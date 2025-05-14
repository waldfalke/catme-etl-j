package com.catmepim.converter.highvolume.core;

import java.io.BufferedInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Provides fallback ZIP extraction capabilities when POI's OPC package handler
 * detects a potential ZIP bomb. Uses java.util.zip directly with control over buffer sizes
 * and extraction limits.
 * 
 * <p>Reference: HighVolumeExcelConverter-Contract-v2.0.1.md (§7, §9)</p>
 *
 * @invariant Maintains safe extraction with configurable memory limits.
 * @invariant Can extract individual entries without loading entire archive.
 */
public class FallbackZipExtractor implements AutoCloseable {
    
    private static final Logger logger = LoggerFactory.getLogger(FallbackZipExtractor.class);
    private static final int DEFAULT_BUFFER_SIZE = 8192;
    private static final long MAX_ENTRY_SIZE = 100 * 1024 * 1024; // 100 MB max for single entry
    private static final int MAX_INFLATION_FACTOR = 1000; // Max 1000x inflation allowed
    
    private final Path zipFilePath;
    private final Map<String, Long> entrySizes = new HashMap<>();
    
    /**
     * Constructs a FallbackZipExtractor for the given ZIP file.
     *
     * @param zipFilePath Path to the ZIP file
     * @pre zipFilePath != null and represents an existing, readable file
     * @post A new FallbackZipExtractor is created and the ZIP entries are scanned
     * @throws IOException if an I/O error occurs
     */
    public FallbackZipExtractor(Path zipFilePath) throws IOException {
        if (zipFilePath == null) {
            throw new IllegalArgumentException("ZIP file path must not be null");
        }
        this.zipFilePath = zipFilePath;
        scanZipEntries();
    }
    
    /**
     * Scans all entries in the ZIP file to create a map of entry names to compressed sizes.
     * This allows for quick lookups and size validations without re-parsing the central directory.
     *
     * @throws IOException if an I/O error occurs
     * @pre zipFilePath points to a valid ZIP file
     * @post entrySizes map is populated with all ZIP entry names and their compressed sizes
     */
    private void scanZipEntries() throws IOException {
        logger.info("Scanning ZIP entries in file: {}", zipFilePath);
        entrySizes.clear();
        
        try (ZipInputStream zis = new ZipInputStream(
                new BufferedInputStream(new FileInputStream(zipFilePath.toFile())))) {
            
            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                entrySizes.put(entry.getName(), entry.getCompressedSize());
                logger.debug("Found ZIP entry: {} (compressed size: {} bytes)", 
                        entry.getName(), entry.getCompressedSize());
                zis.closeEntry();
            }
        }
        
        logger.info("ZIP scan complete. Found {} entries", entrySizes.size());
    }
    
    /**
     * Extracts a specific entry from the ZIP file based on its name.
     * Uses a controlled buffer size and sets limits on the maximum extracted size
     * to prevent zip bombs from consuming excessive memory.
     *
     * @param entryName Name of the entry to extract
     * @return InputStream containing the extracted entry data
     * @throws IOException if an I/O error occurs or safety limits are exceeded
     * @pre entryName != null and represents an entry in the ZIP file
     * @post Returns an InputStream for the requested entry if safety limits are met
     */
    public InputStream extractEntry(String entryName) throws IOException {
        if (entryName == null) {
            throw new IllegalArgumentException("Entry name must not be null");
        }
        
        logger.info("Extracting entry '{}' from ZIP file", entryName);
        
        try (ZipInputStream zis = new ZipInputStream(
                new BufferedInputStream(new FileInputStream(zipFilePath.toFile()), DEFAULT_BUFFER_SIZE))) {
            
            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                if (entry.getName().equals(entryName)) {
                    logger.debug("Found entry '{}', extracting with safety controls", entryName);
                    return extractEntryWithSafety(zis, entry);
                }
                zis.closeEntry();
            }
        }
        
        throw new IOException("Entry not found: " + entryName);
    }
    
    /**
     * Safely extracts a ZIP entry with memory and expansion controls.
     * 
     * @param zis ZipInputStream positioned at the start of the entry content
     * @param entry ZipEntry to extract
     * @return InputStream for the extracted content
     * @throws IOException if safety limits are exceeded or I/O error occurs
     * @pre zis is positioned at the start of entry's data
     * @post Returns InputStream for the safely extracted data
     */
    private InputStream extractEntryWithSafety(ZipInputStream zis, ZipEntry entry) throws IOException {
        // For safety with potentially malicious files, we'll read the entry fully 
        // with controlled buffers and then return a ByteArrayInputStream
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        byte[] buffer = new byte[DEFAULT_BUFFER_SIZE];
        
        long totalBytesRead = 0;
        long compressedSize = entry.getCompressedSize();
        if (compressedSize <= 0) {
            // Size unknown, use the scan result or default
            compressedSize = entrySizes.getOrDefault(entry.getName(), 0L);
            if (compressedSize <= 0) {
                // Still unknown, assume small for now
                compressedSize = 10 * 1024; // 10 KB guess
            }
        }
        
        int bytesRead;
        while ((bytesRead = zis.read(buffer)) != -1) {
            totalBytesRead += bytesRead;
            
            // Safety check: enforce max size limit
            if (totalBytesRead > MAX_ENTRY_SIZE) {
                throw new IOException("Entry '" + entry.getName() + "' exceeds maximum allowed size of " + 
                        MAX_ENTRY_SIZE + " bytes");
            }
            
            // Safety check: enforce max inflation factor
            if (compressedSize > 0 && totalBytesRead > compressedSize * MAX_INFLATION_FACTOR) {
                throw new IOException("Entry '" + entry.getName() + "' inflation ratio exceeds safety limit of " + 
                        MAX_INFLATION_FACTOR + ":1");
            }
            
            baos.write(buffer, 0, bytesRead);
        }
        
        logger.info("Extracted '{}': {} bytes (inflation ratio: {}:1)", 
                entry.getName(), totalBytesRead, 
                compressedSize > 0 ? String.format("%.2f", ((double) totalBytesRead / compressedSize)) : "unknown");
        
        return new java.io.ByteArrayInputStream(baos.toByteArray());
    }
    
    /**
     * Checks if an entry exists in the ZIP file.
     *
     * @param entryName Name of the entry to check
     * @return true if the entry exists, false otherwise
     * @pre entryName != null
     * @post Returns whether the specified entry exists in the ZIP file
     */
    public boolean hasEntry(String entryName) {
        if (entryName == null) {
            throw new IllegalArgumentException("Entry name must not be null");
        }
        return entrySizes.containsKey(entryName);
    }
    
    /**
     * Gets a list of all entry names in the ZIP file.
     *
     * @return Array of entry names
     * @pre zip file has been scanned successfully
     * @post Returns an array of all entry names in the ZIP
     */
    public String[] getEntryNames() {
        return entrySizes.keySet().toArray(new String[0]);
    }
    
    @Override
    public void close() {
        // No resources to close in this implementation
        // The ZipInputStream instances are all closed in their respective methods
    }
} 