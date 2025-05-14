package com.catmepim.converter.highvolume.core;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.catmepim.converter.highvolume.exception.ConversionException;
import com.catmepim.converter.highvolume.exception.ZipBombDetectedException;

/**
 * Safe wrapper for streaming XML entries from an OPC (XLSX) package.
 * Provides safe access to shared strings, sheets, and other XML components.
 * 
 * <p>Reference: HighVolumeExcelConverter-Contract-v2.0.1.md (§7, §9)</p>
 *
 * @invariant Provides safe streaming access to OPC package parts.
 * @invariant Detects potential zip bombs and allows fallback handling.
 */
public class SafePOIEntryStreamer {
    
    private static final Logger logger = LoggerFactory.getLogger(SafePOIEntryStreamer.class);
    private final OPCPackage opcPackage;
    
    /**
     * Constructs a SafePOIEntryStreamer for the given OPC package.
     *
     * @param opcPackage The OPC package to stream from
     * @pre opcPackage != null and is opened
     * @post A SafePOIEntryStreamer instance is created for the given package
     */
    public SafePOIEntryStreamer(OPCPackage opcPackage) {
        if (opcPackage == null) {
            throw new IllegalArgumentException("OPCPackage must not be null");
        }
        this.opcPackage = opcPackage;
    }
    
    /**
     * Gets the shared strings XML part as a stream.
     * 
     * @return InputStream for the shared strings XML
     * @throws IOException if I/O error occurs
     * @throws ZipBombDetectedException if a potential zip bomb is detected
     * @pre opcPackage is initialized and valid
     * @post Returns input stream for sharedStrings.xml if found, otherwise throws exception
     */
    public InputStream getSharedStringsStream() throws IOException, ZipBombDetectedException {
        try {
            PackagePart sharedStringsPart = null;
            
            // Find the shared strings part
            for (PackagePart part : opcPackage.getParts()) {
                if (part.getPartName().getName().endsWith("sharedStrings.xml")) {
                    sharedStringsPart = part;
                    break;
                }
            }
            
            if (sharedStringsPart == null) {
                logger.warn("No sharedStrings.xml found in XLSX file");
                throw new ConversionException("No sharedStrings.xml found in XLSX file");
            }
            
            logger.debug("Found sharedStrings.xml part: {}", sharedStringsPart.getPartName());
            try {
                return sharedStringsPart.getInputStream();
            } catch (Exception e) {
                // Check if it's a potential zip bomb
                if (e.getMessage() != null && 
                    (e.getMessage().contains("Zip bomb detected") || 
                     e.getMessage().contains("Unexpected decompression ratio"))) {
                    logger.warn("Potential zip bomb detected in sharedStrings.xml");
                    throw new ZipBombDetectedException("Potential zip bomb detected in sharedStrings.xml", e);
                }
                throw e;
            }
        } catch (ZipBombDetectedException e) {
            throw e; // Re-throw to be handled by caller
        } catch (Exception e) {
            logger.error("Error getting sharedStrings.xml: {}", e.getMessage(), e);
            throw new IOException("Error accessing sharedStrings.xml", e);
        }
    }
    
    /**
     * Gets the sheet XML part with the given name as a stream.
     * 
     * @param sheetName Name of the sheet to get
     * @return InputStream for the sheet XML
     * @throws IOException if I/O error occurs
     * @throws ZipBombDetectedException if a potential zip bomb is detected
     * @pre opcPackage is initialized and valid, sheetName is not null
     * @post Returns input stream for the requested sheet if found, otherwise throws exception
     */
    public InputStream getSheetStream(String sheetName) throws IOException, ZipBombDetectedException {
        if (sheetName == null) {
            throw new IllegalArgumentException("Sheet name must not be null");
        }
        
        try {
            // First, find the workbook
            PackagePart workbookPart = null;
            for (PackagePart part : opcPackage.getParts()) {
                if (part.getPartName().getName().endsWith("workbook.xml")) {
                    workbookPart = part;
                    break;
                }
            }
            
            if (workbookPart == null) {
                throw new ConversionException("No workbook.xml found in XLSX file");
            }
            
            // Get relationships from workbook to sheets
            PackageRelationshipCollection relationships = workbookPart.getRelationships();
            
            // Find the sheet with the given name
            PackagePart targetSheet = null;
            String targetSheetRId = null;
            
            // This is simplistic - in a real implementation, you'd parse workbook.xml
            // to match sheet name to relationship ID
            // For now, just find any sheet relationship and use that (first sheet)
            for (PackageRelationship rel : relationships) {
                if (rel.getRelationshipType().endsWith("worksheet")) {
                    targetSheetRId = rel.getId();
                    break;
                }
            }
            
            if (targetSheetRId != null) {
                targetSheet = opcPackage.getPart(workbookPart.getRelationship(targetSheetRId));
            }
            
            if (targetSheet == null) {
                throw new ConversionException("Sheet not found: " + sheetName);
            }
            
            logger.debug("Found sheet part: {}", targetSheet.getPartName());
            try {
                return targetSheet.getInputStream();
            } catch (Exception e) {
                // Check if it's a potential zip bomb
                if (e.getMessage() != null && 
                    (e.getMessage().contains("Zip bomb detected") || 
                     e.getMessage().contains("Unexpected decompression ratio"))) {
                    logger.warn("Potential zip bomb detected in sheet: {}", sheetName);
                    throw new ZipBombDetectedException("Potential zip bomb detected in sheet: " + sheetName, e);
                }
                throw e;
            }
        } catch (ZipBombDetectedException e) {
            throw e; // Re-throw to be handled by caller
        } catch (Exception e) {
            logger.error("Error getting sheet {}: {}", sheetName, e.getMessage(), e);
            throw new IOException("Error accessing sheet: " + sheetName, e);
        }
    }
    
    /**
     * Gets first sheet in the workbook.
     * 
     * @return InputStream for the first sheet
     * @throws IOException if I/O error occurs
     * @throws ZipBombDetectedException if a potential zip bomb is detected
     * @pre opcPackage is initialized and valid
     * @post Returns input stream for the first sheet if found, otherwise throws exception
     */
    public InputStream getFirstSheetStream() throws IOException, ZipBombDetectedException {
        try {
            // Find the workbook
            PackagePart workbookPart = null;
            for (PackagePart part : opcPackage.getParts()) {
                if (part.getPartName().getName().endsWith("workbook.xml")) {
                    workbookPart = part;
                    break;
                }
            }
            
            if (workbookPart == null) {
                throw new ConversionException("No workbook.xml found in XLSX file");
            }
            
            // Get relationships from workbook to sheets
            PackageRelationshipCollection relationships = workbookPart.getRelationships();
            
            // Find the first sheet relationship
            PackagePart firstSheet = null;
            
            for (PackageRelationship rel : relationships) {
                if (rel.getRelationshipType().endsWith("worksheet")) {
                    PackagePart sheetPart = opcPackage.getPart(workbookPart.getRelationship(rel.getId()));
                    firstSheet = sheetPart;
                    break;
                }
            }
            
            if (firstSheet == null) {
                throw new ConversionException("No sheets found in workbook");
            }
            
            logger.debug("Found first sheet part: {}", firstSheet.getPartName());
            try {
                return firstSheet.getInputStream();
            } catch (Exception e) {
                // Check if it's a potential zip bomb
                if (e.getMessage() != null && 
                    (e.getMessage().contains("Zip bomb detected") || 
                     e.getMessage().contains("Unexpected decompression ratio"))) {
                    logger.warn("Potential zip bomb detected in first sheet");
                    throw new ZipBombDetectedException("Potential zip bomb detected in first sheet", e);
                }
                throw e;
            }
        } catch (ZipBombDetectedException e) {
            throw e; // Re-throw to be handled by caller
        } catch (Exception e) {
            logger.error("Error getting first sheet: {}", e.getMessage(), e);
            throw new IOException("Error accessing first sheet", e);
        }
    }
    
    /**
     * Close the underlying OPC package.
     * 
     * @throws IOException if I/O error occurs
     * @pre opcPackage is initialized and valid
     * @post opcPackage is closed and resources are released
     */
    public void close() throws IOException {
        try {
            if (opcPackage != null) {
                opcPackage.close();
            }
        } catch (IOException e) {
            logger.error("Error closing OPC package: {}", e.getMessage(), e);
            throw e;
        }
    }
} 