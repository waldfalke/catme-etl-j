package com.catmepim.converter.highvolume.core;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import com.catmepim.converter.highvolume.exception.ConversionException;
import com.catmepim.converter.highvolume.exception.ZipBombDetectedException;

/**
 * Provides lazy loading and caching of shared strings from Excel XLSX files.
 * Optimizes memory usage by loading strings on-demand and implementing an LRU cache
 * for the most frequently accessed strings.
 * 
 * <p>Reference: HighVolumeExcelConverter-Contract-v2.0.1.md (§7, §5)</p>
 *
 * @invariant Provides efficient access to shared strings while controlling memory usage.
 * @invariant Maintains a partial cache of most recently used strings.
 */
public class LazySharedStringsProvider {
    
    private static final Logger logger = LoggerFactory.getLogger(LazySharedStringsProvider.class);
    
    private static final int DEFAULT_CACHE_SIZE = 10000; // Default number of strings to cache in memory
    
    private final SafePOIEntryStreamer entryStreamer;
    private final FallbackZipExtractor fallbackExtractor;
    
    private final int cacheSize;
    private final Map<Integer, String> stringCache; // LRU cache for most recently used strings
    private int totalStringsCount = -1; // Total number of strings in the shared strings table
    private boolean parsedSuccessfully = false;
    
    /**
     * Constructs a LazySharedStringsProvider with default cache size.
     *
     * @param entryStreamer the SafePOIEntryStreamer to use for XML access
     * @param fallbackExtractor the FallbackZipExtractor to use if a zip bomb is detected
     * @pre entryStreamer != null and initialized
     * @pre fallbackExtractor != null and initialized
     * @post A new LazySharedStringsProvider is created with default cache size
     */
    public LazySharedStringsProvider(SafePOIEntryStreamer entryStreamer, FallbackZipExtractor fallbackExtractor) {
        this(entryStreamer, fallbackExtractor, DEFAULT_CACHE_SIZE);
    }
    
    /**
     * Constructs a LazySharedStringsProvider with specified cache size.
     *
     * @param entryStreamer the SafePOIEntryStreamer to use for XML access
     * @param fallbackExtractor the FallbackZipExtractor to use if a zip bomb is detected
     * @param cacheSize the maximum number of strings to cache in memory
     * @pre entryStreamer != null and initialized
     * @pre fallbackExtractor != null and initialized
     * @pre cacheSize > 0
     * @post A new LazySharedStringsProvider is created with the specified cache size
     */
    public LazySharedStringsProvider(SafePOIEntryStreamer entryStreamer, FallbackZipExtractor fallbackExtractor, int cacheSize) {
        if (entryStreamer == null) {
            throw new IllegalArgumentException("SafePOIEntryStreamer must not be null");
        }
        if (fallbackExtractor == null) {
            throw new IllegalArgumentException("FallbackZipExtractor must not be null");
        }
        if (cacheSize <= 0) {
            throw new IllegalArgumentException("Cache size must be positive");
        }
        
        this.entryStreamer = entryStreamer;
        this.fallbackExtractor = fallbackExtractor;
        this.cacheSize = cacheSize;
        
        // Create an LRU cache with the specified size
        this.stringCache = new LinkedHashMap<Integer, String>(cacheSize + 1, 0.75f, true) {
            private static final long serialVersionUID = 1L;
            
            @Override
            protected boolean removeEldestEntry(Map.Entry<Integer, String> eldest) {
                return size() > cacheSize;
            }
        };
    }
    
    /**
     * Initializes the shared strings provider by parsing the count/unique count
     * from the shared strings XML file.
     *
     * @throws IOException if an I/O error occurs
     * @pre Not yet initialized or initialization failed previously
     * @post totalStringsCount is set to the total number of strings
     */
    public void initialize() throws IOException {
        if (parsedSuccessfully) {
            return; // Already initialized
        }
        
        logger.info("Initializing LazySharedStringsProvider");
        
        try {
            // Try to get shared strings through POI-safe path first
            try (InputStream sharedStringsStream = entryStreamer.getSharedStringsStream()) {
                parseSharedStringsCounts(sharedStringsStream);
            }
        } catch (ZipBombDetectedException e) {
            logger.warn("POI detected potential zip bomb in sharedStrings.xml, falling back to manual extraction");
            // Fall back to manual extraction
            if (fallbackExtractor.hasEntry("xl/sharedStrings.xml")) {
                try (InputStream fallbackStream = fallbackExtractor.extractEntry("xl/sharedStrings.xml")) {
                    parseSharedStringsCounts(fallbackStream);
                }
            } else {
                throw new IOException("Could not find sharedStrings.xml in fallback mode");
            }
        }
    }
    
    /**
     * Parses the shared strings XML file to get the count of strings.
     *
     * @param inputStream the InputStream for the shared strings XML
     * @throws IOException if an I/O error occurs during parsing
     * @pre inputStream != null and positioned at start of XML
     * @post totalStringsCount is set and parsedSuccessfully is true if parsing succeeds
     */
    private void parseSharedStringsCounts(InputStream inputStream) throws IOException {
        try {
            SAXParserFactory factory = SAXParserFactory.newInstance();
            factory.setNamespaceAware(true);
            factory.setValidating(false);
            
            // Prevent XML attacks
            factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
            factory.setFeature("http://xml.org/sax/features/external-general-entities", false);
            factory.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
            
            SAXParser saxParser = factory.newSAXParser();
            XMLReader xmlReader = saxParser.getXMLReader();
            
            final AtomicInteger count = new AtomicInteger(-1);
            
            DefaultHandler handler = new DefaultHandler() {
                @Override
                public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
                    if ("sst".equals(localName) || "sst".equals(qName)) {
                        String countStr = attributes.getValue("count");
                        if (countStr != null) {
                            count.set(Integer.parseInt(countStr));
                        }
                    }
                }
            };
            
            xmlReader.setContentHandler(handler);
            xmlReader.parse(new InputSource(inputStream));
            
            totalStringsCount = count.get();
            parsedSuccessfully = totalStringsCount >= 0;
            
            logger.info("SharedStrings count: {}", totalStringsCount);
            
        } catch (ParserConfigurationException | SAXException e) {
            throw new IOException("Error parsing shared strings XML", e);
        }
    }
    
    /**
     * Gets a string from the shared strings table by index.
     * The string is loaded from the XML file if not already in the cache.
     *
     * @param index the index of the string to retrieve
     * @return the string at the specified index
     * @throws IOException if an I/O error occurs
     * @throws IndexOutOfBoundsException if the index is negative or >= total count
     * @pre index >= 0 and index < totalStringsCount
     * @post returns the string at the specified index
     */
    public String getString(int index) throws IOException {
        if (!parsedSuccessfully) {
            initialize();
        }
        
        if (index < 0 || (totalStringsCount >= 0 && index >= totalStringsCount)) {
            throw new IndexOutOfBoundsException("String index out of bounds: " + index);
        }
        
        // Check cache first
        String cachedString = stringCache.get(index);
        if (cachedString != null) {
            return cachedString;
        }
        
        // Not in cache, need to parse the XML to find this string
        logger.debug("String at index {} not in cache, loading from XML", index);
        
        List<String> loadedStrings = new ArrayList<>();
        try {
            // Try to get shared strings through POI-safe path first
            try (InputStream sharedStringsStream = entryStreamer.getSharedStringsStream()) {
                loadStringAtIndex(sharedStringsStream, index, loadedStrings);
            }
        } catch (ZipBombDetectedException e) {
            logger.warn("POI detected potential zip bomb in sharedStrings.xml when accessing string {}, falling back to manual extraction", index);
            // Fall back to manual extraction
            if (fallbackExtractor.hasEntry("xl/sharedStrings.xml")) {
                try (InputStream fallbackStream = fallbackExtractor.extractEntry("xl/sharedStrings.xml")) {
                    loadStringAtIndex(fallbackStream, index, loadedStrings);
                }
            } else {
                throw new IOException("Could not find sharedStrings.xml in fallback mode");
            }
        }
        
        if (loadedStrings.isEmpty()) {
            throw new IOException("Failed to load string at index " + index);
        }
        
        // Add to cache
        String result = loadedStrings.get(0);
        stringCache.put(index, result);
        
        return result;
    }
    
    /**
     * Loads a specific string from the shared strings XML by index.
     *
     * @param inputStream the InputStream for the shared strings XML
     * @param targetIndex the index of the string to load
     * @param resultList the list to add the loaded string to
     * @throws IOException if an I/O error occurs during parsing
     * @pre inputStream != null and positioned at start of XML
     * @pre targetIndex >= 0
     * @pre resultList != null
     * @post If string is found, it is added to resultList
     */
    private void loadStringAtIndex(InputStream inputStream, final int targetIndex, final List<String> resultList) throws IOException {
        try {
            SAXParserFactory factory = SAXParserFactory.newInstance();
            factory.setNamespaceAware(true);
            factory.setValidating(false);
            
            // Prevent XML attacks
            factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
            factory.setFeature("http://xml.org/sax/features/external-general-entities", false);
            factory.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
            
            SAXParser saxParser = factory.newSAXParser();
            XMLReader xmlReader = saxParser.getXMLReader();
            
            DefaultHandler handler = new DefaultHandler() {
                private int currentIndex = -1;
                private boolean inSI = false;
                private boolean inT = false;
                private StringBuilder currentText = new StringBuilder();
                
                @Override
                public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
                    if ("si".equals(localName) || "si".equals(qName)) {
                        currentIndex++;
                        inSI = true;
                        currentText.setLength(0);
                    } else if (inSI && ("t".equals(localName) || "t".equals(qName))) {
                        inT = true;
                    }
                }
                
                @Override
                public void endElement(String uri, String localName, String qName) throws SAXException {
                    if ("si".equals(localName) || "si".equals(qName)) {
                        inSI = false;
                        if (currentIndex == targetIndex) {
                            resultList.add(currentText.toString());
                            throw new SAXException("TargetFoundException"); // Used to short-circuit parsing
                        }
                    } else if (inSI && ("t".equals(localName) || "t".equals(qName))) {
                        inT = false;
                    }
                }
                
                @Override
                public void characters(char[] ch, int start, int length) throws SAXException {
                    if (inSI && inT && currentIndex == targetIndex) {
                        currentText.append(ch, start, length);
                    }
                }
            };
            
            xmlReader.setContentHandler(handler);
            try {
                xmlReader.parse(new InputSource(inputStream));
            } catch (SAXException e) {
                if (!"TargetFoundException".equals(e.getMessage())) {
                    throw e; // Re-throw if not our intentional exception
                }
                // Otherwise we found our target string and short-circuited the parse
            }
            
        } catch (ParserConfigurationException | SAXException e) {
            if (!"TargetFoundException".equals(e.getMessage())) {
                throw new IOException("Error parsing shared strings XML", e);
            }
        }
    }
    
    /**
     * Gets the total number of strings in the shared strings table.
     *
     * @return the total number of strings, or -1 if not yet initialized
     * @throws IOException if an I/O error occurs during initialization
     * @pre none
     * @post Returns total string count or initializes if needed
     */
    public int getTotalStringsCount() throws IOException {
        if (!parsedSuccessfully) {
            initialize();
        }
        return totalStringsCount;
    }
    
    /**
     * Gets the current count of cached strings.
     *
     * @return the number of strings currently in the cache
     * @pre none
     * @post Returns current cache size
     */
    public int getCachedStringsCount() {
        return stringCache.size();
    }
    
    /**
     * Clears the string cache.
     *
     * @pre none
     * @post The string cache is empty
     */
    public void clearCache() {
        stringCache.clear();
        logger.debug("SharedStrings cache cleared");
    }
} 