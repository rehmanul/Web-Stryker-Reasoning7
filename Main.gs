/**
 * Web Stryker R7 - Advanced Company & Product Data Extraction System
 * Version: 3.1.5
 */

// Global configuration
const CONFIG = {
  MAX_RETRIES: 3,
  TIMEOUT_MS: 30000,
  USER_AGENT: "Mozilla/5.0 (compatible; GoogleAppsScript/1.0)",
  ENABLE_ADVANCED_FEATURES: true,
  FALLBACK_TO_BASIC: true,
  MAX_PRODUCT_PAGES: 10,
  MAX_CRAWL_DEPTH: 3,
  
  // API configuration
  API: {
    AZURE: {
      OPENAI: {
        ENDPOINT: "https://fetcher.openai.azure.com/",
        KEY: "",
        DEPLOYMENT: "gpt-4",
        MAX_TOKENS: 1000,
        TEMPERATURE: 0.3
      }
    },
    KNOWLEDGE_GRAPH: {
      KEY: ""
    }
  },
  
  // Extraction settings
  EXTRACTION: {
    FOLLOW_LINKS: true,
    MAX_PRODUCTS: 20,
    EXTRACT_IMAGES: true,
    DETAILED_LOGGING: true
  }
};

// Object to track extraction states
const extractionStates = {};

// Global stats for reporting
let globalStats = {
  processed: 0,
  remaining: 0,
  success: 0,
  fail: 0,
  apiCalls: {
    azure: { success: 0, fail: 0 },
    knowledgeGraph: { success: 0, fail: 0 }
  },
  companyData: {
    found: 0,
    emails: 0,
    phones: 0,
    addresses: 0,
    descriptions: 0,
    types: 0
  },
  productData: {
    found: 0,
    images: 0,
    descriptions: 0,
    categories: 0
  }
};

/**
 * Required doGet function for web app deployment
 * @param {Object} e - The event parameter
 * @return {HtmlOutput} The HTML page to be displayed
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('WebAppInterface')
    .evaluate()
    .setTitle('Web Stryker R7')
    .setFaviconUrl('https://www.google.com/images/icons/product/sheets-32.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Include CSS and JavaScript files in the web app
 * @param {string} filename - The name of the file to include
 * @return {string} The contents of the file
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Setup spreadsheet structure with all required sheets
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create Data sheet if it doesn't exist
  let dataSheet = ss.getSheetByName("Data");
  if (!dataSheet) {
    dataSheet = ss.insertSheet("Data");
    
    // Setup data headers based on the required schema
    const headers = [
      "URL", "Status", "Company Name", "Addresses", "Emails", "Phones", 
      "Company Description", "Company Type", "Product Name", "Product URL", 
      "Product Category", "Product Subcategory", "Product Family", 
      "Quantity", "Price", "Product Description", "Specifications", 
      "Images", "Extraction Date"
    ];
    
    const headerRange = dataSheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight("bold");
    
    // Format data sheet
    dataSheet.setFrozenRows(1);
    dataSheet.setColumnWidths(1, headers.length, 200);
    
    // Apply formatting
    headerRange.setBackground("#E0E0E0");
    
    // Create filter
    dataSheet.getRange(1, 1, 1, headers.length).createFilter();
  }
  
  // Create Extraction Log sheet if it doesn't exist
  let extractionLogSheet = ss.getSheetByName("Extraction Log");
  if (!extractionLogSheet) {
    extractionLogSheet = ss.insertSheet("Extraction Log");
    
    // Setup extraction log headers
    const logHeaders = [
      "Timestamp", "URL", "Extraction ID", "Operation", "Status", "Details", "Duration (ms)"
    ];
    
    const logHeaderRange = extractionLogSheet.getRange(1, 1, 1, logHeaders.length);
    logHeaderRange.setValues([logHeaders]);
    logHeaderRange.setFontWeight("bold");
    logHeaderRange.setBackground("#E0E0E0");
    
    // Format extraction log sheet
    extractionLogSheet.setFrozenRows(1);
    extractionLogSheet.setColumnWidth(1, 180); // Timestamp
    extractionLogSheet.setColumnWidth(2, 250); // URL
    extractionLogSheet.setColumnWidth(3, 200); // Extraction ID
    extractionLogSheet.setColumnWidth(4, 150); // Operation
    extractionLogSheet.setColumnWidth(5, 120); // Status
    extractionLogSheet.setColumnWidth(6, 350); // Details
    extractionLogSheet.setColumnWidth(7, 120); // Duration
    
    // Create filter
    extractionLogSheet.getRange(1, 1, 1, logHeaders.length).createFilter();
  }
  
  // Create Error Log sheet if it doesn't exist
  let errorLogSheet = ss.getSheetByName("Error Log");
  if (!errorLogSheet) {
    errorLogSheet = ss.insertSheet("Error Log");
    
    // Setup error log headers
    const errorHeaders = [
      "Timestamp", "URL", "Extraction ID", "Error Type", "Error Message", "Stack Trace"
    ];
    
    const errorHeaderRange = errorLogSheet.getRange(1, 1, 1, errorHeaders.length);
    errorHeaderRange.setValues([errorHeaders]);
    errorHeaderRange.setFontWeight("bold");
    errorHeaderRange.setBackground("#FFCDD2"); // Light red for error log
    
    // Format error log sheet
    errorLogSheet.setFrozenRows(1);
    errorLogSheet.setColumnWidth(1, 180); // Timestamp
    errorLogSheet.setColumnWidth(2, 250); // URL
    errorLogSheet.setColumnWidth(3, 200); // Extraction ID
    errorLogSheet.setColumnWidth(4, 150); // Error Type
    errorLogSheet.setColumnWidth(5, 350); // Error Message
    errorLogSheet.setColumnWidth(6, 400); // Stack Trace
    
    // Create filter
    errorLogSheet.getRange(1, 1, 1, errorHeaders.length).createFilter();
  }
  
  // Create Stats Dashboard sheet if it doesn't exist
  let statsSheet = ss.getSheetByName("Data Center");
  if (!statsSheet) {
    statsSheet = ss.insertSheet("Data Center");
    
    // Set up dashboard structure
    statsSheet.getRange("A1:R1").merge();
    statsSheet.getRange("A1").setValue("Web Stryker R7 - Extraction Overview");
    statsSheet.getRange("A1").setFontWeight("bold");
    statsSheet.getRange("A1").setBackground("#3a1c71");
    statsSheet.getRange("A1").setFontColor("white");
    statsSheet.getRange("A1").setFontSize(16);
    statsSheet.getRange("A1").setHorizontalAlignment("center");
    
    // Main Task Report section
    statsSheet.getRange("A3:R3").merge();
    statsSheet.getRange("A3").setValue("Main Task Report");
    statsSheet.getRange("A3").setFontWeight("bold");
    statsSheet.getRange("A3").setBackground("#4285f4");
    statsSheet.getRange("A3").setFontColor("white");
    statsSheet.getRange("A3").setHorizontalAlignment("center");
    
    const mainReportHeaders = [
      "URL", "Processed", "Remaining", "Success", "Fail", 
      "Company Name Found", "Emails Found", "Phones Found", "Addresses Found", 
      "Company logo URL Found", "Categories Found", "Product Names Found", 
      "Product Image URLs Found", "Product Type Found"
    ];
    
    statsSheet.getRange(4, 1, 1, mainReportHeaders.length).setValues([mainReportHeaders]);
    statsSheet.getRange(4, 1, 1, mainReportHeaders.length).setFontWeight("bold");
    statsSheet.getRange(4, 1, 1, mainReportHeaders.length).setBackground("#e0e0e0");
    
    // API Call Report section
    statsSheet.getRange("A7:K7").merge();
    statsSheet.getRange("A7").setValue("API Call Report");
    statsSheet.getRange("A7").setFontWeight("bold");
    statsSheet.getRange("A7").setBackground("#4285f4");
    statsSheet.getRange("A7").setFontColor("white");
    statsSheet.getRange("A7").setHorizontalAlignment("center");
    
    const apiReportHeaders = [
      "URL", "Processed", "Remaining", "Success", "Fail",
      "Azure OpenAI Success", "Azure OpenAI Fail",
      "Google Knowledge Graph API Key Success", "Google Knowledge Graph API Key Fail",
      "Remarks On API Failure"
    ];
    
    statsSheet.getRange(8, 1, 1, apiReportHeaders.length).setValues([apiReportHeaders]);
    statsSheet.getRange(8, 1, 1, apiReportHeaders.length).setFontWeight("bold");
    statsSheet.getRange(8, 1, 1, apiReportHeaders.length).setBackground("#e0e0e0");
    
    // Product Categories Report section
    statsSheet.getRange("A11:G11").merge();
    statsSheet.getRange("A11").setValue("Product Categories Report");
    statsSheet.getRange("A11").setFontWeight("bold");
    statsSheet.getRange("A11").setBackground("#d76d77");
    statsSheet.getRange("A11").setFontColor("white");
    statsSheet.getRange("A11").setHorizontalAlignment("center");
    
    const categoriesReportHeaders = [
      "Company", "Categories Found", "Single Categories Found For", "Multi Categories Found For",
      "Categories Unpredictable For", "Categories Not Found", "Remarks On Categories"
    ];
    
    statsSheet.getRange(12, 1, 1, categoriesReportHeaders.length).setValues([categoriesReportHeaders]);
    statsSheet.getRange(12, 1, 1, categoriesReportHeaders.length).setFontWeight("bold");
    statsSheet.getRange(12, 1, 1, categoriesReportHeaders.length).setBackground("#e0e0e0");
    
    // Product Detailed Titles Report section
    statsSheet.getRange("I11:R11").merge();
    statsSheet.getRange("I11").setValue("Product Detailed Titles Report");
    statsSheet.getRange("I11").setFontWeight("bold");
    statsSheet.getRange("I11").setBackground("#d76d77");
    statsSheet.getRange("I11").setFontColor("white");
    statsSheet.getRange("I11").setHorizontalAlignment("center");
    
    const detailsReportHeaders = [
      "Company", "Detailed Product Title Fetched For", "Detail Title Fetch Failed For",
      "Image Link Collects For", "Image Link Collect Failed For", "Total Failed",
      "Remarks On Categories", "Azure OpenAI Fail"
    ];
    
    statsSheet.getRange(12, 9, 1, detailsReportHeaders.length).setValues([detailsReportHeaders]);
    statsSheet.getRange(12, 9, 1, detailsReportHeaders.length).setFontWeight("bold");
    statsSheet.getRange(12, 9, 1, detailsReportHeaders.length).setBackground("#e0e0e0");
    
    // Initialize stats row with zeros
    const statsRow = Array(mainReportHeaders.length).fill(0);
    statsSheet.getRange(5, 1, 1, statsRow.length).setValues([statsRow]);
    
    const apiStatsRow = Array(apiReportHeaders.length).fill(0);
    statsSheet.getRange(9, 1, 1, apiStatsRow.length).setValues([apiStatsRow]);
    
    const categoryStatsRow = Array(categoriesReportHeaders.length).fill(0);
    statsSheet.getRange(13, 1, 1, categoryStatsRow.length).setValues([categoryStatsRow]);
    
    const detailsStatsRow = Array(detailsReportHeaders.length).fill(0);
    statsSheet.getRange(13, 9, 1, detailsStatsRow.length).setValues([detailsStatsRow]);
  }
}

/**
 * Log extraction operation to the Extraction Log sheet
 * @param {string} url - The URL being processed
 * @param {string} extractionId - The extraction ID
 * @param {string} operation - The operation being performed
 * @param {string} status - The status of the operation
 * @param {string} details - Additional details about the operation
 * @param {number} duration - Duration of the operation in milliseconds (optional)
 */
function logExtractionOperation(url, extractionId, operation, status, details, duration = null) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName("Extraction Log");
    
    if (!logSheet) {
      console.log("Extraction Log sheet not found. Creating it...");
      setupSpreadsheet();
      return logExtractionOperation(url, extractionId, operation, status, details, duration);
    }
    
    const timestamp = new Date().toISOString();
    const logData = [timestamp, url, extractionId, operation, status, details, duration];
    
    // Append to the log sheet
    logSheet.appendRow(logData);
    
    // Console log for debugging
    console.log(`[${operation}] ${status}: ${details}`);
    
  } catch (error) {
    console.error(`Error logging extraction operation: ${error.message}`);
    logError(url, extractionId, "LoggingError", error.message, error.stack);
  }
}

/**
 * Log error to the Error Log sheet
 * @param {string} url - The URL being processed
 * @param {string} extractionId - The extraction ID
 * @param {string} errorType - The type of error
 * @param {string} errorMessage - The error message
 * @param {string} stackTrace - The error stack trace (optional)
 */
function logError(url, extractionId, errorType, errorMessage, stackTrace = null) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const errorSheet = ss.getSheetByName("Error Log");
    
    if (!errorSheet) {
      console.log("Error Log sheet not found. Creating it...");
      setupSpreadsheet();
      return logError(url, extractionId, errorType, errorMessage, stackTrace);
    }
    
    const timestamp = new Date().toISOString();
    const errorData = [timestamp, url, extractionId, errorType, errorMessage, stackTrace];
    
    // Append to the error log sheet
    errorSheet.appendRow(errorData);
    
    // Console log for debugging
    console.error(`[ERROR] ${errorType}: ${errorMessage}`);
    
    // Update global stats
    globalStats.fail++;
    
  } catch (error) {
    // If we can't log to the error sheet, at least log to console
    console.error(`Failed to log error: ${error.message}`);
    console.error(`Original error: ${errorMessage}`);
  }
}

/**
 * Measure execution time of a function and log the result
 * @param {string} url - The URL being processed
 * @param {string} extractionId - The extraction ID
 * @param {string} operation - The operation being performed
 * @param {Function} func - The function to execute
 * @param {Array} args - Arguments to pass to the function
 * @return {*} The result of the function execution
 */
function timeAndLogOperation(url, extractionId, operation, func, ...args) {
  const startTime = new Date().getTime();
  let result;
  let status = "Success";
  let details = "Operation completed successfully";
  
  try {
    // Execute the function
    result = func.apply(null, args);
    return result;
  } catch (error) {
    // Log the error
    status = "Error";
    details = error.message;
    logError(url, extractionId, "OperationError", error.message, error.stack);
    throw error;
  } finally {
    // Calculate duration and log the operation
    const endTime = new Date().getTime();
    const duration = endTime - startTime;
    logExtractionOperation(url, extractionId, operation, status, details, duration);
  }
}

/**
 * Update extraction progress
 * @param {string} extractionId - The extraction ID
 * @param {number} progress - Progress percentage (0-100)
 * @param {string} stage - Current extraction stage
 */
function updateExtractionProgress(extractionId, progress, stage) {
  if (!extractionId || !extractionStates[extractionId]) return;
  
  extractionStates[extractionId].progress = progress;
  extractionStates[extractionId].stage = stage;
}

/**
 * Get current extraction progress
 * @param {string} extractionId - The extraction ID
 * @return {Object} Current progress information
 */
function getExtractionProgress(extractionId) {
  if (!extractionId || !extractionStates[extractionId]) {
    return {
      found: false,
      progress: 0,
      stage: "Not found"
    };
  }
  
  return {
    found: true,
    progress: extractionStates[extractionId].progress,
    stage: extractionStates[extractionId].stage,
    paused: extractionStates[extractionId].paused,
    stopped: extractionStates[extractionId].stopped
  };
}

/**
 * Process a URL submitted through the web interface
 * @param {string} url - The URL to extract data from
 * @param {string} extractionId - Unique ID for this extraction process
 * @return {Object} The extraction results
 */
function processUrlFromWebApp(url, extractionId) {
  try {
    // Start timing the overall extraction
    const startTime = new Date().getTime();
    
    // Initialize extraction state
    if (extractionId) {
      extractionStates[extractionId] = {
        paused: false,
        stopped: false,
        url: url,
        startTime: new Date().toISOString(),
        progress: 0,
        stage: "Initializing"
      };
    }
    
    // Setup spreadsheet if needed
    setupSpreadsheet();
    
    // Get the active spreadsheet and sheets
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName("Data");
    
    // Log extraction start
    logExtractionOperation(url, extractionId, "Extraction", "Started", "Beginning extraction process");
    
    // Update progress
    updateExtractionProgress(extractionId, 5, "Validating URL");
    
    // Validate the URL
    if (!isValidUrl(url)) {
      logError(url, extractionId, "ValidationError", "Invalid URL format");
      return {
        success: false,
        error: "Invalid URL format"
      };
    }
    
    // Initialize URL entry with "In Progress" status
    initializeUrlEntry(url);
    updateExtractionStatus(dataSheet, url, "In Progress");
    
    // Update progress
    updateExtractionProgress(extractionId, 10, "Starting extraction");
    
    // Extract data using the integrated extraction system
    let extractedData;
    try {
      extractedData = timeAndLogOperation(
        url,
        extractionId,
        "DataExtraction",
        extractCompanyAndProductData,
        url,
        CONFIG,
        extractionId
      );
    } catch (error) {
      logError(url, extractionId, "ExtractionError", error.message, error.stack);
      updateExtractionStatus(dataSheet, url, "Failed");
      
      const endTime = new Date().getTime();
      logExtractionOperation(
        url, 
        extractionId, 
        "Extraction", 
        "Failed", 
        `Extraction failed: ${error.message}`,
        endTime - startTime
      );
      
      return {
        success: false,
        error: `Extraction failed: ${error.message}`
      };
    }
    
    if (!extractedData) {
      updateExtractionStatus(dataSheet, url, "Failed");
      logExtractionOperation(url, extractionId, "Extraction", "Failed", "No data extracted");
      
      return {
        success: false,
        error: "Failed to extract data from URL"
      };
    }
    
    // Process the extracted data
    try {
      const processedData = timeAndLogOperation(
        url,
        extractionId,
        "DataProcessing",
        processExtractedData,
        extractedData,
        url,
        CONFIG
      );
      
      // Update with processed data
      extractedData = processedData || extractedData;
    } catch (error) {
      logError(url, extractionId, "ProcessingError", error.message, error.stack);
      // Continue despite processing error
    }
    
    // Store data in the spreadsheet
    try {
      timeAndLogOperation(
        url,
        extractionId,
        "DataStorage",
        storeExtractedData,
        dataSheet,
        extractedData
      );
    } catch (error) {
      logError(url, extractionId, "StorageError", error.message, error.stack);
      // Continue despite storage error
    }
    
    // Update status
    updateExtractionStatus(dataSheet, url, "Completed");
    
    // Update progress
    updateExtractionProgress(extractionId, 100, "Completed");
    
    // Update stats
    updateStatsForExtraction(url, extractedData);
    
    // Calculate total duration
    const endTime = new Date().getTime();
    const totalDuration = endTime - startTime;
    
    // Log completion
    logExtractionOperation(
      url, 
      extractionId, 
      "Extraction", 
      "Completed", 
      `Extraction completed successfully`,
      totalDuration
    );
    
    // Clean up extraction state after a delay
    Utilities.sleep(1000);
    delete extractionStates[extractionId];
    
    // Increment global stats
    globalStats.processed++;
    globalStats.success++;
    
    // Return results for display
    return {
      success: true,
      data: extractedData
    };
    
  } catch (error) {
    const errorMessage = `Error processing URL: ${error.message}`;
    logError(url, extractionId, "ProcessingError", errorMessage, error.stack);
    
    // Update status to Failed
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const dataSheet = ss.getSheetByName("Data");
      updateExtractionStatus(dataSheet, url, "Failed");
    } catch (e) {
      // Ignore error
    }
    
    // Clean up extraction state
    if (extractionId) {
      delete extractionStates[extractionId];
    }
    
    return {
      success: false,
      error: errorMessage
    };
  }
}

/**
 * Process multiple URLs in batch mode
 * @param {Array<string>} urls - Array of URLs to process
 * @return {Object} Batch processing results
 */
function processBatchUrls(urls) {
  // Validate input
  if (!urls || !Array.isArray(urls) || urls.length === 0) {
    return {
      success: false,
      error: "No valid URLs provided"
    };
  }
  
  // Initialize batch state
  const batchId = "batch-" + Date.now().toString();
  const results = {
    success: true,
    total: urls.length,
    processed: 0,
    successful: 0,
    failed: 0,
    failures: []
  };
  
  // Update global stats
  globalStats.remaining = urls.length;
  
  // Process each URL
  urls.forEach((url, index) => {
    try {
      // Generate extraction ID for this URL
      const extractionId = `${batchId}-${index}`;
      
      // Process the URL
      const result = processUrlFromWebApp(url, extractionId);
      
      // Update batch results
      results.processed++;
      if (result.success) {
        results.successful++;
      } else {
        results.failed++;
        results.failures.push({
          url: url,
          error: result.error
        });
      }
      
      // Update global stats
      globalStats.remaining--;
      
    } catch (error) {
      // Log error and continue with next URL
      logError(url, batchId, "BatchProcessingError", error.message, error.stack);
      
      results.processed++;
      results.failed++;
      results.failures.push({
        url: url,
        error: error.message
      });
      
      // Update global stats
      globalStats.remaining--;
    }
    
    // Small delay to prevent quota issues
    Utilities.sleep(1000);
  });
  
  // Update success flag if any failures
  if (results.failed > 0) {
    results.success = false;
  }
  
  return results;
}

/**
 * Process all URLs in the Data sheet that haven't been processed yet
 * @return {Object} Batch processing results
 */
function processAllPendingUrls() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName("Data");
    
    if (!dataSheet) {
      return {
        success: false,
        error: "Data sheet not found"
      };
    }
    
    // Get all rows from the Data sheet
    const dataRange = dataSheet.getDataRange();
    const dataValues = dataRange.getValues();
    
    if (dataValues.length <= 1) {
      return {
        success: false,
        error: "No URLs found in the Data sheet"
      };
    }
    
    // Find URLs with "Pending" status
    const pendingUrls = [];
    
    for (let i = 1; i < dataValues.length; i++) {
      const url = dataValues[i][0];
      const status = dataValues[i][1];
      
      if (url && (status === "Pending" || status === "" || status === undefined)) {
        pendingUrls.push(url);
      }
    }
    
    if (pendingUrls.length === 0) {
      return {
        success: false,
        error: "No pending URLs found"
      };
    }
    
    // Process all pending URLs
    return processBatchUrls(pendingUrls);
    
  } catch (error) {
    return {
      success: false,
      error: `Error processing pending URLs: ${error.message}`
    };
  }
}

/**
 * Initialize entry in the data sheet for a URL with "Pending" status
 * @param {string} url - The URL to initialize
 */
function initializeUrlEntry(url) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName("Data");
    
    // Check if URL already exists in the data sheet
    const dataRange = dataSheet.getDataRange();
    const dataValues = dataRange.getValues();
    
    for (let i = 1; i < dataValues.length; i++) {
      if (dataValues[i][0] === url) {
        // URL already exists, update status to "Pending"
        dataSheet.getRange(i + 1, 2).setValue("Pending");
        return;
      }
    }
    
    // URL doesn't exist, add new row with just the URL and status
    const newRow = [url, "Pending"];
    dataSheet.appendRow(newRow);
    
  } catch (error) {
    console.error(`Error initializing URL entry: ${error.message}`);
  }
}

/**
 * Update extraction status for a URL in the data sheet
 * @param {Sheet} dataSheet - The Data sheet
 * @param {string} url - The URL to update
 * @param {string} status - The new status
 */
function updateExtractionStatus(dataSheet, url, status) {
  try {
    const dataRange = dataSheet.getDataRange();
    const dataValues = dataRange.getValues();
    
    for (let i = 1; i < dataValues.length; i++) {
      if (dataValues[i][0] === url) {
        dataSheet.getRange(i + 1, 2).setValue(status);
        return;
      }
    }
  } catch (error) {
    console.error(`Error updating extraction status: ${error.message}`);
  }
}

/**
 * Enhanced URL validation
 * @param {string} url - URL to validate
 * @return {boolean} Whether URL is valid
 */
function isValidUrl(url) {
  try {
    const regex = /^(https?:\/\/)?([\da-z\.-]+)\.([a-z\.]{2,6})([\/\w \.-]*)*\/?$/;
    return regex.test(url);
  } catch (e) {
    return false;
  }
}

/**
 * Fetch with retry logic and control for pause/stop
 * @param {string} url - URL to fetch
 * @param {Object} config - Configuration options
 * @param {string} extractionId - Unique ID for this extraction process
 * @return {HTTPResponse|null} The response or null on failure
 */
function fetchWithRetryAndControl(url, config, extractionId = null) {
  const options = {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      "User-Agent": config.USER_AGENT
    },
    validateHttpsCertificates: true
  };
  
  let lastError = null;
  
  for (let attempt = 0; attempt < config.MAX_RETRIES; attempt++) {
    // Check if extraction is stopped
    if (extractionId && extractionStates[extractionId] && extractionStates[extractionId].stopped) {
      console.log("Extraction stopped during fetch retry");
      return null;
    }
    
    // Wait if paused
    waitIfPaused(extractionId);
    
    try {
      return UrlFetchApp.fetch(url, options);
    } catch (error) {
      lastError = error;
      // Exponential backoff
      Utilities.sleep(Math.pow(2, attempt) * 1000 + Math.random() * 1000);
    }
  }
  
  console.log("All retry attempts failed: " + lastError);
  return null;
}

/**
 * Wait if extraction is paused
 * @param {string} extractionId - Unique ID for this extraction process
 */
function waitIfPaused(extractionId) {
  if (!extractionId || !extractionStates[extractionId]) return;
  
  // Check if paused in a loop with short intervals
  while (extractionStates[extractionId].paused && !extractionStates[extractionId].stopped) {
    Utilities.sleep(500); // Check every 500ms
  }
}

/**
 * Pause an extraction
 * @param {string} extractionId - Unique ID for the extraction to pause
 * @return {Object} Status response
 */
function pauseExtraction(extractionId) {
  if (!extractionId || !extractionStates[extractionId]) {
    return { success: false, message: "Extraction not found" };
  }
  
  extractionStates[extractionId].paused = true;
  return { success: true, message: "Extraction paused" };
}

/**
 * Resume an extraction
 * @param {string} extractionId - Unique ID for the extraction to resume
 * @return {Object} Status response
 */
function resumeExtraction(extractionId) {
  if (!extractionId || !extractionStates[extractionId]) {
    return { success: false, message: "Extraction not found" };
  }
  
  extractionStates[extractionId].paused = false;
  return { success: true, message: "Extraction resumed" };
}

/**
 * Stop an extraction
 * @param {string} extractionId - Unique ID for the extraction to stop
 * @return {Object} Status response
 */
function stopExtraction(extractionId) {
  if (!extractionId || !extractionStates[extractionId]) {
    return { success: false, message: "Extraction not found" };
  }
  
  extractionStates[extractionId].stopped = true;
  extractionStates[extractionId].paused = false; // Unpause if paused
  
  return { success: true, message: "Extraction stopped" };
}

/**
 * Main function to extract company and product data
 * @param {string} url - The base URL to extract data from
 * @param {Object} config - Configuration options
 * @param {string} extractionId - Unique ID for this extraction process
 * @return {Object} The extracted company and product data
 */
function extractCompanyAndProductData(url, config, extractionId) {
  try {
    updateExtractionProgress(extractionId, 15, "Fetching website content");
    
    // Fetch the base URL
    const response = fetchWithRetryAndControl(url, config, extractionId);
    
    if (!response) {
      console.log(`Failed to fetch URL: ${url}`);
      return null;
    }
    
    const content = response.getContentText();
    
    updateExtractionProgress(extractionId, 25, "Extracting company information");
    
    // Extract company data
    const companyData = extractCompanyInfo(content, url, config, extractionId);
    companyData.url = url;
    companyData.extractionDate = new Date().toISOString();
    
    updateExtractionProgress(extractionId, 40, "Extracting contact information");
    
    // Extract contact information
    const contactInfo = extractContactInfo(content, url);
    companyData.emails = contactInfo.emails.join(", ");
    companyData.phones = contactInfo.phones.join(", ");
    companyData.addresses = contactInfo.addresses.join("; ");
    
    updateExtractionProgress(extractionId, 60, "Discovering product information");
    
    // Extract product information if enabled
    const productInfo = extractProductInfo(content, url, config, extractionId);
    
    // Combine product info with company data
    companyData.productName = productInfo.productName || "";
    companyData.productUrl = productInfo.productUrl || "";
    companyData.productCategory = productInfo.mainCategory || "";
    companyData.productSubcategory = productInfo.subCategory || "";
    companyData.productFamily = productInfo.productFamily || "";
    companyData.quantity = productInfo.productQuantity || "";
    companyData.price = productInfo.price || "";
    companyData.productDescription = productInfo.detailedDescription || "";
    companyData.specifications = productInfo.specifications || "";
    companyData.images = productInfo.productImages || "";
    
    updateExtractionProgress(extractionId, 80, "Enriching data with AI analysis");
    
    // Enrich data with AI if API key is available
    if (config.API.AZURE.OPENAI.KEY) {
      try {
        const enrichedData = enrichDataWithAI(companyData, config);
        if (enrichedData) {
          // Merge AI enriched data
          Object.assign(companyData, enrichedData);
          globalStats.apiCalls.azure.success++;
        }
      } catch (error) {
        console.error(`AI enrichment error: ${error.message}`);
        globalStats.apiCalls.azure.fail++;
      }
    }
    
    updateExtractionProgress(extractionId, 90, "Finalizing extraction");
    
    // Verify with Knowledge Graph if API key is available
    if (config.API.KNOWLEDGE_GRAPH.KEY) {
      try {
        const knowledgeGraphData = queryKnowledgeGraph(companyData.companyName, config);
        if (knowledgeGraphData) {
          // Merge knowledge graph data
          if (knowledgeGraphData.companyType && !companyData.companyType) {
            companyData.companyType = knowledgeGraphData.companyType;
          }
          globalStats.apiCalls.knowledgeGraph.success++;
        }
      } catch (error) {
        console.error(`Knowledge Graph query error: ${error.message}`);
        globalStats.apiCalls.knowledgeGraph.fail++;
      }
    }
    
    return companyData;
    
  } catch (error) {
    console.error(`Error extracting company and product data: ${error.message}`);
    return null;
  }
}

/**
 * Extract company information from HTML content
 * @param {string} content - HTML content
 * @param {string} url - Source URL
 * @param {Object} config - Configuration options
 * @param {string} extractionId - Unique ID for this extraction process
 * @return {Object} Company information
 */
function extractCompanyInfo(content, url, config, extractionId) {
  try {
    // Initialize company data object
    const companyData = {
      companyName: "",
      companyDescription: "",
      companyType: ""
    };
    
    // Extract company name (from title, meta tags, or prominent headings)
    const titleMatch = content.match(/<title>(.*?)<\/title>/i);
    if (titleMatch) {
      // Clean up title to get company name
      let title = titleMatch[1].trim();
      
      // Remove common suffixes like "Home", "Official Website", etc.
      title = title.replace(/\s*[-|]\s*(Home|Official Website|Official Site|Welcome).*$/i, "");
      title = title.replace(/\s*[-|]\s*.*?(homepage|official).*$/i, "");
      
      companyData.companyName = title;
    }
    
    // Try to get a more precise company name from structured data
    const structuredData = extractStructuredData(content);
    if (structuredData && structuredData.organization && structuredData.organization.name) {
      companyData.companyName = structuredData.organization.name;
    }
    
    // Look for organization name in common patterns
    const orgNameMatch = content.match(/<meta\s+(?:property|name)="(?:og:site_name|twitter:site)"[^>]*content="([^"]*)"[^>]*>/i);
    if (orgNameMatch) {
      companyData.companyName = orgNameMatch[1].trim();
    }
    
    // Look for company name in common heading patterns
    const headingMatch = content.match(/<h1[^>]*>(.*?)<\/h1>/i);
    if (headingMatch) {
      const headingText = cleanHtml(headingMatch[1]);
      if (headingText.length < 50) { // Avoid long headings that are likely not company names
        companyData.companyName = headingText;
      }
    }
    
    // Extract logo URL
    const logoPatterns = [
      /<img[^>]*\b(?:id|class)="[^"]*\b(?:logo|brand|company-logo)\b[^"]*"[^>]*src="([^"]*)"/i,
      /<img[^>]*\balt="[^"]*\b(?:logo|brand|company-logo)\b[^"]*"[^>]*src="([^"]*)"/i,
      /<img[^>]*\bsrc="([^"]*logo[^"]*)"/i
    ];
    
    for (const pattern of logoPatterns) {
      const logoMatch = content.match(pattern);
      if (logoMatch) {
        companyData.logo = resolveUrl(logoMatch[1], url);
        break;
      }
    }
    
    // Extract company description
    const metaDescription = content.match(/<meta\s+name="description"\s+content="([^"]*)"/i);
    if (metaDescription) {
      companyData.companyDescription = metaDescription[1].trim();
    }
    
    // Look for about us sections for better description
    const aboutSectionPatterns = [
      /<(?:div|section)[^>]*\b(?:id|class)="[^"]*\babout\b[^"]*"[^>]*>([\s\S]*?)<\/(?:div|section)>/i,
      /<h\d[^>]*>\s*About\s+(?:Us|Company)\s*<\/h\d>([\s\S]*?)(?:<h\d|<\/div|<\/section)/i
    ];
    
    for (const pattern of aboutSectionPatterns) {
      const aboutMatch = content.match(pattern);
      if (aboutMatch) {
        const aboutText = cleanHtml(aboutMatch[1]);
        if (aboutText.length > companyData.companyDescription.length) {
          companyData.companyDescription = aboutText;
        }
        break;
      }
    }
    
    // Try OG description for better company description
    const ogDescription = content.match(/<meta\s+property="og:description"\s+content="([^"]*)"/i);
    if (ogDescription && (!companyData.companyDescription || companyData.companyDescription.length < ogDescription[1].length)) {
      companyData.companyDescription = ogDescription[1].trim();
    }
    
    // Extract company type
    // This is more challenging as it's often not explicitly stated
    // Look for industry keywords or categories
    const industryKeywords = [
      {regex: /\b(?:tech|software|application|app|digital|IT|information technology)\b/i, type: "Technology"},
      {regex: /\b(?:manufacturing|factory|production|industrial)\b/i, type: "Manufacturing"},
      {regex: /\b(?:retail|shop|store|e-commerce|marketplace)\b/i, type: "Retail"},
      {regex: /\b(?:healthcare|medical|hospital|clinic|pharma|health)\b/i, type: "Healthcare"},
      {regex: /\b(?:financial|bank|insurance|investment|finance)\b/i, type: "Financial Services"},
      {regex: /\b(?:education|school|university|college|learning|teaching)\b/i, type: "Education"},
      {regex: /\b(?:food|restaurant|catering|bakery|café)\b/i, type: "Food & Beverage"},
      {regex: /\b(?:media|news|publishing|entertainment)\b/i, type: "Media & Entertainment"},
      {regex: /\b(?:construction|building|architecture|engineering)\b/i, type: "Construction"},
      {regex: /\b(?:travel|tourism|hotel|accommodation|vacation)\b/i, type: "Travel & Tourism"},
      {regex: /\b(?:consulting|consultant|advisor|professional services)\b/i, type: "Consulting"},
      {regex: /\b(?:automotive|car|vehicle|transportation)\b/i, type: "Automotive"},
      {regex: /\b(?:energy|utility|power|oil|gas|electricity)\b/i, type: "Energy & Utilities"},
      {regex: /\b(?:tofu|vegan|plant-based|vegetarian|organic food)\b/i, type: "Plant-based Foods"}
    ];
    
    // First try to find company type in about section or description
    const textToAnalyze = companyData.companyDescription || content;
    
    for (const industry of industryKeywords) {
      if (industry.regex.test(textToAnalyze)) {
        companyData.companyType = industry.type;
        break;
      }
    }
    
    // If we found company name, but not type, look at URL for clues
    if (companyData.companyName && !companyData.companyType) {
      const domainParts = extractDomainParts(url);
      for (const industry of industryKeywords) {
        if (domainParts.some(part => industry.regex.test(part))) {
          companyData.companyType = industry.type;
          break;
        }
      }
    }
    
    updateExtractionProgress(extractionId, 30, "Company information extracted");
    
    // Update stats
    if (companyData.companyName) globalStats.companyData.found++;
    if (companyData.companyDescription) globalStats.companyData.descriptions++;
    if (companyData.companyType) globalStats.companyData.types++;
    
    return companyData;
  } catch (e) {
    console.error(`Error extracting company info: ${e.message}`);
    return {
      companyName: "",
      companyDescription: "",
      companyType: ""
    };
  }
}

/**
 * Extract contact information from HTML content
 * @param {string} content - HTML content
 * @param {string} url - Source URL
 * @return {Object} Contact information (emails, phones, addresses)
 */
function extractContactInfo(content, url) {
  try {
    const contactInfo = {
      emails: [],
      phones: [],
      addresses: []
    };
    
    // Look for contact page link
    let contactPageUrl = null;
    const contactLinkPatterns = [
      /<a[^>]*\bhref="([^"]*contact[^"]*)"/i,
      /<a[^>]*\bhref="([^"]*about-us[^"]*)"/i,
      /<a[^>]*\bhref="([^"]*get-in-touch[^"]*)"/i
    ];
    
    for (const pattern of contactLinkPatterns) {
      const match = content.match(pattern);
      if (match) {
        contactPageUrl = resolveUrl(match[1], url);
        break;
      }
    }
    
    // If contact page found, fetch and analyze it
    let contactPageContent = "";
    if (contactPageUrl && contactPageUrl !== url) {
      try {
        const response = UrlFetchApp.fetch(contactPageUrl, {
          muteHttpExceptions: true,
          followRedirects: true
        });
        contactPageContent = response.getContentText();
      } catch (error) {
        console.error(`Error fetching contact page: ${error.message}`);
      }
    }
    
    // Combine main content and contact page content for analysis
    const combinedContent = content + (contactPageContent || "");
    
    // Extract emails
    const emailPattern = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b/g;
    const emailMatches = combinedContent.match(emailPattern) || [];
    
    if (emailMatches && emailMatches.length > 0) {
      contactInfo.emails = Array.from(new Set(emailMatches)); // Remove duplicates
      globalStats.companyData.emails += contactInfo.emails.length;
    }
    
    // Extract phone numbers (various formats)
    const phonePatterns = [
      /\b\+\d{1,3}[\s.-]?\(?\d{1,4}\)?[\s.-]?\d{1,4}[\s.-]?\d{1,9}\b/g,  // International format
      /\b\(\d{3}\)[\s.-]?\d{3}[\s.-]?\d{4}\b/g,  // US format (xxx) xxx-xxxx
      /\b\d{3}[\s.-]?\d{3}[\s.-]?\d{4}\b/g,      // Simple format xxx-xxx-xxxx
      /\b\d{2,3}[\s.-]?\d{2,4}[\s.-]?\d{4,5}\b/g  // European formats
    ];
    
    const foundPhones = [];
    for (const pattern of phonePatterns) {
      const phoneMatches = combinedContent.match(pattern) || [];
      foundPhones.push(...phoneMatches);
    }
    
    if (foundPhones.length > 0) {
      contactInfo.phones = Array.from(new Set(foundPhones)); // Remove duplicates
      globalStats.companyData.phones += contactInfo.phones.length;
    }
    
    // Extract addresses
    // Look for contact section
    const contactSectionPatterns = [
      /<(?:div|section)[^>]*\b(?:id|class)="[^"]*\b(?:contact|address|location)[^"]*"[^>]*>([\s\S]*?)<\/(?:div|section)>/i,
      /<h\d[^>]*>\s*(?:Contact|Address|Location|Find Us)\s*<\/h\d>([\s\S]*?)(?:<h\d|<\/div|<\/section)/i
    ];
    
    let contactSection = "";
    for (const pattern of contactSectionPatterns) {
      const match = combinedContent.match(pattern);
      if (match) {
        contactSection = match[1];
        break;
      }
    }
    
    // If contact section found, look for address patterns
    if (contactSection) {
      // Look for potential address elements
      const addressElements = [
        ...contactSection.matchAll(/<p[^>]*>([\s\S]*?)<\/p>/gi),
        ...contactSection.matchAll(/<(?:div|span)[^>]*\b(?:id|class)="[^"]*\b(?:address|location)[^"]*"[^>]*>([\s\S]*?)<\/(?:div|span)>/gi)
      ];
      
      for (const element of addressElements) {
        if (element && element[1]) {
          const cleanAddress = cleanHtml(element[1]);
          
          // Check if this looks like an address (contains numbers and common address words)
          if (/\d+/.test(cleanAddress) && 
              /\b(?:street|st|avenue|ave|road|rd|boulevard|blvd|lane|ln|drive|dr|way|place|pl|square|sq|county|city|town|village|state|province|country)\b/i.test(cleanAddress)) {
            contactInfo.addresses.push(cleanAddress);
          }
        }
      }
    }
    
    // If no addresses found yet, try generic patterns
    if (contactInfo.addresses.length === 0) {
      const addressPatterns = [
        // Street, City, State ZIP format
        /\d+\s+[A-Za-z0-9\s.,]+\s+(?:Street|St|Avenue|Ave|Road|Rd|Boulevard|Blvd|Drive|Dr|Lane|Ln|Court|Ct|Way|Place|Pl|Square|Sq)[,.\s]*(?:[A-Za-z\s]+)[,.\s]*(?:[A-Z]{2}|\b[A-Za-z]+\b)[,.\s]*(?:\d{5}(?:-\d{4})?)?/gi,
        
        // European format
        /\d+\s+[A-Za-z0-9\s.,]+\s+(?:Street|St|Avenue|Ave|Road|Rd|Boulevard|Blvd)[,.\s]*(?:[A-Za-z\s]+)[,.\s]*(?:[A-Z]{1,2}\d{1,2}\s+\d[A-Z]{2}|\d{4,5})/gi,
        
        // P.O. Box format
        /P\.?O\.?\s+Box\s+\d+[,.\s]*(?:[A-Za-z\s]+)[,.\s]*(?:[A-Z]{2}|\b[A-Za-z]+\b)[,.\s]*(?:\d{5}(?:-\d{4})?)?/gi
      ];
      
      for (const pattern of addressPatterns) {
        const addressMatches = combinedContent.match(pattern) || [];
        contactInfo.addresses.push(...addressMatches.map(addr => addr.trim()));
      }
    }
    
    // Remove duplicates and update stats
    contactInfo.addresses = Array.from(new Set(contactInfo.addresses));
    if (contactInfo.addresses.length > 0) {
      globalStats.companyData.addresses += contactInfo.addresses.length;
    }
    
    return contactInfo;
  } catch (e) {
    console.error(`Error extracting contact info: ${e.message}`);
    return {
      emails: [],
      phones: [],
      addresses: []
    };
  }
}

/**
 * Extract product information from HTML content
 * @param {string} content - HTML content
 * @param {string} url - Source URL
 * @param {Object} config - Configuration options
 * @param {string} extractionId - Unique ID for this extraction process
 * @return {Object} Product information
 */
function extractProductInfo(content, url, config, extractionId) {
  try {
    updateExtractionProgress(extractionId, 45, "Analyzing product structure");
    
    const productInfo = {
      mainCategory: "",
      subCategory: "",
      productFamily: "",
      productName: "",
      productUrl: "",
      productQuantity: "",
      price: "",
      detailedDescription: "",
      specifications: "",
      productImages: ""
    };
    
    // 1. Look for product listings on the main page
    const productLinks = extractProductLinks(content, url);
    
    // 2. Extract categories from menu structure or breadcrumbs
    const categories = extractCategories(content);
    if (categories.length > 0) {
      productInfo.mainCategory = categories[0] || "";
      if (categories.length > 1) {
        productInfo.subCategory = categories[1] || "";
      }
      if (categories.length > 2) {
        productInfo.productFamily = categories[2] || "";
      }
      
      // Update global stats
      globalStats.productData.categories += categories.length;
    }
    
    updateExtractionProgress(extractionId, 50, "Discovering products");
    
    // If no product links found, look for products section link
    let productsPageUrl = null;
    if (productLinks.length === 0) {
      const productsLinkPatterns = [
        /<a[^>]*\bhref="([^"]*products[^"]*)"/i,
        /<a[^>]*\bhref="([^"]*catalogue[^"]*)"/i,
        /<a[^>]*\bhref="([^"]*catalog[^"]*)"/i,
        /<a[^>]*\bhref="([^"]*shop[^"]*)"/i
      ];
      
      for (const pattern of productsLinkPatterns) {
        const match = content.match(pattern);
        if (match) {
          productsPageUrl = resolveUrl(match[1], url);
          break;
        }
      }
      
      // If products page found, fetch and analyze it
      if (productsPageUrl && productsPageUrl !== url) {
        try {
          const response = UrlFetchApp.fetch(productsPageUrl, {
            muteHttpExceptions: true,
            followRedirects: true
          });
          const productsPageContent = response.getContentText();
          
          // Extract product links from products page
          const additionalLinks = extractProductLinks(productsPageContent, productsPageUrl);
          productLinks.push(...additionalLinks);
        } catch (error) {
          console.error(`Error fetching products page: ${error.message}`);
        }
      }
    }
    
    // If still no product links, try to extract product data from the current page
    if (productLinks.length === 0) {
      // Look for product sections directly on the page
      const productSectionPatterns = [
        /<(?:div|section)[^>]*\b(?:id|class)="[^"]*\b(?:product|item)[^"]*"[^>]*>([\s\S]*?)<\/(?:div|section)>/i,
        /<h\d[^>]*>\s*(?:Products|Our Products|Featured Products)\s*<\/h\d>([\s\S]*?)(?:<h\d|<\/div|<\/section)/i
      ];
      
      for (const pattern of productSectionPatterns) {
        const match = content.match(pattern);
        if (match) {
          const productData = extractProductDetails(match[1], url);
          
          if (productData.name) {
            productInfo.productName = productData.name;
            productInfo.productUrl = url;
            productInfo.detailedDescription = productData.description || "";
            productInfo.price = productData.price || "";
            productInfo.productQuantity = productData.quantity || "";
            productInfo.specifications = productData.specifications || "";
            productInfo.productImages = productData.images.join(", ");
            
            // Update global stats
            globalStats.productData.found++;
            if (productData.images.length > 0) {
              globalStats.productData.images += productData.images.length;
            }
            if (productData.description) {
              globalStats.productData.descriptions++;
            }
            
            break;
          }
        }
      }
    }
    
    // 3. Follow product links if configured to do so and we have links
    if (config.EXTRACTION.FOLLOW_LINKS && productLinks.length > 0) {
      updateExtractionProgress(extractionId, 55, "Analyzing product details");
      
      // Limit the number of products to process
      const linksToProcess = productLinks.slice(0, Math.min(productLinks.length, config.MAX_PRODUCT_PAGES));
      const allProducts = [];
      
      // Process each product link
      for (let i = 0; i < linksToProcess.length; i++) {
        const productLink = linksToProcess[i];
        
        // Update progress
        updateExtractionProgress(
          extractionId, 
          55 + Math.floor((i / linksToProcess.length) * 20), 
          `Processing product ${i+1} of ${linksToProcess.length}`
        );
        
        // Check if extraction is stopped
        if (extractionId && extractionStates[extractionId] && extractionStates[extractionId].stopped) {
          console.log("Extraction stopped during product processing");
          break;
        }
        
        // Wait if paused
        waitIfPaused(extractionId);
        
        // Fetch and process product page
        console.log(`Processing product: ${productLink.url}`);
        const productData = processProductPage(productLink.url, config, extractionId);
        
        if (productData) {
          // Add to products array
          allProducts.push({
            name: productData.name,
            url: productLink.url,
            description: productData.description,
            price: productData.price,
            quantity: productData.quantity,
            specifications: productData.specifications,
            images: productData.images
          });
          
          // Update global stats
          globalStats.productData.found++;
          if (productData.images.length > 0) {
            globalStats.productData.images += productData.images.length;
          }
          if (productData.description) {
            globalStats.productData.descriptions++;
          }
        }
        
        // Small delay between product page requests
        Utilities.sleep(300);
      }
      
      // Process all collected products
      if (allProducts.length > 0) {
        // Concatenate product names with commas
        productInfo.productName = allProducts.map(p => p.name).filter(Boolean).join(", ");
        
        // Use the first product's URL as the main product URL
        if (allProducts[0].url) {
          productInfo.productUrl = allProducts[0].url;
        }
        
        // Collect all unique product images
        const allImages = new Set();
        allProducts.forEach(product => {
          if (product.images && Array.isArray(product.images)) {
            product.images.forEach(img => allImages.add(img));
          }
        });
        productInfo.productImages = Array.from(allImages).join(", ");
        
        // Combine descriptions (up to a reasonable length)
        const descriptions = allProducts
          .map(p => p.description)
          .filter(Boolean)
          .join("\n\n")
          .slice(0, 1000);
        
        if (descriptions) {
          productInfo.detailedDescription = descriptions;
        }
      }
    }
    
    return productInfo;
  } catch (e) {
    console.error(`Error extracting product info: ${e.message}`);
    return {
      mainCategory: "",
      subCategory: "",
      productFamily: "",
      productName: "",
      productUrl: "",
      productQuantity: "",
      price: "",
      detailedDescription: "",
      specifications: "",
      productImages: ""
    };
  }
}

/**
 * Extract product links from HTML content
 * @param {string} content - HTML content
 * @param {string} baseUrl - Base URL for resolving relative links
 * @return {Array} Array of product links with text and URL
 */
function extractProductLinks(content, baseUrl) {
  try {
    const productLinks = [];
    
    // Look for links in product sections, grids, or listings
    const productSectionPatterns = [
      /<(?:div|section|ul)[^>]*\b(?:id|class)="[^"]*\b(?:product|item|listing|catalog|shop)[^"]*"[^>]*>([\s\S]*?)<\/(?:div|section|ul)>/gi,
      /<h\d[^>]*>\s*(?:Products|Our Products|Featured Products|Shop|Catalog)\s*<\/h\d>([\s\S]*?)(?:<h\d|<footer|<\/main|<\/body)/gi
    ];
    
    let productSections = [];
    
    // Extract product sections
    for (const pattern of productSectionPatterns) {
      let match;
      const regex = new RegExp(pattern);
      while ((match = regex.exec(content)) !== null) {
        productSections.push(match[1]);
      }
    }
    
    // If no dedicated product sections found, use the whole content as fallback
    if (productSections.length === 0) {
      productSections = [content];
    }
    
    // Extract links from product sections
    for (const section of productSections) {
      const linkPattern = /<a\s+[^>]*href="([^"]*)"[^>]*>([\s\S]*?)<\/a>/gi;
      let linkMatch;
      
      while ((linkMatch = linkPattern.exec(section)) !== null) {
        const href = linkMatch[1];
        const text = cleanHtml(linkMatch[2]).trim();
        
        // Skip empty links, non-product links, or navigation links
        if (!href || href === "#" || href.startsWith("javascript:") || 
            href.includes("login") || href.includes("cart") || 
            href.includes("account") || href.includes("contact")) {
          continue;
        }
        
        // Skip links without any text content
        if (!text) continue;
        
        // Resolve relative URL
        const fullUrl = resolveUrl(href, baseUrl);
        
        // Only include links from the same domain
        if (isSameDomain(fullUrl, baseUrl) && !isExcludedPath(fullUrl)) {
          // Check if this link seems product-related
          if (isLikelyProductLink(fullUrl, text)) {
            productLinks.push({
              url: fullUrl,
              text: text
            });
          }
        }
      }
    }
    
    // If we found product links with product section patterns, return those
    if (productLinks.length > 0) {
      return productLinks;
    }
    
    // If no product links found yet, look for image-based product links
    const imageProductPatterns = [
      /<div[^>]*\b(?:id|class)="[^"]*\b(?:product|item)[^"]*"[^>]*>[\s\S]*?<img[^>]*src="([^"]*)"[\s\S]*?<a[^>]*href="([^"]*)"[^>]*>([\s\S]*?)<\/a>/gi,
      /<a[^>]*href="([^"]*)"[^>]*>[\s\S]*?<img[^>]*\b(?:id|class)="[^"]*\b(?:product|item)[^"]*"[^>]*src="([^"]*)"[\s\S]*?<\/a>/gi
    ];
    
    for (const pattern of imageProductPatterns) {
      let match;
      const regex = new RegExp(pattern);
      while ((match = regex.exec(content)) !== null) {
        let href, text;
        
        // Extract URL and text based on pattern match
        if (pattern.toString().indexOf('img[^>]*src="([^"]*)"') < pattern.toString().indexOf('a[^>]*href="([^"]*)"')) {
          // First pattern: image then link
          href = match[2];
          text = cleanHtml(match[3]).trim();
        } else {
          // Second pattern: link then image
          href = match[1];
          text = ""; // No direct text, use URL elements
        }
        
        // Skip non-product links
        if (!href || href === "#" || href.startsWith("javascript:")) continue;
        
        // Resolve relative URL
        const fullUrl = resolveUrl(href, baseUrl);
        
        // Only include links from the same domain
        if (isSameDomain(fullUrl, baseUrl) && !isExcludedPath(fullUrl)) {
          // If no text content, use URL path as text
          if (!text) {
            const urlPath = new URL(fullUrl).pathname;
            const pathSegments = urlPath.split('/').filter(Boolean);
            if (pathSegments.length > 0) {
              text = pathSegments[pathSegments.length - 1].replace(/-|_/g, ' ');
            }
          }
          
          // Add to product links
          if (text) {
            productLinks.push({
              url: fullUrl,
              text: text
            });
          }
        }
      }
    }
    
    return productLinks;
  } catch (e) {
    console.error(`Error extracting product links: ${e.message}`);
    return [];
  }
}

/**
 * Check if a URL path should be excluded from product links
 * @param {string} url - URL to check
 * @return {boolean} Whether the URL should be excluded
 */
function isExcludedPath(url) {
  try {
    const excludedPaths = [
      'about', 'contact', 'privacy', 'terms', 'faq', 'help', 'support',
      'blog', 'news', 'login', 'register', 'account', 'cart', 'checkout',
      'search', 'sitemap', 'careers', 'jobs', 'press', 'media'
    ];
    
    const parsedUrl = new URL(url);
    const path = parsedUrl.pathname.toLowerCase();
    
    return excludedPaths.some(excluded => 
      path === `/${excluded}` || 
      path === `/${excluded}/` || 
      path.includes(`/${excluded}/`)
    );
  } catch (e) {
    return false;
  }
}

/**
 * Check if a link is likely a product link based on URL and text
 * @param {string} url - Link URL
 * @param {string} text - Link text
 * @return {boolean} Whether the link is likely a product link
 */
function isLikelyProductLink(url, text) {
  try {
    // Check URL for product-related terms
    const productTermsInUrl = [
      'product', 'item', 'shop', 'buy', 'purchase', 'catalog', 'catalogue',
      'collection', 'goods', 'merchandise', 'sale', 'order', 'category'
    ];
    
    const parsedUrl = new URL(url);
    const urlLower = parsedUrl.href.toLowerCase();
    const pathLower = parsedUrl.pathname.toLowerCase();
    
    const hasProductTermInUrl = productTermsInUrl.some(term => 
      urlLower.includes(term) || 
      pathLower.includes(term)
    );
    
    // Check if URL has product ID pattern
    const hasProductIdPattern = /\/p\/|\/product\/|\/item\/|\/prod[_-]?id\/|\/sku\/|\/id\/\d+/.test(pathLower);
    
    // Check text for product indicators
    const textLower = text.toLowerCase();
    const hasProductIndicatorInText = textLower.includes('buy') || 
                                      textLower.includes('shop') || 
                                      textLower.includes('view') ||
                                      textLower.includes('details') ||
                                      textLower.includes('more');
    
    // Check if link text is concise (likely product name) and not navigational
    const isConciseText = text.length > 0 && text.length < 50 && 
                        !textLower.includes('about') && 
                        !textLower.includes('contact') &&
                        !textLower.includes('home');
    
    return hasProductTermInUrl || hasProductIdPattern || hasProductIndicatorInText || isConciseText;
  } catch (e) {
    return false;
  }
}

/**
 * Extract category information from HTML content
 * @param {string} content - HTML content
 * @return {Array} Array of categories (main, sub, family)
 */
function extractCategories(content) {
  try {
    const categories = [];
    
    // Try to extract from breadcrumbs
    const breadcrumbPatterns = [
      /<(?:nav|div|ul)[^>]*\b(?:id|class)="[^"]*\b(?:breadcrumb|path|navigation)[^"]*"[^>]*>([\s\S]*?)<\/(?:nav|div|ul)>/i,
      /<ol[^>]*\b(?:id|class)="[^"]*\b(?:breadcrumb)[^"]*"[^>]*>([\s\S]*?)<\/ol>/i
    ];
    
    for (const pattern of breadcrumbPatterns) {
      const breadcrumbMatch = content.match(pattern);
      if (breadcrumbMatch) {
        const breadcrumbContent = breadcrumbMatch[1];
        const linkPattern = /<a[^>]*>([\s\S]*?)<\/a>/gi;
        let linkMatch;
        
        while ((linkMatch = linkPattern.exec(breadcrumbContent)) !== null) {
          const text = cleanHtml(linkMatch[1]).trim();
          
          // Skip "Home", "Index", etc.
          if (text && !["home", "index", "main", "start"].includes(text.toLowerCase())) {
            categories.push(text);
          }
        }
        
        if (categories.length > 0) {
          break; // Found categories in breadcrumbs
        }
      }
    }
    
    // If breadcrumbs not found, try menu categories
    if (categories.length === 0) {
      const menuPatterns = [
        /<(?:nav|div|ul)[^>]*\b(?:id|class)="[^"]*\b(?:menu|main-menu|primary-menu|navigation)[^"]*"[^>]*>([\s\S]*?)<\/(?:nav|div|ul)>/i,
        /<(?:nav|div|ul)[^>]*\b(?:id|class)="[^"]*\b(?:category|categories|navbar)[^"]*"[^>]*>([\s\S]*?)<\/(?:nav|div|ul)>/i
      ];
      
      for (const pattern of menuPatterns) {
        const menuMatch = content.match(pattern);
        if (menuMatch) {
          const menuContent = menuMatch[1];
          
          // Try to find category links
          const categoryPaths = [];
          const firstLevelItems = [...menuContent.matchAll(/<li[^>]*>([\s\S]*?)<\/li>/gi)];
          
          for (const item of firstLevelItems) {
            const itemContent = item[1];
            const linkMatch = itemContent.match(/<a[^>]*>([\s\S]*?)<\/a>/i);
            
            if (linkMatch) {
              const text = cleanHtml(linkMatch[1]).trim();
              
              // Skip "Home", "About", "Contact", etc.
              if (text && !["home", "about", "contact", "faq", "help"].includes(text.toLowerCase())) {
                // Check for submenu
                const submenuMatch = itemContent.match(/<ul[^>]*>([\s\S]*?)<\/ul>/i);
                
                if (submenuMatch) {
                  // This is a category with subcategories
                  const submenuContent = submenuMatch[1];
                  const subItems = [...submenuContent.matchAll(/<li[^>]*>([\s\S]*?)<\/li>/gi)];
                  
                  for (const subItem of subItems) {
                    const subItemContent = subItem[1];
                    const subLinkMatch = subItemContent.match(/<a[^>]*>([\s\S]*?)<\/a>/i);
                    
                    if (subLinkMatch) {
                      const subText = cleanHtml(subLinkMatch[1]).trim();
                      
                      if (subText) {
                        categoryPaths.push([text, subText]);
                      }
                    }
                  }
                } else {
                  // This is a standalone category
                  categoryPaths.push([text]);
                }
              }
            }
          }
          
          // If we found category paths, use the first one
          if (categoryPaths.length > 0) {
            return categoryPaths[0];
          }
        }
      }
    }
    
    // If still no categories found, look for h1/h2 headers that might indicate categories
    if (categories.length === 0) {
      const h1Match = content.match(/<h1[^>]*>([\s\S]*?)<\/h1>/i);
      if (h1Match) {
        const h1Text = cleanHtml(h1Match[1]).trim();
        
        // Check if this looks like a category (not too long, not site title)
        if (h1Text && h1Text.length < 50 && !h1Text.toLowerCase().includes("welcome")) {
          categories.push(h1Text);
          
          // Look for potential subcategory in h2
          const h2Match = content.match(/<h2[^>]*>([\s\S]*?)<\/h2>/i);
          if (h2Match) {
            const h2Text = cleanHtml(h2Match[1]).trim();
            if (h2Text && h2Text.length < 50) {
              categories.push(h2Text);
            }
          }
        }
      }
    }
    
    return categories;
  } catch (e) {
    console.error(`Error extracting categories: ${e.message}`);
    return [];
  }
}

/**
 * Process a product page to extract detailed product information
 * @param {string} url - Product page URL
 * @param {Object} config - Configuration options
 * @param {string} extractionId - Unique ID for this extraction process
 * @return {Object} Detailed product information
 */
function processProductPage(url, config, extractionId) {
  try {
    const response = fetchWithRetryAndControl(url, config, extractionId);
    
    if (!response) {
      console.log(`Failed to fetch product page: ${url}`);
      return null;
    }
    
    const content = response.getContentText();
    
    // Extract product details from the page
    return extractProductDetails(content, url);
    
  } catch (e) {
    console.error(`Error processing product page: ${e.message}`);
    return null;
  }
}

/**
 * Extract product details from HTML content
 * @param {string} content - HTML content of product page
 * @param {string} url - Product page URL
 * @return {Object} Detailed product information
 */
function extractProductDetails(content, url) {
  try {
    const productData = {
      name: "",
      description: "",
      price: "",
      quantity: "",
      specifications: "",
      images: []
    };
    
    // Extract product name (prioritize structured data)
    const structuredData = extractStructuredData(content);
    if (structuredData && structuredData.product && structuredData.product.name) {
      productData.name = structuredData.product.name;
    } else {
      // Try common product name patterns
      const productNamePatterns = [
        /<h1[^>]*\b(?:id|class)="[^"]*\b(?:product|item|title)[^"]*"[^>]*>([\s\S]*?)<\/h1>/i,
        /<h1[^>]*>([\s\S]*?)<\/h1>/i,
        /<div[^>]*\b(?:id|class)="[^"]*\b(?:product|item)[^"]*-title"[^>]*>([\s\S]*?)<\/div>/i
      ];
      
      for (const pattern of productNamePatterns) {
        const match = content.match(pattern);
        if (match) {
          productData.name = cleanHtml(match[1]).trim();
          break;
        }
      }
      
      // If still no name, use page title as fallback
      if (!productData.name) {
        const titleMatch = content.match(/<title>(.*?)<\/title>/i);
        if (titleMatch) {
          productData.name = titleMatch[1].trim().split('|')[0].split('-')[0].trim();
        }
      }
    }
    
    // Extract product price
    if (structuredData && structuredData.product && structuredData.product.offers && 
        structuredData.product.offers.price) {
      productData.price = structuredData.product.offers.price;
    } else {
      // Try common price patterns
      const pricePatterns = [
        /<(?:div|span)[^>]*\b(?:id|class)="[^"]*\b(?:price|product-price)[^"]*"[^>]*>([\s\S]*?)<\/(?:div|span)>/i,
        /<meta[^>]*\bitemprop="price"[^>]*\bcontent="([^"]*)"/i,
        /[$€£¥]\s*\d+(?:\.\d{1,2})?/,
        /\d+(?:\.\d{1,2})?\s*[$€£¥]/
      ];
      
      for (const pattern of pricePatterns) {
        const match = content.match(pattern);
        if (match) {
          productData.price = cleanHtml(match[1] || match[0]).trim();
          break;
        }
      }
    }
    
    // Extract product description
    if (structuredData && structuredData.product && structuredData.product.description) {
      productData.description = structuredData.product.description;
    } else {
      // Try common description patterns
      const descriptionPatterns = [
        /<(?:div|section)[^>]*\b(?:id|class)="[^"]*\b(?:product|item)[^"]*-description[^"]*"[^>]*>([\s\S]*?)<\/(?:div|section)>/i,
        /<(?:div|section)[^>]*\b(?:id|class)="[^"]*\b(?:description)[^"]*"[^>]*>([\s\S]*?)<\/(?:div|section)>/i,
        /<div[^>]*\bitemprop="description"[^>]*>([\s\S]*?)<\/div>/i,
        /<meta[^>]*\bname="description"[^>]*\bcontent="([^"]*)"/i
      ];
      
      for (const pattern of descriptionPatterns) {
        const match = content.match(pattern);
        if (match) {
          productData.description = cleanHtml(match[1] || match[0]).trim();
          break;
        }
      }
    }
    
    // Extract product quantity/size information
    const quantityPatterns = [
      /<(?:div|span)[^>]*\b(?:id|class)="[^"]*\b(?:size|quantity|volume|weight|dimension)[^"]*"[^>]*>([\s\S]*?)<\/(?:div|span)>/i,
      /<span[^>]*\bitemprop="size"[^>]*>([\s\S]*?)<\/span>/i,
      /<select[^>]*\b(?:id|name)="[^"]*\b(?:size|quantity|volume|weight)[^"]*"[^>]*>([\s\S]*?)<\/select>/i,
      /(\d+(?:\.\d+)?\s*(?:ml|l|g|kg|oz|lb|pack|piece|count|ct))/i
    ];
    
    for (const pattern of quantityPatterns) {
      const match = content.match(pattern);
      if (match) {
        productData.quantity = cleanHtml(match[1] || match[0]).trim();
        break;
      }
    }
    
    // Extract product specifications
    const specPatterns = [
      /<(?:div|section|table)[^>]*\b(?:id|class)="[^"]*\b(?:specification|technical|details|specs)[^"]*"[^>]*>([\s\S]*?)<\/(?:div|section|table)>/i,
      /<(?:div|section)[^>]*\bitemprop="additionalProperty"[^>]*>([\s\S]*?)<\/(?:div|section)>/i,
      /<h\d[^>]*>\s*(?:Specifications|Technical Details|Tech Specs|Additional Information)\s*<\/h\d>([\s\S]*?)(?:<h\d|<\/div|<\/section)/i,
      /<table[^>]*\b(?:id|class)="[^"]*\b(?:product|item)[^"]*-attributes"[^>]*>([\s\S]*?)<\/table>/i
    ];
    
    for (const pattern of specPatterns) {
      const match = content.match(pattern);
      if (match) {
        productData.specifications = cleanHtml(match[1]).trim();
        break;
      }
    }
    
    // Extract product images
    if (structuredData && structuredData.product && structuredData.product.image) {
      const images = Array.isArray(structuredData.product.image) 
        ? structuredData.product.image 
        : [structuredData.product.image];
      
      productData.images = images.map(img => resolveUrl(img, url));
    } else {
      // Try common image patterns
      const imagePatterns = [
        /<img[^>]*\b(?:id|class)="[^"]*\b(?:product|item)[^"]*-image[^"]*"[^>]*src="([^"]*)"/gi,
        /<div[^>]*\b(?:id|class)="[^"]*\b(?:product|item)[^"]*-gallery[^"]*"[^>]*>[\s\S]*?<img[^>]*src="([^"]*)"/gi,
        /<a[^>]*\b(?:id|class|rel)="[^"]*\b(?:lightbox|gallery)[^"]*"[^>]*href="([^"]*)"/gi,
        /<meta[^>]*\bproperty="og:image"[^>]*\bcontent="([^"]*)"/i
      ];
      
      for (const pattern of imagePatterns) {
        let match;
        const regex = new RegExp(pattern);
        while ((match = regex.exec(content)) !== null) {
          const imgUrl = resolveUrl(match[1], url);
          productData.images.push(imgUrl);
        }
      }
      
      // If no product-specific images found, look for any image that might be a product image
      if (productData.images.length === 0) {
        const anyImagePattern = /<img[^>]*src="([^"]*)"[^>]*>/gi;
        let anyImageMatch;
        
        while ((anyImageMatch = anyImagePattern.exec(content)) !== null) {
          const src = anyImageMatch[1];
          
          // Skip tiny images, icons, logos, etc.
          if (src.includes("icon") || src.includes("logo") || src.includes("banner") || 
              src.includes("pixel") || src.endsWith(".svg")) {
            continue;
          }
          
          // Resolve relative URL
          const imageUrl = resolveUrl(src, url);
          productData.images.push(imageUrl);
          
          // Limit to a reasonable number of images
          if (productData.images.length >= 5) break;
        }
      }
    }
    
    return productData;
  } catch (e) {
    console.error(`Error extracting product details: ${e.message}`);
    return {
      name: "",
      description: "",
      price: "",
      quantity: "",
      specifications: "",
      images: []
    };
  }
}

/**
 * Store extracted data in the Data sheet with the simplified schema
 * @param {Sheet} dataSheet - The Data sheet
 * @param {Object} data - The extracted data
 */
function storeExtractedData(dataSheet, data) {
  try {
    // Find if URL already exists
    const dataRange = dataSheet.getDataRange();
    const dataValues = dataRange.getValues();
    let rowIndex = -1;
    
    for (let i = 1; i < dataValues.length; i++) {
      if (dataValues[i][0] === data.url) {
        rowIndex = i + 1; // +1 because array is 0-indexed, but sheets are 1-indexed
        break;
      }
    }
    
    // Prepare row data with the simplified schema
    const rowData = [
      data.url || "",                       // URL
      "Completed",                          // Status
      data.companyName || "",               // Company Name
      data.addresses || "",                 // Addresses
      data.emails || "",                    // Emails
      data.phones || "",                    // Phones 
      data.companyDescription || "",        // Company Description
      data.companyType || "",               // Company Type
      data.productName || "",               // Product Name
      data.productUrl || "",                // Product URL
      data.productCategory || "",           // Product Category
      data.productSubcategory || "",        // Product Subcategory
      data.productFamily || "",             // Product Family
      data.quantity || "",                  // Quantity
      data.price || "",                     // Price
      data.productDescription || "",        // Product Description
      data.specifications || "",            // Specifications
      data.images || "",                    // Images
      data.extractionDate || new Date().toISOString() // Extraction Date
    ];
    
    if (rowIndex > 0) {
      // Update existing row
      dataSheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      // Add new row
      dataSheet.appendRow(rowData);
    }
    
  } catch (error) {
    console.error(`Error storing extracted data: ${error.message}`);
    throw error;
  }
}

/**
 * Process extracted data using the DataProcessor module
 * @param {Object} data - Raw extracted data
 * @param {string} url - Source URL
 * @param {Object} config - Configuration options
 * @return {Object} Processed data
 */
function processExtractedData(data, url, config) {
  try {
    // Clean text fields
    const textFields = ['companyName', 'companyDescription', 'productName', 'productDescription', 'specifications'];
    textFields.forEach(field => {
      if (data[field]) {
        data[field] = cleanText(data[field]);
      }
    });
    
    // Normalize price
    if (data.price) {
      data.price = normalizePrice(data.price);
    }
    
    // Process company type
    if (data.companyDescription && (!data.companyType || data.companyType === "Technology")) {
      data.companyType = determineCompanyType(data.companyDescription);
    }
    
    // Process product category
    if (data.productName && !data.productCategory) {
      data.productCategory = determineProductCategory(data.productName, data.productDescription);
    }
    
    return data;
  } catch (error) {
    console.error(`Error processing extracted data: ${error.message}`);
    return data; // Return original data if processing fails
  }
}

/**
 * Clean text content
 * @param {string} text - Text to clean
 * @returns {string} Cleaned text
 */
function cleanText(text) {
  if (!text) return '';
  
  return text
    .trim()
    .replace(/\s+/g, ' ') // Normalize whitespace
    .replace(/[^\S\r\n]+/g, ' ') // Remove multiple spaces
    .replace(/[\r\n]+/g, '\n') // Normalize line breaks
    .replace(/&nbsp;/g, ' ') // Replace HTML spaces
    .replace(/&amp;/g, '&') // Replace HTML entities
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

/**
 * Normalize price value
 * @param {string|number} price - Price to normalize
 * @returns {string} Normalized price
 */
function normalizePrice(price) {
  if (typeof price === 'number') return price.toString();
  
  // Handle various currency formats
  if (typeof price === 'string') {
    // Remove currency symbols but keep the value
    return price.trim();
  }
  
  return price.toString();
}

/**
 * Determine company type from company description
 * @param {string} description - Company description
 * @returns {string} Company type
 */
function determineCompanyType(description) {
  const typePatterns = [
    { pattern: /\b(?:food|restaurant|catering|bakery|café|organic|vegan|vegetarian|plant-based|tofu|natural food)\b/i, type: "Food & Beverage" },
    { pattern: /\b(?:tech|software|application|app|digital|IT|information technology)\b/i, type: "Technology" },
    { pattern: /\b(?:manufacturing|factory|production|industrial)\b/i, type: "Manufacturing" },
    { pattern: /\b(?:retail|shop|store|e-commerce|marketplace)\b/i, type: "Retail" },
    { pattern: /\b(?:healthcare|medical|hospital|clinic|pharma|health)\b/i, type: "Healthcare" },
    { pattern: /\b(?:financial|bank|insurance|investment|finance)\b/i, type: "Financial Services" },
    { pattern: /\b(?:education|school|university|college|learning|teaching)\b/i, type: "Education" }
  ];
  
  for (const {pattern, type} of typePatterns) {
    if (pattern.test(description)) {
      return type;
    }
  }
  
  return "Other";
}

/**
 * Determine product category from product name and description
 * @param {string} name - Product name
 * @param {string} description - Product description
 * @returns {string} Product category
 */
function determineProductCategory(name, description) {
  const text = (name + " " + (description || "")).toLowerCase();
  
  const categoryPatterns = [
    { pattern: /\b(?:food|meal|snack|drink|beverage|juice|water|soda|coffee|tea)\b/i, category: "Food & Beverage" },
    { pattern: /\b(?:tofu|seitan|vegan|plant-based|vegetarian)\b/i, category: "Plant-based Foods" },
    { pattern: /\b(?:clothing|shirt|pants|jacket|dress|apparel|wear|fashion)\b/i, category: "Clothing" },
    { pattern: /\b(?:electronics|device|gadget|computer|laptop|phone|headphone|speaker)\b/i, category: "Electronics" },
    { pattern: /\b(?:furniture|chair|table|sofa|bed|cabinet|desk)\b/i, category: "Furniture" },
    { pattern: /\b(?:beauty|cosmetic|makeup|skincare|perfume|fragrance)\b/i, category: "Beauty & Personal Care" },
    { pattern: /\b(?:book|novel|textbook|magazine|publication)\b/i, category: "Books & Media" },
    { pattern: /\b(?:toy|game|puzzle|board game|video game)\b/i, category: "Toys & Games" },
    { pattern: /\b(?:tool|hardware|equipment|machinery|device)\b/i, category: "Tools & Equipment" },
    { pattern: /\b(?:sports|fitness|exercise|workout|athletic)\b/i, category: "Sports & Fitness" }
  ];
  
  for (const {pattern, category} of categoryPatterns) {
    if (pattern.test(text)) {
      return category;
    }
  }
  
  return "Other";
}

/**
 * Extract structured data (JSON-LD) from HTML content
 * @param {string} content - HTML content
 * @return {Object} Structured data object
 */
function extractStructuredData(content) {
  try {
    const jsonLdMatch = content.match(/<script[^>]*type="application\/ld\+json"[^>]*>([\s\S]*?)<\/script>/i);
    if (jsonLdMatch) {
      return JSON.parse(jsonLdMatch[1]);
    }
    return {};
  } catch (e) {
    console.error(`Error extracting structured data: ${e.message}`);
    return {};
  }
}

/**
 * Clean HTML content to plain text
 * @param {string} html - HTML content
 * @return {string} Plain text
 */
function cleanHtml(html) {
  if (!html) return '';
  
  // Remove HTML tags
  let text = html.replace(/<[^>]+>/g, ' ');
  
  // Decode HTML entities
  text = text.replace(/&nbsp;/g, ' ');
  text = text.replace(/&amp;/g, '&');
  text = text.replace(/&lt;/g, '<');
  text = text.replace(/&gt;/g, '>');
  text = text.replace(/&quot;/g, '"');
  text = text.replace(/&apos;/g, "'");
  text = text.replace(/&#(\d+);/g, (match, dec) => String.fromCharCode(dec));
  
  // Normalize whitespace
  text = text.replace(/\s+/g, ' ').trim();
  
  return text;
}

/**
 * Resolve relative URL to absolute URL
 * @param {string} url - Relative or absolute URL
 * @param {string} base - Base URL
 * @return {string} Absolute URL
 */
function resolveUrl(url, base) {
  try {
    // Check if URL is already absolute
    if (url.match(/^https?:\/\//)) {
      return url;
    }
    
    // Handle protocol-relative URLs
    if (url.startsWith('//')) {
      const baseProtocol = base.split('://')[0];
      return `${baseProtocol}:${url}`;
    }
    
    // Parse base URL
    let baseUrl = base;
    if (!baseUrl.endsWith('/')) {
      // Remove path component if base doesn't end with /
      const lastSlashIndex = baseUrl.lastIndexOf('/');
      if (lastSlashIndex > 8) { // 8 is the minimum index for 'http://' or 'https://'
        baseUrl = baseUrl.substring(0, lastSlashIndex + 1);
      } else {
        baseUrl = baseUrl + '/';
      }
    }
    
    // Handle root-relative URLs
    if (url.startsWith('/')) {
      // Extract domain from base URL
      const domainMatch = baseUrl.match(/^(https?:\/\/[^/]+)\//);
      if (domainMatch) {
        return domainMatch[1] + url;
      }
      return baseUrl + url.substring(1);
    }
    
    // Handle relative URLs
    return baseUrl + url;
  } catch (e) {
    console.error(`Error resolving URL: ${e.message}`);
    return url;
  }
}

/**
 * Check if two URLs are from the same domain
 * @param {string} url1 - First URL
 * @param {string} url2 - Second URL
 * @return {boolean} Whether URLs are from the same domain
 */
function isSameDomain(url1, url2) {
  try {
    // Extract domain from URLs
    const getDomain = (url) => {
      const match = url.match(/^https?:\/\/([^/]+)/i);
      return match ? match[1] : '';
    };
    
    const domain1 = getDomain(url1);
    const domain2 = getDomain(url2);
    
    // Compare domains
    return domain1 === domain2;
  } catch (e) {
    console.error(`Error comparing domains: ${e.message}`);
    return false;
  }
}

/**
 * Extract domain parts from URL for analysis
 * @param {string} url - URL to analyze
 * @return {Array} Domain parts
 */
function extractDomainParts(url) {
  try {
    const domainMatch = url.match(/^https?:\/\/([^/]+)/i);
    if (domainMatch) {
      const domain = domainMatch[1];
      // Split domain into parts and remove common TLDs
      return domain.split('.')
        .filter(part => !['com', 'org', 'net', 'edu', 'gov', 'co', 'io', 'ai', 'app'].includes(part));
    }
    return [];
  } catch (e) {
    console.error(`Error extracting domain parts: ${e.message}`);
    return [];
  }
}

/**
 * Enrich the extracted data with AI
 * @param {Object} data - Extracted data
 * @param {Object} config - Configuration options
 * @return {Object} Enriched data
 */
function enrichDataWithAI(data, config) {
  try {
    // Skip if no API key or endpoint
    if (!config.API.AZURE.OPENAI.KEY || !config.API.AZURE.OPENAI.ENDPOINT) {
      return null;
    }
    
    const apiKey = config.API.AZURE.OPENAI.KEY;
    const endpoint = config.API.AZURE.OPENAI.ENDPOINT;
    const deployment = config.API.AZURE.OPENAI.DEPLOYMENT;
    
    // Prepare prompt for company data enrichment
    const prompt = `
      Analyze this company data and provide enriched information:
      
      Company Name: ${data.companyName || 'Unknown'}
      Company Description: ${data.companyDescription || 'None provided'}
      Company Type/Industry: ${data.companyType || 'Unknown'}
      Products: ${data.productName || 'None found'}
      
      Please provide:
      1. A more accurate company type/industry classification
      2. A categorization of the products found
      3. If the company appears to be focused on specific markets or demographics
      
      Format your response as JSON with keys: refinedCompanyType, productCategories, targetMarket
    `;
    
    // Call Azure OpenAI API
    const apiUrl = `${endpoint}openai/deployments/${deployment}/chat/completions?api-version=2023-05-15`;
    
    const requestBody = {
      messages: [
        {
          role: "system",
          content: "You are an AI assistant that specializes in analyzing company and product information to provide structured business intelligence data."
        },
        {
          role: "user",
          content: prompt
        }
      ],
      temperature: config.API.AZURE.OPENAI.TEMPERATURE,
      max_tokens: config.API.AZURE.OPENAI.MAX_TOKENS
    };
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'api-key': apiKey
      },
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseData = JSON.parse(response.getContentText());
    
    if (responseData.error) {
      throw new Error(responseData.error.message);
    }
    
    if (responseData.choices && responseData.choices.length > 0) {
      try {
        const aiResponse = responseData.choices[0].message.content;
        
        // Extract JSON from response text
        const jsonMatch = aiResponse.match(/\{[\s\S]*\}/);
        if (jsonMatch) {
          const enrichedData = JSON.parse(jsonMatch[0]);
          
          // Create result object with enriched data
          const result = {};
          
          if (enrichedData.refinedCompanyType && 
             (!data.companyType || data.companyType === "Other" || data.companyType === "Technology")) {
            result.companyType = enrichedData.refinedCompanyType;
          }
          
          if (enrichedData.productCategories && 
             (!data.productCategory || data.productCategory === "Other")) {
            result.productCategory = 
              Array.isArray(enrichedData.productCategories) 
                ? enrichedData.productCategories[0] 
                : enrichedData.productCategories;
          }
          
          if (enrichedData.targetMarket && !data.targetMarket) {
            result.targetMarket = enrichedData.targetMarket;
          }
          
          return result;
        }
      } catch (error) {
        console.error(`Error parsing AI response: ${error.message}`);
      }
    }
    
    return null;
  } catch (error) {
    console.error(`Error enriching data with AI: ${error.message}`);
    return null;
  }
}

/**
 * Query Google Knowledge Graph API
 * @param {string} query - The query (company name)
 * @param {Object} config - Configuration options
 * @return {Object} Knowledge Graph data
 */
function queryKnowledgeGraph(query, config) {
  try {
    // Skip if no API key
    if (!config.API.KNOWLEDGE_GRAPH.KEY) {
      return null;
    }
    
    const apiKey = config.API.KNOWLEDGE_GRAPH.KEY;
    const encodedQuery = encodeURIComponent(query);
    const apiUrl = `https://kgsearch.googleapis.com/v1/entities:search?query=${encodedQuery}&key=${apiKey}&limit=1&types=Organization&types=Corporation`;
    
    const options = {
      method: 'get',
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseData = JSON.parse(response.getContentText());
    
    if (responseData.itemListElement && responseData.itemListElement.length > 0) {
      const item = responseData.itemListElement[0].result;
      
      return {
        companyType: item.description || "",
        companyDescription: item.detailedDescription ? item.detailedDescription.articleBody : ""
      };
    }
    
    return null;
  } catch (error) {
    console.error(`Error querying Knowledge Graph: ${error.message}`);
    return null;
  }
}

/**
 * Update stats dashboard with extraction results
 * @param {string} url - URL that was processed
 * @param {Object} data - Extracted data
 */
function updateStatsForExtraction(url, data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const statsSheet = ss.getSheetByName("Data Center");
    
    if (!statsSheet) {
      console.error("Data Center sheet not found");
      return;
    }
    
    // Update Main Task Report stats
    const mainStatsRow = statsSheet.getRange(5, 1, 1, 14).getValues()[0];
    
    // URL (use the first one if not present)
    if (!mainStatsRow[0]) {
      statsSheet.getRange(5, 1).setValue(url);
    }
    
    // Processed count
    statsSheet.getRange(5, 2).setValue(mainStatsRow[1] + 1);
    
    // Remaining (no change in single extraction)
    
    // Success count
    statsSheet.getRange(5, 4).setValue(mainStatsRow[3] + 1);
    
    // Company Name Found
    if (data.companyName) {
      statsSheet.getRange(5, 6).setValue(mainStatsRow[5] + 1);
    }
    
    // Emails Found
    if (data.emails && data.emails.length > 0) {
      statsSheet.getRange(5, 7).setValue(mainStatsRow[6] + 1);
    }
    
    // Phones Found
    if (data.phones && data.phones.length > 0) {
      statsSheet.getRange(5, 8).setValue(mainStatsRow[7] + 1);
    }
    
    // Addresses Found
    if (data.addresses && data.addresses.length > 0) {
      statsSheet.getRange(5, 9).setValue(mainStatsRow[8] + 1);
    }
    
    // Company logo URL Found
    if (data.logo) {
      statsSheet.getRange(5, 10).setValue(mainStatsRow[9] + 1);
    }
    
    // Categories Found
    if (data.productCategory) {
      statsSheet.getRange(5, 11).setValue(mainStatsRow[10] + 1);
    }
    
    // Product Names Found
    if (data.productName) {
      statsSheet.getRange(5, 12).setValue(mainStatsRow[11] + 1);
    }
    
    // Product Image URLs Found
    if (data.images && data.images.length > 0) {
      statsSheet.getRange(5, 13).setValue(mainStatsRow[12] + 1);
    }
    
    // Product Type Found
    if (data.productCategory) {
      statsSheet.getRange(5, 14).setValue(mainStatsRow[13] + 1);
    }
    
    // Update API Report
    const apiStatsRow = statsSheet.getRange(9, 1, 1, 10).getValues()[0];
    
    // URL (use the first one if not present)
    if (!apiStatsRow[0]) {
      statsSheet.getRange(9, 1).setValue(url);
    }
    
    // Processed count
    statsSheet.getRange(9, 2).setValue(apiStatsRow[1] + 1);
    
    // Success count
    statsSheet.getRange(9, 4).setValue(apiStatsRow[3] + 1);
    
    // Azure OpenAI Success/Fail
    statsSheet.getRange(9, 6).setValue(globalStats.apiCalls.azure.success);
    statsSheet.getRange(9, 7).setValue(globalStats.apiCalls.azure.fail);
    
    // Knowledge Graph Success/Fail
    statsSheet.getRange(9, 8).setValue(globalStats.apiCalls.knowledgeGraph.success);
    statsSheet.getRange(9, 9).setValue(globalStats.apiCalls.knowledgeGraph.fail);
    
  } catch (error) {
    console.error(`Error updating stats: ${error.message}`);
  }
}

/**
 * Get recent extractions for the web app
 * @param {number} limit - Maximum number of results to return
 * @return {Array} Recent extractions
 */
function getRecentExtractions(limit = 10) {
  try {
    // Get the active spreadsheet and Data sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName("Data");
    
    if (!dataSheet) {
      return [];
    }
    
    // Get all data
    const dataRange = dataSheet.getDataRange();
    const dataValues = dataRange.getValues();
    
    if (dataValues.length <= 1) {
      // Only header row exists
      return [];
    }
    
    // Get headers
    const headers = dataValues[0];
    
    // Convert data to objects with header keys
    const dataObjects = [];
    
    for (let i = 1; i < dataValues.length; i++) {
      const rowData = {};
      
      for (let j = 0; j < headers.length; j++) {
        rowData[headers[j]] = dataValues[i][j];
      }
      
      // Add row index for reference
      rowData._rowIndex = i + 1;
      
      dataObjects.push(rowData);
    }
    
    // Sort by extraction date (descending) if available
    dataObjects.sort((a, b) => {
      // Use the dedicated Extraction Date column if available
      if (a['Extraction Date'] && b['Extraction Date']) {
        return new Date(b['Extraction Date']) - new Date(a['Extraction Date']);
      }
      // Otherwise use row order
      return b._rowIndex - a._rowIndex;
    });
    
    // Limit results
    return dataObjects.slice(0, limit);
    
  } catch (error) {
    console.error(`Error getting recent extractions: ${error.message}`);
    return [];
  }
}

/**
 * Save API keys to the script properties
 * @param {Object} keys - API keys object
 * @return {Object} Status result
 */
function saveApiKeys(keys) {
  try {
    // Update configuration
    if (keys.azureOpenaiKey) {
      CONFIG.API.AZURE.OPENAI.KEY = keys.azureOpenaiKey;
    }
    
    if (keys.azureOpenaiEndpoint) {
      CONFIG.API.AZURE.OPENAI.ENDPOINT = keys.azureOpenaiEndpoint;
    }
    
    if (keys.azureOpenaiDeployment) {
      CONFIG.API.AZURE.OPENAI.DEPLOYMENT = keys.azureOpenaiDeployment;
    }
    
    if (keys.knowledgeGraphApiKey) {
      CONFIG.API.KNOWLEDGE_GRAPH.KEY = keys.knowledgeGraphApiKey;
    }
    
    // Save to Properties service for persistence
    const scriptProperties = PropertiesService.getScriptProperties();
    
    scriptProperties.setProperties({
      'AZURE_OPENAI_KEY': CONFIG.API.AZURE.OPENAI.KEY,
      'AZURE_OPENAI_ENDPOINT': CONFIG.API.AZURE.OPENAI.ENDPOINT,
      'AZURE_OPENAI_DEPLOYMENT': CONFIG.API.AZURE.OPENAI.DEPLOYMENT,
      'KNOWLEDGE_GRAPH_API_KEY': CONFIG.API.KNOWLEDGE_GRAPH.KEY
    });
    
    return {
      success: true,
      message: "API keys saved successfully"
    };
  } catch (error) {
    console.error(`Error saving API keys: ${error.message}`);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * Load API keys from script properties
 * @return {Object} API keys
 */
function loadApiKeys() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const azureKey = scriptProperties.getProperty('AZURE_OPENAI_KEY') || "";
    const azureEndpoint = scriptProperties.getProperty('AZURE_OPENAI_ENDPOINT') || "https://fetcher.openai.azure.com/";
    const azureDeployment = scriptProperties.getProperty('AZURE_OPENAI_DEPLOYMENT') || "gpt-4";
    const kgKey = scriptProperties.getProperty('KNOWLEDGE_GRAPH_API_KEY') || "";
    
    // Update configuration
    CONFIG.API.AZURE.OPENAI.KEY = azureKey;
    CONFIG.API.AZURE.OPENAI.ENDPOINT = azureEndpoint;
    CONFIG.API.AZURE.OPENAI.DEPLOYMENT = azureDeployment;
    CONFIG.API.KNOWLEDGE_GRAPH.KEY = kgKey;
    
    return {
      azureOpenaiKey: azureKey,
      azureOpenaiEndpoint: azureEndpoint,
      azureOpenaiDeployment: azureDeployment,
      knowledgeGraphApiKey: kgKey
    };
  } catch (error) {
    console.error(`Error loading API keys: ${error.message}`);
    return {
      azureOpenaiKey: "",
      azureOpenaiEndpoint: "https://fetcher.openai.azure.com/",
      azureOpenaiDeployment: "gpt-4",
      knowledgeGraphApiKey: ""
    };
  }
}

/**
 * Add menu item to run the extraction
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Web Stryker R7")
    .addItem("Extract from URL", "showExtractDialog")
    .addItem("Process All Pending URLs", "runBatchExtraction")
    .addSeparator()
    .addItem("Open Web App", "openWebApp")
    .addSeparator()
    .addItem("Configure API Keys", "showApiKeysDialog")
    .addToUi();
  
  // Load API keys from script properties
  loadApiKeys();
}

/**
 * Run batch extraction for all pending URLs
 */
function runBatchExtraction() {
  const result = processAllPendingUrls();
  
  const ui = SpreadsheetApp.getUi();
  if (result.success) {
    ui.alert(
      "Batch Extraction Completed", 
      `Processed ${result.processed} URLs\nSuccessful: ${result.successful}\nFailed: ${result.failed}`, 
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      "Batch Extraction Failed", 
      result.error || "An unknown error occurred", 
      ui.ButtonSet.OK
    );
  }
}

/**
 * Show dialog to extract from a URL
 */
function showExtractDialog() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; margin: 20px; }
      label { display: block; margin-top: 15px; font-weight: bold; }
      input { width: 100%; padding: 8px; margin-top: 5px; }
      .button-container { margin-top: 20px; text-align: right; }
      button { padding: 8px 15px; background-color: #4285f4; color: white; border: none; border-radius: 4px; cursor: pointer; }
      button:hover { background-color: #3367d6; }
      .status { margin-top: 20px; display: none; }
      .spinner { border: 4px solid #f3f3f3; border-top: 4px solid #3498db; border-radius: 50%; width: 20px; height: 20px; animation: spin 2s linear infinite; display: inline-block; margin-right: 10px; }
      @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
    </style>
    <h2>Extract Data from URL</h2>
    <form id="extract-form">
      <label for="url">Website URL:</label>
      <input type="text" id="url" name="url" placeholder="https://example.com" required>
      
      <div class="button-container">
        <button type="submit">Extract</button>
      </div>
    </form>
    
    <div id="status" class="status">
      <div class="spinner"></div>
      <span id="status-message">Extracting data...</span>
    </div>
    
    <script>
      document.getElementById('extract-form').addEventListener('submit', function(e) {
        e.preventDefault();
        
        // Show status
        document.getElementById('status').style.display = 'block';
        
        const url = document.getElementById('url').value;
        const extractionId = 'dialog-' + Date.now().toString();
        
        google.script.run
          .withSuccessHandler(function(result) {
            if (result.success) {
              document.getElementById('status-message').textContent = 'Extraction completed successfully!';
              setTimeout(function() {
                google.script.host.close();
              }, 2000);
            } else {
              document.getElementById('status-message').textContent = 'Error: ' + result.error;
            }
          })
          .withFailureHandler(function(error) {
            document.getElementById('status-message').textContent = 'Error: ' + error.message;
          })
          .processUrlFromWebApp(url, extractionId);
      });
    </script>
  `)
    .setWidth(400)
    .setHeight(250)
    .setTitle("Extract from URL");
  
  SpreadsheetApp.getUi().showModalDialog(html, "Extract from URL");
}

/**
 * Open the web app in a dialog
 */
function openWebApp() {
  const ui = SpreadsheetApp.getUi();
  
  // Get the deployment web app URL
  const scriptId = ScriptApp.getScriptId();
  const webAppUrl = `https://script.google.com/macros/s/${scriptId}/exec`;
  
  const html = HtmlService.createHtmlOutput(`
    <script>
      window.open("${webAppUrl}", "_blank");
      setTimeout(function() {
        google.script.host.close();
      }, 100);
    </script>
  `)
    .setWidth(1)
    .setHeight(1);
  
  ui.showModalDialog(html, "Opening Web App...");
}

/**
 * Show dialog to configure API keys
 */
function showApiKeysDialog() {
  // Load current keys
  const keys = loadApiKeys();
  
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; margin: 20px; }
      label { display: block; margin-top: 15px; font-weight: bold; }
      input { width: 100%; padding: 8px; margin-top: 5px; }
      .button-container { margin-top: 20px; text-align: right; }
      button { padding: 8px 15px; background-color: #4285f4; color: white; border: none; border-radius: 4px; cursor: pointer; }
      button:hover { background-color: #3367d6; }
      .section-title { margin-top: 25px; border-bottom: 1px solid #eee; padding-bottom: 5px; }
    </style>
    <h2>API Keys Configuration</h2>
    
    <div class="section-title">
      <h3>Azure OpenAI Configuration</h3>
    </div>
    <form id="api-keys-form">
      <label for="azure-key">API Key:</label>
      <input type="password" id="azure-key" name="azure-key" placeholder="Enter your Azure OpenAI API Key" value="${keys.azureOpenaiKey || ''}">
      
      <label for="azure-endpoint">Endpoint URL:</label>
      <input type="text" id="azure-endpoint" name="azure-endpoint" placeholder="https://your-resource.openai.azure.com/" value="${keys.azureOpenaiEndpoint || 'https://fetcher.openai.azure.com/'}">
      
      <label for="azure-deployment">Deployment Name:</label>
      <input type="text" id="azure-deployment" name="azure-deployment" placeholder="gpt-4" value="${keys.azureOpenaiDeployment || 'gpt-4'}">
      
      <div class="section-title">
        <h3>Google Knowledge Graph API</h3>
      </div>
      
      <label for="kg-key">API Key:</label>
      <input type="text" id="kg-key" name="kg-key" placeholder="Enter your Google Knowledge Graph API Key" value="${keys.knowledgeGraphApiKey || ''}">
      
      <div class="button-container">
        <button type="submit">Save Configuration</button>
      </div>
    </form>
    
    <script>
      document.getElementById('api-keys-form').addEventListener('submit', function(e) {
        e.preventDefault();
        
        const azureOpenaiKey = document.getElementById('azure-key').value;
        const azureOpenaiEndpoint = document.getElementById('azure-endpoint').value;
        const azureOpenaiDeployment = document.getElementById('azure-deployment').value;
        const knowledgeGraphApiKey = document.getElementById('kg-key').value;
        
        google.script.run
          .withSuccessHandler(function(result) {
            if (result.success) {
              alert('Configuration saved successfully!');
              google.script.host.close();
            } else {
              alert('Error: ' + result.message);
            }
          })
          .withFailureHandler(function(error) {
            alert('Error: ' + error.message);
          })
          .saveApiKeys({
            azureOpenaiKey: azureOpenaiKey,
            azureOpenaiEndpoint: azureOpenaiEndpoint,
            azureOpenaiDeployment: azureOpenaiDeployment,
            knowledgeGraphApiKey: knowledgeGraphApiKey
          });
      });
    </script>
  `)
    .setWidth(500)
    .setHeight(550)
    .setTitle("API Keys Configuration");
  
  SpreadsheetApp.getUi().showModalDialog(html, "Configure API Keys");
}
/**
 * Determine product category from product name and description
 * @param {string} name - Product name
 * @param {string} description - Product description
 * @returns {string} Product category
 */
function determineProductCategory(name, description) {
  const textToAnalyze = (name + " " + (description || "")).toLowerCase();
  
  const categoryPatterns = [
    { pattern: /\b(?:tofu|vegan|vegetarian|plant-based|organic|food)\b/i, category: "Food" },
    { pattern: /\b(?:electronics|device|gadget|computer|laptop|phone|tablet)\b/i, category: "Electronics" },
    { pattern: /\b(?:clothing|shirt|pants|dress|jacket|apparel|wear)\b/i, category: "Clothing" },
    { pattern: /\b(?:furniture|chair|table|desk|sofa|bed)\b/i, category: "Furniture" },
    { pattern: /\b(?:book|novel|textbook|magazine|publication)\b/i, category: "Books" }
  ];
  
  for (const {pattern, category} of categoryPatterns) {
    if (pattern.test(textToAnalyze)) {
      return category;
    }
  }
  
  return "Other";
}

/**
 * Extract structured data (JSON-LD) from HTML content
 * @param {string} content - HTML content
 * @return {Object} Structured data object
 */
function extractStructuredData(content) {
  try {
    const jsonLdMatch = content.match(/<script type="application\/ld\+json">([\s\S]*?)<\/script>/i);
    if (jsonLdMatch) {
      return JSON.parse(jsonLdMatch[1]);
    }
    return {};
  } catch (e) {
    return {};
  }
}

/**
 * Clean HTML content to plain text
 * @param {string} html - HTML content
 * @return {string} Plain text
 */
function cleanHtml(html) {
  // Remove HTML tags
  let text = html.replace(/<[^>]+>/g, ' ');
  
  // Decode HTML entities
  text = text.replace(/&nbsp;/g, ' ');
  text = text.replace(/&amp;/g, '&');
  text = text.replace(/&lt;/g, '<');
  text = text.replace(/&gt;/g, '>');
  text = text.replace(/&quot;/g, '"');
  text = text.replace(/&apos;/g, "'");
  
  // Normalize whitespace
  text = text.replace(/\s+/g, ' ').trim();
  
  return text;
}

/**
 * Update stats for an extraction
 * @param {string} url - The URL that was extracted
 * @param {Object} data - The extracted data
 */
function updateStatsForExtraction(url, data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const statsSheet = ss.getSheetByName("Data Center");
    
    if (!statsSheet) {
      console.log("Stats sheet not found");
      return;
    }
    
    // Update main task report
    updateMainTaskReport(statsSheet, url, data);
    
    // Update API call report
    updateApiCallReport(statsSheet, url);
    
    // Update product categories report
    updateCategoriesReport(statsSheet, data);
    
  } catch (error) {
    console.error(`Error updating stats: ${error.message}`);
  }
}

/**
 * Add menu item to run the extraction
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Web Stryker R7")
    .addItem("Extract from URL", "showExtractDialog")
    .addItem("Process All Pending URLs", "processAllPendingUrls")
    .addItem("Open Web App", "openWebApp")
    .addSeparator()
    .addItem("Configure API Keys", "showApiKeysDialog")
    .addItem("View Data Center", "openDataCenter")
    .addToUi();
  
  // Load API keys from script properties
  loadApiKeys();
  
  // Initialize the spreadsheet if needed
  setupSpreadsheet();
}

/**
 * Open the Data Center sheet
 */
function openDataCenter() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const statsSheet = ss.getSheetByName("Data Center");
  
  if (statsSheet) {
    statsSheet.activate();
  } else {
    setupSpreadsheet();
    ss.getSheetByName("Data Center").activate();
  }
}
