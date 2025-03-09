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
