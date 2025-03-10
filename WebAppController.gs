/**
 * WebAppController - Controls the web app interface and interactions
 * Implements domain-driven design principles for cleaner architecture
 */
class WebAppController {
  constructor() {
    this.extractionService = new ExtractionService(CONFIG);
    this.sheetRepository = new SheetRepository();
  }
  
  /**
   * Initialize the web app - called when the page loads
   */
  initialize() {
    this.setupEventListeners();
    this.loadApiKeys();
    this.loadRecentExtractions();
  }
  
  /**
   * Set up all event listeners for the web app
   */
  setupEventListeners() {
    // URL form submission
    document.getElementById('url-form').addEventListener('submit', (e) => {
      e.preventDefault();
      this.processUrl();
    });
    
    // API settings toggle
    document.getElementById('api-settings-toggle').addEventListener('click', () => {
      this.toggleApiSettings();
    });
    
    // API keys form submission
    document.getElementById('api-keys-form').addEventListener('submit', (e) => {
      e.preventDefault();
      this.saveApiKeys();
    });
    
    // Control buttons
    document.getElementById('pause-btn').addEventListener('click', () => this.pauseExtraction());
    document.getElementById('resume-btn').addEventListener('click', () => this.resumeExtraction());
    document.getElementById('stop-btn').addEventListener('click', () => this.stopExtraction());
    document.getElementById('reset-btn').addEventListener('click', () => this.resetForm());
    
    // Tab change event for loading recent extractions
    document.getElementById('recent-tab').addEventListener('click', () => {
      this.loadRecentExtractions();
    });
  }
  
  /**
   * Process URL submission
   */
  processUrl() {
    const urlInput = document.getElementById('url-input');
    const url = urlInput.value.trim();
    
    if (!url) {
      this.showError('Please enter a valid URL');
      return;
    }
    
    // Hide error message and results
    document.getElementById('error-message').style.display = 'none';
    document.getElementById('results-container').style.display = 'none';
    
    // Show loading overlay
    this.showLoading();
    
    // Update button states
    this.updateButtonStates(true);
    
    // Generate unique extraction ID
    this.extractionInProgress = true;
    this.extractionPaused = false;
    this.currentExtractionId = 'webapp-' + Date.now().toString();
    
    // Start progress check interval
    this.startProgressChecking();
    
    // Call the server-side function
    google.script.run
      .withSuccessHandler((response) => this.handleExtractionSuccess(response))
      .withFailureHandler((error) => this.handleExtractionError(error))
      .processUrlFromWebApp(url, this.currentExtractionId);
  }
  
  /**
   * Show loading overlay with initial state
   */
  showLoading() {
    document.getElementById('loading-overlay').style.display = 'flex';
    document.getElementById('extraction-status').textContent = 'Extracting data...';
    document.getElementById('progress-bar').style.width = '0%';
    document.getElementById('progress-bar').setAttribute('aria-valuenow', '0');
    document.getElementById('progress-text').textContent = 'Starting extraction...';
  }
  
  /**
   * Start checking extraction progress
   */
  startProgressChecking() {
    if (this.progressInterval) {
      clearInterval(this.progressInterval);
    }
    
    this.progressInterval = setInterval(() => {
      if (this.currentExtractionId) {
        google.script.run
          .withSuccessHandler((progressData) => this.updateProgress(progressData))
          .getExtractionProgress(this.currentExtractionId);
      }
    }, 1000);
  }
  
  /**
   * Update progress display
   * @param {Object} progressData - Progress data
   */
  updateProgress(progressData) {
    if (!progressData.found) {
      return;
    }
    
    if (progressData.stopped) {
      clearInterval(this.progressInterval);
      document.getElementById('loading-overlay').style.display = 'none';
      this.showError('Extraction was stopped');
      this.updateButtonStates(false);
      return;
    }
    
    if (progressData.paused) {
      document.getElementById('extraction-status').textContent = 'Extraction paused';
      document.getElementById('progress-text').textContent = 'Waiting to resume...';
      return;
    }
    
    // Update progress bar
    const progressBar = document.getElementById('progress-bar');
    progressBar.style.width = progressData.progress + '%';
    progressBar.setAttribute('aria-valuenow', progressData.progress);
    
    // Update progress text
    document.getElementById('progress-text').textContent = progressData.stage || 'Processing...';
  }
  
  /**
   * Handle successful extraction
   * @param {Object} response - Extraction response
   */
  handleExtractionSuccess(response) {
    // Stop progress checking
    clearInterval(this.progressInterval);
    
    // Hide loading overlay
    document.getElementById('loading-overlay').style.display = 'none';
    
    // Reset extraction state
    this.updateButtonStates(false);
    
    if (response.success) {
      this.displayResults(response.data);
      
      // Refresh recent extractions if needed
      if (document.getElementById('recent-tab').classList.contains('active')) {
        this.loadRecentExtractions();
      }
    } else {
      this.showError(response.error || 'An unknown error occurred during extraction');
    }
  }
  
  /**
   * Handle extraction error
   * @param {Object} error - Error object
   */
  handleExtractionError(error) {
    // Stop progress checking
    clearInterval(this.progressInterval);
    
    // Hide loading overlay
    document.getElementById('loading-overlay').style.display = 'none';
    
    // Reset extraction state
    this.updateButtonStates(false);
    
    // Show error message
    this.showError(error.message || 'An error occurred during extraction');
  }
  
  /**
   * Display extraction results
   * @param {Object} data - Extracted data
   */
  displayResults(data) {
    // Company Information
    document.getElementById('company-name').textContent = data.companyName || 'Not found';
    document.getElementById('company-type').textContent = data.companyType || 'Not specified';
    document.getElementById('company-description').textContent = data.companyDescription || 'No description available';
    document.getElementById('company-emails').textContent = data.emails || 'No email addresses found';
    document.getElementById('company-phones').textContent = data.phones || 'No phone numbers found';
    document.getElementById('company-address').textContent = data.addresses || 'No address found';
    
    // Format extraction date
    let extractionDate = 'Unknown';
    try {
      if (data.extractionDate) {
        const date = new Date(data.extractionDate);
        extractionDate = date.toLocaleString();
      }
    } catch (e) {
      console.error('Error formatting date:', e);
    }
    document.getElementById('extraction-date').textContent = extractionDate;
    
    // Product Information
    document.getElementById('product-name').textContent = data.productName || 'No products found';
    
    // Product URL with link if available
    if (data.productUrl) {
      document.getElementById('product-url').innerHTML = `<a href="${data.productUrl}" target="_blank">${data.productUrl}</a>`;
    } else {
      document.getElementById('product-url').textContent = 'No product URL available';
    }
    
    document.getElementById('product-category').textContent = data.productCategory || 'Not categorized';
    document.getElementById('product-subcategory').textContent = data.productSubcategory || 'No subcategory';
    document.getElementById('product-quantity').textContent = data.quantity || 'Not specified';
    document.getElementById('product-price').textContent = data.price || 'No price information';
    document.getElementById('product-description').textContent = data.productDescription || 'No product description available';
    document.getElementById('product-specifications').textContent = data.specifications || 'No specifications available';
    
    // Display product images
    this.displayProductImages(data.images);
    
    // Show results container
    document.getElementById('results-container').style.display = 'block';
  }
  
  /**
   * Display product images
   * @param {string} imagesString - Comma-separated image URLs
   */
  displayProductImages(imagesString) {
    const imagesContainer = document.getElementById('images-container');
    imagesContainer.innerHTML = '';
    
    if (imagesString) {
      const imageUrls = imagesString.split(',').filter(url => url.trim());
      
      if (imageUrls.length > 0) {
        for (const url of imageUrls) {
          const trimmedUrl = url.trim();
          if (trimmedUrl) {
            const img = document.createElement('img');
            img.src = trimmedUrl;
            img.alt = 'Product image';
            img.className = 'image-preview';
            img.onerror = function() { this.style.display = 'none'; };
            imagesContainer.appendChild(img);
          }
        }
      } else {
        imagesContainer.textContent = 'No product images found';
      }
    } else {
      imagesContainer.textContent = 'No product images found';
    }
  }
  
  /**
   * Show error message
   * @param {string} message - Error message
   */
  showError(message) {
    const errorElement = document.getElementById('error-message');
    errorElement.textContent = message;
    errorElement.style.display = 'block';
  }
  
  /**
   * Update control button states
   * @param {boolean} isExtracting - Whether extraction is in progress
   */
  updateButtonStates(isExtracting) {
    this.extractionInProgress = isExtracting;
    
    document.getElementById('extract-btn').disabled = isExtracting;
    document.getElementById('pause-btn').disabled = !isExtracting || this.extractionPaused;
    document.getElementById('resume-btn').disabled = !isExtracting || !this.extractionPaused;
    document.getElementById('stop-btn').disabled = !isExtracting;
    document.getElementById('reset-btn').disabled = isExtracting;
  }
  
  /**
   * Pause extraction
   */
  pauseExtraction() {
    if (this.extractionInProgress && !this.extractionPaused && this.currentExtractionId) {
      this.extractionPaused = true;
      
      // Update button states
      this.updateButtonStates(true);
      
      // Call server-side function
      google.script.run
        .withSuccessHandler((response) => {
          console.log('Extraction paused:', response);
        })
        .pauseExtraction(this.currentExtractionId);
    }
  }
  
  /**
   * Resume extraction
   */
  resumeExtraction() {
    if (this.extractionInProgress && this.extractionPaused && this.currentExtractionId) {
      this.extractionPaused = false;
      
      // Update button states
      this.updateButtonStates(true);
      
      // Call server-side function
      google.script.run
        .withSuccessHandler((response) => {
          console.log('Extraction resumed:', response);
        })
        .resumeExtraction(this.currentExtractionId);
    }
  }
  
  /**
   * Stop extraction
   */
  stopExtraction() {
    if (this.extractionInProgress && this.currentExtractionId) {
      // Call server-side function
      google.script.run
        .withSuccessHandler((response) => {
          console.log('Extraction stopped:', response);
          
          // Reset state
          this.extractionInProgress = false;
          this.extractionPaused = false;
          
          // Update UI
          clearInterval(this.progressInterval);
          document.getElementById('loading-overlay').style.display = 'none';
          this.showError('Extraction was stopped by user');
          this.updateButtonStates(false);
        })
        .stopExtraction(this.currentExtractionId);
    }
  }
  
  /**
   * Reset form
   */
  resetForm() {
    document.getElementById('url-input').value = '';
    document.getElementById('error-message').style.display = 'none';
    document.getElementById('results-container').style.display = 'none';
    
    // Reset extraction state
    this.extractionInProgress = false;
    this.extractionPaused = false;
    this.currentExtractionId = null;
    
    this.updateButtonStates(false);
  }
  
  /**
   * Toggle API settings panel
   */
  toggleApiSettings() {
    const settingsCard = document.getElementById('api-settings-card');
    if (settingsCard.style.display === 'block') {
      settingsCard.style.display = 'none';
    } else {
      settingsCard.style.display = 'block';
    }
  }
  
  /**
   * Load API keys
   */
  loadApiKeys() {
    google.script.run
      .withSuccessHandler((keys) => {
        document.getElementById('azure-key').value = keys.azureOpenaiKey || '';
        document.getElementById('azure-endpoint').value = keys.azureOpenaiEndpoint || 'https://fetcher.openai.azure.com/';
        document.getElementById('azure-deployment').value = keys.azureOpenaiDeployment || 'gpt-4';
        document.getElementById('kg-key').value = keys.knowledgeGraphApiKey || '';
      })
      .loadApiKeys();
  }
  
  /**
   * Save API keys
   */
  saveApiKeys() {
    const keys = {
      azureOpenaiKey: document.getElementById('azure-key').value,
      azureOpenaiEndpoint: document.getElementById('azure-endpoint').value,
      azureOpenaiDeployment: document.getElementById('azure-deployment').value,
      knowledgeGraphApiKey: document.getElementById('kg-key').value
    };
    
    google.script.run
      .withSuccessHandler((result) => {
        if (result.success) {
          alert('API keys saved successfully');
          document.getElementById('api-settings-card').style.display = 'none';
        } else {
          alert('Error saving API keys: ' + result.message);
        }
      })
      .withFailureHandler((error) => {
        alert('Error: ' + error.message);
      })
      .saveApiKeys(keys);
  }
  
  /**
   * Load recent extractions
   */
  loadRecentExtractions() {
    document.getElementById('recent-loading').style.display = 'block';
    document.getElementById('recent-error').style.display = 'none';
    document.getElementById('no-extractions').style.display = 'none';
    document.getElementById('recent-container').innerHTML = '';
    
    google.script.run
      .withSuccessHandler((extractions) => this.displayRecentExtractions(extractions))
      .withFailureHandler((error) => {
        document.getElementById('recent-loading').style.display = 'none';
        document.getElementById('recent-error').style.display = 'block';
        console.error('Error loading recent extractions:', error);
      })
      .getRecentExtractions(12);
  }
  
  /**
   * Display recent extractions
   * @param {Array} extractions - Recent extractions
   */
  displayRecentExtractions(extractions) {
    document.getElementById('recent-loading').style.display = 'none';
    
    if (!extractions || extractions.length === 0) {
      document.getElementById('no-extractions').style.display = 'block';
      return;
    }
    
    const container = document.getElementById('recent-container');
    
    extractions.forEach((extraction) => {
      // Format date
      let extractionDate = 'Unknown';
      try {
        if (extraction['Extraction Date']) {
          const date = new Date(extraction['Extraction Date']);
          extractionDate = date.toLocaleDateString();
        }
      } catch (e) {
        console.error('Error formatting date:', e);
      }
      
      // Determine status class
      let statusClass = 'status-pending';
      if (extraction['Status'] === 'Completed') {
        statusClass = 'status-completed';
      } else if (extraction['Status'] === 'Failed') {
        statusClass = 'status-failed';
      } else if (extraction['Status'] === 'In Progress') {
        statusClass = 'status-progress';
      }
      
      // Create card
      const col = document.createElement('div');
      col.className = 'col-md-4 mb-4';
      col.innerHTML = `
        <div class="card result-card h-100">
          <div class="card-header d-flex justify-content-between align-items-center">
            <h6 class="mb-0 text-truncate" title="${extraction['Company Name'] || 'Unknown Company'}">
              <i class="fas fa-building btn-icon"></i>${extraction['Company Name'] || 'Unknown Company'}
            </h6>
            <span class="status-badge ${statusClass}">${extraction['Status'] || 'Unknown'}</span>
          </div>
          <div class="card-body">
            <p class="small mb-2 text-truncate" title="${extraction['URL'] || ''}">
              <i class="fas fa-link btn-icon"></i>${extraction['URL'] || 'No URL'}
            </p>
            <p class="small mb-2">
              <i class="fas fa-calendar-alt btn-icon"></i>${extractionDate}
            </p>
            <div class="small mb-2">
              <i class="fas fa-shopping-bag btn-icon"></i>
              <strong>Products:</strong> ${extraction['Product Name'] ? extraction['Product Name'].substring(0, 50) + (extraction['Product Name'].length > 50 ? '...' : '') : 'None found'}
            </div>
            <div class="small mb-2">
              <i class="fas fa-envelope btn-icon"></i>
              <strong>Contacts:</strong> ${extraction['Emails'] ? 'Found' : 'None found'}
            </div>
          </div>
          <div class="card-footer text-center">
            <button class="btn btn-sm btn-primary" onclick="webAppController.reExtractUrl('${extraction['URL']}')">
              <i class="fas fa-sync-alt btn-icon"></i>Extract Again
            </button>
          </div>
        </div>
      `;
      
      container.appendChild(col);
    });
  }
  
  /**
   * Re-extract URL from recent extractions
   * @param {string} url - URL to extract
   */
  reExtractUrl(url) {
    if (!url) return;
    
    // Switch to extract tab
    document.getElementById('extract-tab').click();
    
    // Fill URL input
    document.getElementById('url-input').value = url;
    
    // Start extraction
    this.processUrl();
  }
}
