/**
 * MetadataExtractor - Extracts Open Graph, Twitter Card, and other metadata
 */
class MetadataExtractor {
  /**
   * Extract metadata from HTML content
   * @param {string} content - HTML content
   * @param {string} url - Source URL
   * @param {CompanyEntity} company - Company entity to populate
   */
  extract(content, url, company) {
    try {
      company.metadata = {
        openGraph: {},
        twitterCard: {},
        schemaOrg: {},
        metaTags: {}
      };
      
      // Extract Open Graph metadata
      const ogTitleMatch = content.match(/<meta\s+property="og:title"\s+content="([^"]*)"/i);
      const ogDescMatch = content.match(/<meta\s+property="og:description"\s+content="([^"]*)"/i);
      const ogImageMatch = content.match(/<meta\s+property="og:image"\s+content="([^"]*)"/i);
      const ogUrlMatch = content.match(/<meta\s+property="og:url"\s+content="([^"]*)"/i);
      const ogTypeMatch = content.match(/<meta\s+property="og:type"\s+content="([^"]*)"/i);
      const ogSiteNameMatch = content.match(/<meta\s+property="og:site_name"\s+content="([^"]*)"/i);
      
      if (ogTitleMatch) company.metadata.openGraph.title = ogTitleMatch[1];
      if (ogDescMatch) company.metadata.openGraph.description = ogDescMatch[1];
      if (ogImageMatch) company.metadata.openGraph.image = ogImageMatch[1];
      if (ogUrlMatch) company.metadata.openGraph.url = ogUrlMatch[1];
      if (ogTypeMatch) company.metadata.openGraph.type = ogTypeMatch[1];
      if (ogSiteNameMatch) company.metadata.openGraph.siteName = ogSiteNameMatch[1];
      
      // Extract Twitter Card metadata
      const twitterTitleMatch = content.match(/<meta\s+name="twitter:title"\s+content="([^"]*)"/i);
      const twitterDescMatch = content.match(/<meta\s+name="twitter:description"\s+content="([^"]*)"/i);
      const twitterImageMatch = content.match(/<meta\s+name="twitter:image"\s+content="([^"]*)"/i);
      const twitterCardMatch = content.match(/<meta\s+name="twitter:card"\s+content="([^"]*)"/i);
      const twitterSiteMatch = content.match(/<meta\s+name="twitter:site"\s+content="([^"]*)"/i);
      
      if (twitterTitleMatch) company.metadata.twitterCard.title = twitterTitleMatch[1];
      if (twitterDescMatch) company.metadata.twitterCard.description = twitterDescMatch[1];
      if (twitterImageMatch) company.metadata.twitterCard.image = twitterImageMatch[1];
      if (twitterCardMatch) company.metadata.twitterCard.card = twitterCardMatch[1];
      if (twitterSiteMatch) company.metadata.twitterCard.site = twitterSiteMatch[1];
      
      // Extract other meta tags
      const descriptionMatch = content.match(/<meta\s+name="description"\s+content="([^"]*)"/i);
      const keywordsMatch = content.match(/<meta\s+name="keywords"\s+content="([^"]*)"/i);
      const authorMatch = content.match(/<meta\s+name="author"\s+content="([^"]*)"/i);
      const viewportMatch = content.match(/<meta\s+name="viewport"\s+content="([^"]*)"/i);
      
      if (descriptionMatch) company.metadata.metaTags.description = descriptionMatch[1];
      if (keywordsMatch) company.metadata.metaTags.keywords = keywordsMatch[1];
      if (authorMatch) company.metadata.metaTags.author = authorMatch[1];
      if (viewportMatch) company.metadata.metaTags.viewport = viewportMatch[1];
      
      // Extract canonical URL
      const canonicalMatch = content.match(/<link\s+rel="canonical"\s+href="([^"]*)"/i);
      if (canonicalMatch) company.metadata.canonical = canonicalMatch[1];
      
      // Use Open Graph or Twitter data to enhance company information if missing
      if (!company.companyName && company.metadata.openGraph.siteName) {
        company.companyName = company.metadata.openGraph.siteName;
      }
      
      if (!company.companyDescription && company.metadata.openGraph.description) {
        company.companyDescription = company.metadata.openGraph.description;
      } else if (!company.companyDescription && company.metadata.metaTags.description) {
        company.companyDescription = company.metadata.metaTags.description;
      }
      
      if (!company.logo && company.metadata.openGraph.image) {
        company.logo = company.metadata.openGraph.image;
      }
    } catch (e) {
      console.error(`Error extracting metadata: ${e.message}`);
    }
  }
}

/**
 * PageCrawler - Crawls website pages to discover content
 */
class PageCrawler {
  /**
   * @param {Object} config - Configuration options
   */
  constructor(config) {
    this.config = config;
    this.maxDepth = config.MAX_CRAWL_DEPTH || 2;
    this.maxPages = config.MAX_PAGES_TO_CRAWL || 10;
    this.visitedUrls = new Set();
  }
  
  /**
   * Crawl website to discover pages
   * @param {string} baseUrl - Starting URL
   * @param {string} initialContent - HTML content of the base URL
   * @param {string} extractionId - Extraction ID for tracking
   * @return {Object} Map of URL to page content
   */
  async crawl(baseUrl, initialContent, extractionId) {
    const pageContents = new Map();
    pageContents.set(baseUrl, initialContent);
    this.visitedUrls.add(baseUrl);
    
    // Get important URLs to crawl
    const urlsToVisit = this.extractPriorityUrls(initialContent, baseUrl);
    
    // Breadth-first crawling
    let currentDepth = 1;
    let currentUrlSet = new Set(urlsToVisit);
    let nextUrlSet = new Set();
    
    while (currentDepth <= this.maxDepth && this.visitedUrls.size < this.maxPages) {
      for (const url of currentUrlSet) {
        // Skip if already visited
        if (this.visitedUrls.has(url)) continue;
        
        // Check if extraction is stopped
        if (ExtractionState.isStopped(extractionId)) {
          console.log("Crawling stopped due to extraction cancellation");
          return pageContents;
        }
        
        // Wait if paused
        ExtractionState.waitIfPaused(extractionId);
        
        try {
          // Fetch the page
          const response = await this.fetchWithDelay(url);
          
          if (response) {
            const content = response.getContentText();
            pageContents.set(url, content);
            this.visitedUrls.add(url);
            
            // Extract new URLs for next level
            if (currentDepth < this.maxDepth) {
              const newUrls = this.extractUrls(content, url, baseUrl);
              for (const newUrl of newUrls) {
                if (!this.visitedUrls.has(newUrl)) {
                  nextUrlSet.add(newUrl);
                }
              }
            }
          }
          
          // Stop if we've reached the max pages
          if (this.visitedUrls.size >= this.maxPages) {
            break;
          }
        } catch (error) {
          console.error(`Error crawling ${url}: ${error.message}`);
        }
      }
      
      // Move to next depth
      currentDepth++;
      currentUrlSet = nextUrlSet;
      nextUrlSet = new Set();
    }
    
    return pageContents;
  }
  
  /**
   * Extract priority URLs for crawling
   * @param {string} content - HTML content
   * @param {string} baseUrl - Base URL
   * @return {Array} URLs to visit
   */
  extractPriorityUrls(content, baseUrl) {
    const urls = new Set();
    
    // Important sections to check
    const importantPaths = [
      '/products', '/about', '/about-us', '/company', '/contact', 
      '/catalog', '/shop', '/store'
    ];
    
    // Extract all links
    const linkPattern = /<a\s+[^>]*href="([^"]*)"[^>]*>/gi;
    let match;
    
    while ((match = linkPattern.exec(content)) !== null) {
      let url = match[1].trim();
      
      // Skip empty, anchor, javascript links
      if (!url || url === '#' || url.startsWith('javascript:') || url.startsWith('mailto:')) {
        continue;
      }
      
      // Resolve to absolute URL
      url = this.resolveUrl(url, baseUrl);
      
      // Check if same domain
      if (this.isSameDomain(url, baseUrl)) {
        // Check if this is a priority path
        const urlObj = new URL(url);
        const pathname = urlObj.pathname.toLowerCase();
        
        if (importantPaths.some(path => pathname === path || pathname === path + '/')) {
          urls.add(url);
        } else if (pathname.includes('/product') || pathname.includes('/category')) {
          urls.add(url);
        }
      }
    }
    
    return Array.from(urls);
  }
  
  /**
   * Extract all URLs from HTML content
   * @param {string} content - HTML content
   * @param {string} pageUrl - URL of the current page
   * @param {string} baseUrl - Original base URL
   * @return {Array} Discovered URLs
   */
  extractUrls(content, pageUrl, baseUrl) {
    const urls = new Set();
    const linkPattern = /<a\s+[^>]*href="([^"]*)"[^>]*>/gi;
    let match;
    
    while ((match = linkPattern.exec(content)) !== null) {
      let url = match[1].trim();
      
      // Skip empty, anchor, javascript links
      if (!url || url === '#' || url.startsWith('javascript:') || url.startsWith('mailto:')) {
        continue;
      }
      
      // Resolve to absolute URL
      url = this.resolveUrl(url, pageUrl);
      
      // Only include URLs from the same domain
      if (this.isSameDomain(url, baseUrl) && !this.isExcludedPath(url)) {
        urls.add(url);
      }
    }
    
    return Array.from(urls);
  }
  
  /**
   * Fetch URL with delay
   * @param {string} url - URL to fetch
   * @return {HTTPResponse} Response or null
   */
  async fetchWithDelay(url) {
    try {
      // Add delay to avoid overloading the server
      Utilities.sleep(1000 + Math.random() * 1000);
      
      const options = {
        muteHttpExceptions: true,
        followRedirects: true,
        headers: {
          "User-Agent": this.config.USER_AGENT || "Mozilla/5.0 (compatible; GoogleAppsScript/1.0)"
        }
      };
      
      return UrlFetchApp.fetch(url, options);
    } catch (error) {
      console.error(`Error fetching ${url}: ${error.message}`);
      return null;
    }
  }
  
  /**
   * Check if URL is excluded
   * @param {string} url - URL to check
   * @return {boolean} Whether the URL should be excluded
   */
  isExcludedPath(url) {
    try {
      const excludedPaths = [
        'login', 'register', 'account', 'cart', 'checkout', 'privacy',
        'terms', 'blog', 'news', 'article', 'post', 'search'
      ];
      
      const urlObj = new URL(url);
      const path = urlObj.pathname.toLowerCase();
      
      return excludedPaths.some(excluded => 
        path.includes(`/${excluded}`) || path.includes(`/${excluded}/`)
      );
    } catch (e) {
      return false;
    }
  }
  
  /**
   * Check if two URLs are from the same domain
   * @param {string} url1 - First URL
   * @param {string} url2 - Second URL
   * @return {boolean} Whether URLs are from the same domain
   */
  isSameDomain(url1, url2) {
    try {
      const domain1 = new URL(url1).hostname;
      const domain2 = new URL(url2).hostname;
      return domain1 === domain2;
    } catch (e) {
      return false;
    }
  }
  
  /**
   * Resolve relative URL to absolute URL
   * @param {string} url - Relative or absolute URL
   * @param {string} base - Base URL
   * @return {string} Absolute URL
   */
  resolveUrl(url, base) {
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
      
      // Use URL API for proper resolution
      return new URL(url, base).href;
    } catch (e) {
      console.error(`Error resolving URL: ${e.message}`);
      return url;
    }
  }
}

/**
 * EnhancedProductExtractor - Extracts detailed product information
 */
class EnhancedProductExtractor {
  /**
   * @param {Object} config - Configuration options
   */
  constructor(config) {
    this.config = config;
    this.cloudApiService = new GoogleCloudApiService(config);
  }
  
  /**
   * Extract product information with crawling
   * @param {string} mainContent - Main page HTML content
   * @param {Map} additionalPages - Map of URL to page content from crawling
   * @param {string} baseUrl - Original base URL
   * @param {CompanyEntity} company - Company entity to populate
   * @param {string} extractionId - Extraction ID for tracking
   */
  async extractWithCrawling(mainContent, additionalPages, baseUrl, company, extractionId) {
    try {
      // First get basic product links from main page
      const productLinks = this.extractProductLinks(mainContent, baseUrl);
      
      // Identify product pages from crawled content
      const productPages = new Map();
      
      // Add any product URLs from main page
      for (const link of productLinks) {
        if (additionalPages.has(link.url)) {
          productPages.set(link.url, {
            content: additionalPages.get(link.url),
            text: link.text
          });
        }
      }
      
      // Look for product patterns in other crawled pages
      for (const [url, content] of additionalPages.entries()) {
        if (!productPages.has(url) && this.isLikelyProductPage(content, url)) {
          productPages.set(url, {
            content: content,
            text: this.extractTitleFromContent(content)
          });
        }
      }
      
      // Extract product categories from main navigation
      const categories = this.extractCategories(mainContent);
      
      // Process each product page
      for (const [url, pageData] of productPages.entries()) {
        // Check if extraction is stopped
        if (ExtractionState.isStopped(extractionId)) {
          console.log("Product extraction stopped");
          break;
        }
        
        // Wait if paused
        ExtractionState.waitIfPaused(extractionId);
        
        // Extract product details
        const productEntity = await this.extractDetailedProduct(
          pageData.content, 
          url, 
          pageData.text,
          categories
        );
        
        if (productEntity && productEntity.isValid()) {
          company.products.push(productEntity);
        }
      }
      
      // If no products found, try to extract from main page
      if (company.products.length === 0) {
        const productEntity = await this.extractDetailedProduct(mainContent, baseUrl, "", categories);
        
        if (productEntity && productEntity.isValid()) {
          company.products.push(productEntity);
        }
      }
      
      // Enhance with AI for categorization
      await this.enhanceProductCategories(company.products);
    } catch (e) {
      console.error(`Error in enhanced product extraction: ${e.message}`);
    }
  }
  
  /**
   * Extract detailed product information
   * @param {string} content - HTML content
   * @param {string} url - Product URL
   * @param {string} linkText - Text from link (optional)
   * @param {Array} categories - Available categories
   * @return {ProductEntity} Product entity
   */
  async extractDetailedProduct(content, url, linkText, categories) {
    const product = new ProductEntity();
    product.productUrl = url;
    
    try {
      // Extract basic product details
      this.extractBasicProductDetails(content, url, linkText, product);
      
      // Extract detailed specifications
      this.extractSpecifications(content, product);
      
      // Extract nutritional information
      this.extractNutritionalInfo(content, product);
      
      // Extract certifications and compliance
      this.extractCertifications(content, product);
      
      // Extract pricing and availability
      this.extractPricingInfo(content, product);
      
      // Set category information if available
      if (categories && categories.length > 0) {
        product.mainCategory = categories[0] || "";
        if (categories.length > 1) {
          product.subCategory = categories[1] || "";
        }
        if (categories.length > 2) {
          product.productFamily = categories[2] || "";
        }
      }
      
      // Analyze product images
      await this.analyzeProductImages(content, url, product);
      
      // Enhance with AI analysis
      if (this.config.API?.GOOGLE?.VERTEX_AI?.KEY && 
          (product.productName || product.description)) {
        try {
          const productAnalysis = await this.cloudApiService.analyzeProduct(
            product.productName,
            product.description
          );
          
          if (productAnalysis) {
            // Update categories if not already set
            if (productAnalysis.categories && productAnalysis.categories.length > 0) {
              if (!product.mainCategory && productAnalysis.categories.length > 0) {
                product.mainCategory = productAnalysis.categories[0];
              }
              
              if (!product.subCategory && productAnalysis.categories.length > 1) {
                product.subCategory = productAnalysis.categories[1];
              }
            }
            
            // Add attributes
            if (productAnalysis.attributes) {
              product.attributes = productAnalysis.attributes;
            }
          }
        } catch (error) {
          console.error(`Error analyzing product with AI: ${error.message}`);
        }
      }
      
      return product;
    } catch (error) {
      console.error(`Error extracting detailed product: ${error.message}`);
      return product;
    }
  }
  
  /**
   * Extract basic product details
   * @param {string} content - HTML content
   * @param {string} url - Product URL
   * @param {string} linkText - Text from link (optional)
   * @param {ProductEntity} product - Product entity to populate
   */
  extractBasicProductDetails(content, url, linkText, product) {
    // Extract product name (prioritize structured data)
    const structuredData = this.extractStructuredData(content);
    if (structuredData && structuredData.product && structuredData.product.name) {
      product.productName = structuredData.product.name;
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
          product.productName = this.cleanHtml(match[1]).trim();
          break;
        }
      }
      
      // If still no name, use link text or page title as fallback
      if (!product.productName) {
        if (linkText) {
          product.productName = linkText;
        } else {
          const titleMatch = content.match(/<title>(.*?)<\/title>/i);
          if (titleMatch) {
            product.productName = titleMatch[1].trim().split('|')[0].split('-')[0].trim();
          }
        }
      }
    }
    
    // Extract product price
    if (structuredData && structuredData.product && structuredData.product.offers && 
        structuredData.product.offers.price) {
      product.price = structuredData.product.offers.price;
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
          product.price = this.cleanHtml(match[1] || match[0]).trim();
          break;
        }
      }
    }
    
    // Extract product description
    if (structuredData && structuredData.product && structuredData.product.description) {
      product.description = structuredData.product.description;
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
          product.description = this.cleanHtml(match[1] || match[0]).trim();
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
        product.quantity = this.cleanHtml(match[1] || match[0]).trim();
        break;
      }
    }
    
    // Extract product images
    this.extractProductImages(content, url, product);
  }
  
  /**
   * Extract product specifications
   * @param {string} content - HTML content
   * @param {ProductEntity} product - Product entity to populate
   */
  extractSpecifications(content, url, product) {
    // Setup detailed specifications if not already there
    if (!product.detailedSpecs) {
      product.detailedSpecs = {
        measurements: {},
        packaging: {},
        storage: {},
        preparation: {},
        other: {}
      };
    }
    
    // Look for specification tables/sections
    const specSectionPatterns = [
      /<(?:div|section|table)[^>]*\b(?:id|class)="[^"]*\b(?:specification|technical|details|specs)[^"]*"[^>]*>([\s\S]*?)<\/(?:div|section|table)>/i,
      /<(?:div|section)[^>]*\bitemprop="additionalProperty"[^>]*>([\s\S]*?)<\/(?:div|section)>/i,
      /<h\d[^>]*>\s*(?:Specifications|Technical Details|Tech Specs|Additional Information)\s*<\/h\d>([\s\S]*?)(?:<h\d|<\/div|<\/section)/i,
      /<table[^>]*\b(?:id|class)="[^"]*\b(?:product|item)[^"]*-attributes"[^>]*>([\s\S]*?)<\/table>/i
    ];
    
    let specSection = "";
    for (const pattern of specSectionPatterns) {
      const match = content.match(pattern);
      if (match) {
        specSection = match[1];
        break;
      }
    }
    
    if (specSection) {
      // Look for table rows
      const tableRows = [...specSection.matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/gi)];
      
      // Process table rows
      for (const row of tableRows) {
        const cells = [...row[1].matchAll(/<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi)];
        
        if (cells.length >= 2) {
          const label = this.cleanHtml(cells[0][1]).trim().toLowerCase();
          const value = this.cleanHtml(cells[1][1]).trim();
          
          if (label && value) {
            this.categorizeSpecification(label, value, product);
          }
        }
      }
      
      // If no table rows found, look for definition lists
      if (tableRows.length === 0) {
        const dlItems = [...specSection.matchAll(/<dt[^>]*>([\s\S]*?)<\/dt>\s*<dd[^>]*>([\s\S]*?)<\/dd>/gi)];
        
        for (const item of dlItems) {
          const label = this.cleanHtml(item[1]).trim().toLowerCase();
          const value = this.cleanHtml(item[2]).trim();
          
          if (label && value) {
            this.categorizeSpecification(label, value, product);
          }
        }
      }
      
      // Generate specifications string from collected details
      this.generateSpecificationsString(product);
    }
  }
  
  /**
   * Categorize product specification into appropriate category
   * @param {string} label - Specification label
   * @param {string} value - Specification value
   * @param {ProductEntity} product - Product entity to update
   */
  categorizeSpecification(label, value, product) {
    // Measurements keywords
    const measurementKeywords = [
      'weight', 'height', 'width', 'length', 'depth', 'dimension', 'size',
      'volume', 'capacity', 'diameter', 'thickness', 'density'
    ];
    
    // Packaging keywords
    const packagingKeywords = [
      'package', 'packaging', 'pack', 'box', 'container', 'wrapper',
      'case', 'pallet', 'quantity', 'count'
    ];
    
    // Storage keywords
    const storageKeywords = [
      'storage', 'shelf life', 'store', 'refrigerate', 'freeze', 'temperature',
      'humidity', 'expiration', 'expiry', 'keep'
    ];
    
    // Preparation keywords
    const preparationKeywords = [
      'preparation', 'prepare', 'cook', 'cooking', 'instructions', 'heat',
      'preheat', 'bake', 'microwave', 'serve', 'serving', 'use'
    ];
    
    // Categorize based on label
    if (measurementKeywords.some(keyword => label.includes(keyword))) {
      product.detailedSpecs.measurements[label] = value;
    } else if (packagingKeywords.some(keyword => label.includes(keyword))) {
      product.detailedSpecs.packaging[label] = value;
    } else if (storageKeywords.some(keyword => label.includes(keyword))) {
      product.detailedSpecs.storage[label] = value;
    } else if (preparationKeywords.some(keyword => label.includes(keyword))) {
      product.detailedSpecs.preparation[label] = value;
    } else {
      product.detailedSpecs.other[label] = value;
    }
  }
  
  /**
   * Generate specifications string from detailed specs
   * @param {ProductEntity} product - Product entity to update
   */
  generateSpecificationsString(product) {
    const specs = [];
    
    // Add measurements
    for (const [key, value] of Object.entries(product.detailedSpecs.measurements)) {
      specs.push(`${this.capitalizeFirstLetter(key)}: ${value}`);
    }
    
    // Add packaging details
    for (const [key, value] of Object.entries(product.detailedSpecs.packaging)) {
      specs.push(`${this.capitalizeFirstLetter(key)}: ${value}`);
    }
    
    // Add storage requirements
    for (const [key, value] of Object.entries(product.detailedSpecs.storage)) {
      specs.push(`${this.capitalizeFirstLetter(key)}: ${value}`);
    }
    
    // Add preparation instructions
    for (const [key, value] of Object.entries(product.detailedSpecs.preparation)) {
      specs.push(`${this.capitalizeFirstLetter(key)}: ${value}`);
    }
    
    // Add other specifications
    for (const [key, value] of Object.entries(product.detailedSpecs.other)) {
      specs.push(`${this.capitalizeFirstLetter(key)}: ${value}`);
    }
    
    // Set specifications string
    if (specs.length > 0) {
      product.specifications = specs.join("\n");
    }
  }
  
  /**
   * Extract nutritional information
   * @param {string} content - HTML content
   * @param {ProductEntity} product - Product entity to populate
   */
  extractNutritionalInfo(content, product) {
    // Setup nutritional info if not already there
    if (!product.nutritionalInfo) {
      product.nutritionalInfo = {
        nutritionalValues: {},
        dietary: [],
        allergens: [],
        healthClaims: []
      };
    }
    
    // Look for nutrition tables/sections
    const nutritionSectionPatterns = [
      /<(?:div|section|table)[^>]*\b(?:id|class)="[^"]*\b(?:nutrition|nutritional|nutrition-facts)[^"]*"[^>]*>([\s\S]*?)<\/(?:div|section|table)>/i,
      /<h\d[^>]*>\s*(?:Nutrition|Nutritional Information|Nutrition Facts)\s*<\/h\d>([\s\S]*?)(?:<h\d|<\/div|<\/section)/i
    ];
    
    let nutritionSection = "";
    for (const pattern of nutritionSectionPatterns) {
      const match = content.match(pattern);
      if (match) {
        nutritionSection = match[1];
        break;
      }
    }
    
    if (nutritionSection) {
      // Extract nutritional values from tables
      const tableRows = [...nutritionSection.matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/gi)];
      
      for (const row of tableRows) {
        const cells = [...row[1].matchAll(/<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi)];
        
        if (cells.length >= 2) {
          const nutrient = this.cleanHtml(cells[0][1]).trim().toLowerCase();
          const value = this.cleanHtml(cells[1][1]).trim();
          
          if (nutrient && value) {
            product.nutritionalInfo.nutritionalValues[nutrient] = value;
          }
        }
      }
    }
    
    // Extract allergens
    const allergenPatterns = [
      /<(?:div|section|p)[^>]*\b(?:id|class)="[^"]*\b(?:allergen|allergy)[^"]*"[^>]*>([\s\S]*?)<\/(?:div|section|p)>/i,
      /(?:allergen|contains):\s*([^<.]+)/i,
      /(?:may contain|may contain traces of):\s*([^<.]+)/i
    ];
    
    for (const pattern of allergenPatterns) {
      const match = content.match(pattern);
      if (match) {
        const allergenText = this.cleanHtml(match[1]).trim();
        const allergens = allergenText.split(/,|\band\b/).map(a => a.trim()).filter(Boolean);
        product.nutritionalInfo.allergens.push(...allergens);
        break;
      }
    }
    
    // Extract dietary information
    const dietaryPatterns = [
      /<(?:div|section|p|span)[^>]*\b(?:id|class)="[^"]*\b(?:dietary|diet|vegan|vegetarian|gluten-free)[^"]*"[^>]*>([\s\S]*?)<\/(?:div|section|p|span)>/i,
      /(?:suitable for|diet|dietary):\s*([^<.]+)/i
    ];
    
    for (const pattern of dietaryPatterns) {
      const match = content.match(pattern);
      if (match) {
        const dietaryText = this.cleanHtml(match[1]).trim();
        const dietary = dietaryText.split(/,|\band\b/).map(d => d.trim()).filter(Boolean);
        product.nutritionalInfo.dietary.push(...dietary);
        break;
      }
    }
    
    // Look for common dietary indicators
    const dietaryIndicators = [
      { pattern: /\b(?:vegan|plant.?based)\b/i, value: "Vegan" },
      { pattern: /\b(?:vegetarian)\b/i, value: "Vegetarian" },
      { pattern: /\b(?:gluten.?free)\b/i, value: "Gluten-Free" },
      { pattern: /\b(?:dairy.?free|lactose.?free)\b/i, value: "Dairy-Free" },
      { pattern: /\b(?:nut.?free)\b/i, value: "Nut-Free" },
      { pattern: /\b(?:organic)\b/i, value: "Organic" }
    ];
    
    for (const indicator of dietaryIndicators) {
      if (indicator.pattern.test(content) && 
          !product.nutritionalInfo.dietary.includes(indicator.value)) {
        product.nutritionalInfo.dietary.push(indicator.value);
      }
    }
    
    // Remove duplicates
    product.nutritionalInfo.allergens = [...new Set(product.nutritionalInfo.allergens)];
    product.nutritionalInfo.dietary = [...new Set(product.nutritionalInfo.dietary)];
  }
  
  /**
   * Extract certifications and compliance information
   * @param {string} content - HTML content
   * @param {ProductEntity} product - Product entity to populate
   */
  extractCertifications(content, product) {
    // Setup quality assessment if not already there
    if (!product.qualityAssessment) {
      product.qualityAssessment = {
        certifications: [],
        standards: [],
        qualityIndicators: [],
        safetyMeasures: []
      };
    }
    
    // Common certification patterns
    const certificationPatterns = [
      { pattern: /\biso\s*9001\b/i, value: "ISO 9001" },
      { pattern: /\biso\s*14001\b/i, value: "ISO 14001" },
      { pattern: /\biso\s*22000\b/i, value: "ISO 22000" },
      { pattern: /\bhaccp\b/i, value: "HACCP" },
      { pattern: /\bfda\b/i, value: "FDA Approved" },
      { pattern: /\busda\b/i, value: "USDA Certified" },
      { pattern: /\busda\s*organic\b/i, value: "USDA Organic" },
      { pattern: /\bfair\s*trade\b/i, value: "Fair Trade Certified" },
      { pattern: /\bnon.gmo\b/i, value: "Non-GMO" },
      { pattern: /\bce\s*mark\b/i, value: "CE Mark" },
      { pattern: /\becolabel\b/i, value: "EU Ecolabel" },
      { pattern: /\benergy\s*star\b/i, value: "Energy Star" },
      { pattern: /\bgreen\s*seal\b/i, value: "Green Seal" },
      { pattern: /\bcradle\s*to\s*cradle\b/i, value: "Cradle to Cradle" },
      { pattern: /\bfsc\b/i, value: "FSC Certified" },
      { pattern: /\bmsds\b/i, value: "MSDS Available" }
    ];
    
    // Check for certifications in content
    for (const cert of certificationPatterns) {
      if (cert.pattern.test(content)) {
        product.qualityAssessment.certifications.push(cert.value);
      }
    }
    
    // Look for certification section
    const certSectionPatterns = [
      /<(?:div|section|ul)[^>]*\b(?:id|class)="[^"]*\b(?:certification|certificate|quality|standard)[^"]*"[^>]*>([\s\S]*?)<\/(?:div|section|ul)>/i,
      /<h\d[^>]*>\s*(?:Certifications|Quality|Standards|Compliance)\s*<\/h\d>([\s\S]*?)(?:<h\d|<\/div|<\/section)/i
    ];
    
    let certSection = "";
    for (const pattern of certSectionPatterns) {
      const match = content.match(pattern);
      if (match) {
        certSection = match[1];
        break;
      }
    }
    
    if (certSection) {
      // Extract list items
      const listItems = [...certSection.matchAll(/<li[^>]*>([\s\S]*?)<\/li>/gi)];
      
      for (const item of listItems) {
        const text = this.cleanHtml(item[1]).trim();
        
        if (text) {
          product.qualityAssessment.certifications.push(text);
        }
      }
    }
    
    // Setup compliance verification if not already there
    if (!product.complianceVerification) {
      product.complianceVerification = {
        regional: [],
        industry: [],
        labeling: [],
        safety: []
      };
    }
    
    // Look for common safety and compliance terms
    const safetyPatterns = [
      { pattern: /\bce\s*compliant\b/i, field: "safety", value: "CE Compliant" },
      { pattern: /\bul\s*listed\b/i, field: "safety", value: "UL Listed" },
      { pattern: /\bastm\b/i, field: "industry", value: "ASTM Standards" },
      { pattern: /\bansi\b/i, field: "industry", value: "ANSI Standards" },
      { pattern: /\biec\b/i, field: "industry", value: "IEC Standards" },
      { pattern: /\brohs\b/i, field: "safety", value: "RoHS Compliant" },
      { pattern: /\breached\b/i, field: "safety", value: "REACH Compliant" },
      { pattern: /\bfcc\b/i, field: "regional", value: "FCC Compliant" },
      { pattern: /\blab\s*tested\b/i, field: "safety", value: "Laboratory Tested" }
    ];
    
    for (const pattern of safetyPatterns) {
      if (pattern.pattern.test(content)) {
        product.complianceVerification[pattern.field].push(pattern.value);
      }
    }
    
    // Remove duplicates
    product.qualityAssessment.certifications = [...new Set(product.qualityAssessment.certifications)];
    product.complianceVerification.regional = [...new Set(product.complianceVerification.regional)];
    product.complianceVerification.industry = [...new Set(product.complianceVerification.industry)];
    product.complianceVerification.safety = [...new Set(product.complianceVerification.safety)];
  }
  
  /**
   * Extract pricing and market information
   * @param {string} content - HTML content
   * @param {ProductEntity} product - Product entity to populate
   */
  extractPricingInfo(content, product) {
    // Setup market context if not already there
    if (!product.marketContext) {
      product.marketContext = {
        targetMarket: [],
        competition: [],
        distribution: [],
        pricing: {}
      };
    }
    
    // Extract regular and sale price
    const regularPricePattern = /<(?:div|span)[^>]*\b(?:id|class)="[^"]*\b(?:regular-price|old-price|original-price)[^"]*"[^>]*>([\s\S]*?)<\/(?:div|span)>/i;
    const salePricePattern = /<(?:div|span)[^>]*\b(?:id|class)="[^"]*\b(?:sale-price|special-price|discounted-price)[^"]*"[^>]*>([\s\S]*?)<\/(?:div|span)>/i;
    
    const regularPriceMatch = content.match(regularPricePattern);
    const salePriceMatch = content.match(salePricePattern);
    
    if (regularPriceMatch) {
      product.marketContext.pricing.regularPrice = this.cleanHtml(regularPriceMatch[1]).trim();
    }
    
    if (salePriceMatch) {
      product.marketContext.pricing.salePrice = this.cleanHtml(salePriceMatch[1]).trim();
    }
    
    // Extract availability
    const availabilityPattern = /<(?:div|span)[^>]*\b(?:id|class)="[^"]*\b(?:availability|stock|inventory)[^"]*"[^>]*>([\s\S]*?)<\/(?:div|span)>/i;
    const availabilityMatch = content.match(availabilityPattern);
    
    if (availabilityMatch) {
      product.marketContext.pricing.availability = this.cleanHtml(availabilityMatch[1]).trim();
    }
    
    // Look for shipping options
    const shippingPattern = /<(?:div|section)[^>]*\b(?:id|class)="[^"]*\b(?:shipping|delivery)[^"]*"[^>]*>([\s\S]*?)<\/(?:div|section)>/i;
    const shippingMatch = content.match(shippingPattern);
    
    if (shippingMatch) {
      const shippingText = this.cleanHtml(shippingMatch[1]).trim();
      if (shippingText) {
        product.marketContext.distribution.push("Shipping: " + shippingText);
      }
    }
    
    // Look for target market indicators
    const targetMarketPatterns = [
      { pattern: /\bfor\s*men\b/i, value: "Men" },
      { pattern: /\bfor\s*women\b/i, value: "Women" },
      { pattern: /\bfor\s*kids\b/i, value: "Children" },
      { pattern: /\bfor\s*children\b/i, value: "Children" },
      { pattern: /\bfor\s*babies\b/i, value: "Babies" },
      { pattern: /\bfor\s*seniors\b/i, value: "Seniors" },
      { pattern: /\bfor\s*professionals\b/i, value: "Professionals" },
      { pattern: /\bfor\s*beginners\b/i, value: "Beginners" }
    ];
    
    for (const pattern of targetMarketPatterns) {
      if (pattern.pattern.test(content)) {
        product.marketContext.targetMarket.push(pattern.value);
      }
    }
  }
  
  /**
   * Extract product images
   * @param {string} content - HTML content
   * @param {string} url - Product URL
   * @param {ProductEntity} product - Product entity to populate
   */
  extractProductImages(content, url, product) {
    // Check structured data first
    const structuredData = this.extractStructuredData(content);
    if (structuredData && structuredData.product && structuredData.product.image) {
      const images = Array.isArray(structuredData.product.image) 
        ? structuredData.product.image 
        : [structuredData.product.image];
      
      product.images = images.map(img => this.resolveUrl(img, url));
      return;
    }
    
    // Try common image patterns
    const imagePatterns = [
      /<img[^>]*\b(?:id|class)="[^"]*\b(?:product|item)[^"]*-image[^"]*"[^>]*src="([^"]*)"/gi,
      /<div[^>]*\b(?:id|class)="[^"]*\b(?:product|item)[^"]*-gallery[^"]*"[^>]*>[\s\S]*?<img[^>]*src="([^"]*)"/gi,
      /<a[^>]*\b(?:id|class|rel)="[^"]*\b(?:lightbox|gallery)[^"]*"[^>]*href="([^"]*)"/gi,
      /<meta[^>]*\bproperty="og:image"[^>]*\bcontent="([^"]*)"/i
    ];
    
    const images = [];
    
    for (const pattern of imagePatterns) {
      let match;
      const regex = new RegExp(pattern);
      while ((match = regex.exec(content)) !== null) {
        const imgUrl = this.resolveUrl(match[1], url);
        images.push(imgUrl);
      }
    }
    
    // If no product-specific images found, look for any image that might be a product image
    if (images.length === 0) {
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
        const imageUrl = this.resolveUrl(src, url);
        images.push(imageUrl);
        
        // Limit to a reasonable number of images
        if (images.length >= 5) break;
      }
    }
    
    product.images = images;
  }
  
  /**
   * Analyze product images with Cloud Vision API
   * @param {string} content - HTML content
   * @param {string} url - Product URL
   * @param {ProductEntity} product - Product entity to update
   */
  async analyzeProductImages(content, url, product) {
    // Skip if no Vision API key or no images
    if (!this.config.API?.GOOGLE?.VISION?.KEY || !product.images || product.images.length === 0) {
      return;
    }
    
    try {
      // Analyze the first product image
      const imageAnalysis = await this.cloudApiService.analyzeImage(product.images[0]);
      
      if (imageAnalysis) {
        // Setup image analysis if not already there
        if (!product.imageAnalysis) {
          product.imageAnalysis = {
            labels: [],
            text: "",
            objects: [],
            logos: []
          };
        }
        
        // Add labels
        if (imageAnalysis.labels) {
          product.imageAnalysis.labels = imageAnalysis.labels.map(label => ({
            description: label.description,
            score: label.score
          }));
        }
        
        // Add text extracted from the image
        if (imageAnalysis.text) {
          product.imageAnalysis.text = imageAnalysis.text;
        }
        
        // Add objects
        if (imageAnalysis.objects) {
          product.imageAnalysis.objects = imageAnalysis.objects.map(obj => ({
            name: obj.name,
            score: obj.score
          }));
        }
        
        // Add logos
        if (imageAnalysis.logos) {
          product.imageAnalysis.logos = imageAnalysis.logos.map(logo => ({
            description: logo.description,
            score: logo.score
          }));
        }
        
        // Use image analysis to enhance product information
        this.enhanceProductWithImageAnalysis(product);
      }
    } catch (error) {
      console.error(`Error analyzing product image: ${error.message}`);
    }
  }
  
  /**
   * Enhance product information with image analysis results
   * @param {ProductEntity} product - Product entity to enhance
   */
  enhanceProductWithImageAnalysis(product) {
    if (!product.imageAnalysis) {
      return;
    }
    
    // Use image labels to enhance categories if not already set
    if ((!product.mainCategory || !product.subCategory) && product.imageAnalysis.labels) {
      const categoryKeywords = {
        "Food": ["food", "dish", "cuisine", "meal", "ingredient", "snack", "dessert", "beverage"],
        "Clothing": ["clothing", "shirt", "dress", "pants", "jacket", "apparel", "fashion", "wear"],
        "Electronics": ["electronic", "device", "gadget", "computer", "phone", "television", "laptop"],
        "Home": ["furniture", "decor", "home", "household", "interior", "couch", "table", "chair"],
        "Beauty": ["beauty", "cosmetic", "skin care", "makeup", "lotion", "cream"],
        "Sports": ["sport", "fitness", "exercise", "athletic", "workout"],
        "Toys": ["toy", "game", "play"]
      };
      
      // Find matching category based on image labels
      for (const [category, keywords] of Object.entries(categoryKeywords)) {
        for (const label of product.imageAnalysis.labels) {
          if (keywords.some(keyword => label.description.toLowerCase().includes(keyword))) {
            if (!product.mainCategory) {
              product.mainCategory = category;
            } else if (!product.subCategory && product.mainCategory !== category) {
              product.subCategory = category;
            }
            break;
          }
        }
        
        if (product.mainCategory && product.subCategory) {
          break;
        }
      }
    }
    
    // Use extracted text to enhance product description if not already available
    if (!product.description && product.imageAnalysis.text) {
      const text = product.imageAnalysis.text;
      
      // Use first few sentences (up to 200 chars) as description
      const sentences = text.split(/[.!?]+/).filter(s => s.trim().length > 10);
      if (sentences.length > 0) {
        product.description = sentences.slice(0, 2).join(". ").substring(0, 200);
        if (product.description.length === 200) {
          product.description += "...";
        }
      }
    }
    
    // Add certification logos to quality assessment
    if (product.imageAnalysis.logos && product.qualityAssessment) {
      for (const logo of product.imageAnalysis.logos) {
        const logoText = logo.description.toLowerCase();
        
        // Common certification logos
        const certificationKeywords = [
          "organic", "fairtrade", "rainforest", "iso", "ce", "fda", "usda",
          "vegan", "vegetarian", "gluten free", "non gmo", "certified"
        ];
        
        if (certificationKeywords.some(keyword => logoText.includes(keyword))) {
          product.qualityAssessment.certifications.push(logo.description);
        }
      }
      
      // Remove duplicates
      product.qualityAssessment.certifications = [...new Set(product.qualityAssessment.certifications)];
    }
  }
  
  /**
   * Enhance product categories using AI
   * @param {Array} products - Array of ProductEntity objects
   */
  async enhanceProductCategories(products) {
    if (!this.config.API?.GOOGLE?.VERTEX_AI?.KEY || products.length === 0) {
      return;
    }
    
    try {
      // Loop through products without proper categorization
      for (const product of products) {
        if (!product.mainCategory || !product.subCategory) {
          const productText = `${product.productName} ${product.description}`;
          
          const textAnalysis = await this.cloudApiService.analyzeText(productText);
          
          if (textAnalysis && textAnalysis.classifications) {
            // Use classifications to set categories
            for (const classification of textAnalysis.classifications) {
              if (classification.displayName && classification.confidence > 0.7) {
                const categories = classification.displayName.split("/").map(c => c.trim());
                
                if (categories.length > 0 && !product.mainCategory) {
                  product.mainCategory = categories[0];
                }
                
                if (categories.length > 1 && !product.subCategory) {
                  product.subCategory = categories[1];
                }
                
                if (categories.length > 2 && !product.productFamily) {
                  product.productFamily = categories[2];
                }
                
                break;
              }
            }
          }
        }
      }
    } catch (error) {
      console.error(`Error enhancing product categories: ${error.message}`);
    }
  }
  
  /**
   * Extract product links from HTML content
   * @param {string} content - HTML content
   * @param {string} baseUrl - Base URL for resolving relative links
   * @return {Array} Product links
   */
  extractProductLinks(content, baseUrl) {
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
        const text = this.cleanHtml(linkMatch[2]).trim();
        
        // Skip empty links, non-product links, or navigation links
        if (!href || href === "#" || href.startsWith("javascript:") || 
            href.includes("login") || href.includes("cart") || 
            href.includes("account") || href.includes("contact")) {
          continue;
        }
        
        // Skip links without any text content
        if (!text) continue;
        
        // Resolve relative URL
        const fullUrl = this.resolveUrl(href, baseUrl);
        
        // Only include links from the same domain
        if (this.isSameDomain(fullUrl, baseUrl) && !this.isExcludedPath(fullUrl)) {
          // Check if this link seems product-related
          if (this.isLikelyProductLink(fullUrl, text)) {
            productLinks.push({
              url: fullUrl,
              text: text
            });
          }
        }
      }
    }
    
    return productLinks;
  }
  
  /**
   * Check if a page is likely a product page
   * @param {string} content - HTML content
   * @param {string} url - URL of the page
   * @return {boolean} Whether the page is likely a product page
   */
  isLikelyProductPage(content, url) {
    // URL-based checks
    const productUrlPatterns = [
      /\/product\//i,
      /\/p\/[a-z0-9]/i,
      /\/item\//i,
      /\/catalog\/product/i
    ];
    
    for (const pattern of productUrlPatterns) {
      if (pattern.test(url)) {
        return true;
      }
    }
    
    // Content-based checks
    const productIndicators = [
      /<(?:div|section)[^>]*\b(?:id|class)="[^"]*\b(?:product|single-product)[^"]*"[^>]*>/i,
      /<form[^>]*\b(?:id|class)="[^"]*\b(?:add-to-cart|buy-now)[^"]*"[^>]*>/i,
      /<(?:div|span)[^>]*\b(?:id|class)="[^"]*\b(?:price|product-price)[^"]*"[^>]*>/i,
      /<(?:div|section)[^>]*\b(?:id|class)="[^"]*\b(?:product-details|product-info)[^"]*"[^>]*>/i,
      /<button[^>]*\b(?:id|class)="[^"]*\b(?:add-to-cart|buy-now)[^"]*"[^>]*>/i
    ];
    
    for (const pattern of productIndicators) {
      if (pattern.test(content)) {
        return true;
      }
    }
    
    // Check for schema.org Product markup
    const schemaPattern = /<script[^>]*type="application\/ld\+json"[^>]*>([\s\S]*?)<\/script>/i;
    const schemaMatch = content.match(schemaPattern);
    
    if (schemaMatch) {
      try {
        const schema = JSON.parse(schemaMatch[1]);
        if (schema['@type'] === 'Product' || 
            (schema['@graph'] && schema['@graph'].some(item => item['@type'] === 'Product'))) {
          return true;
        }
      } catch (e) {
        // Ignore JSON parse errors
      }
    }
    
    return false;
  }
  
  /**
   * Extract title from HTML content
   * @param {string} content - HTML content
   * @return {string} Page title
   */
  extractTitleFromContent(content) {
    const h1Match = content.match(/<h1[^>]*>([\s\S]*?)<\/h1>/i);
    if (h1Match) {
      return this.cleanHtml(h1Match[1]).trim();
    }
    
    const titleMatch = content.match(/<title>(.*?)<\/title>/i);
    if (titleMatch) {
      return titleMatch[1].trim().split('|')[0].split('-')[0].trim();
    }
    
    return "";
  }
  
  /**
   * Extract categories from HTML content
   * @param {string} content - HTML content
   * @return {Array} Array of categories (main, sub, family)
   */
  extractCategories(content) {
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
            const text = this.cleanHtml(linkMatch[1]).trim();
            
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
                const text = this.cleanHtml(linkMatch[1]).trim();
                
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
                        const subText = this.cleanHtml(subLinkMatch[1]).trim();
                        
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
      
      return categories;
    } catch (e) {
      console.error(`Error extracting categories: ${e.message}`);
      return [];
    }
  }
  
  /**
   * Extract structured data from HTML content
   * @param {string} content - HTML content
   * @return {Object} Structured data object
   */
  extractStructuredData(content) {
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
  cleanHtml(html) {
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
   * Check if URL path should be excluded
   * @param {string} url - URL to check
   * @return {boolean} Whether the URL should be excluded
   */
  isExcludedPath(url) {
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
  isLikelyProductLink(url, text) {
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
   * Check if two URLs are from the same domain
   * @param {string} url1 - First URL
   * @param {string} url2 - Second URL
   * @return {boolean} Whether URLs are from the same domain
   */
  isSameDomain(url1, url2) {
    try {
      const getDomain = (url) => {
        const match = url.match(/^https?:\/\/([^/]+)/i);
        return match ? match[1] : '';
      };
      
      const domain1 = getDomain(url1);
      const domain2 = getDomain(url2);
      
      return domain1 === domain2;
    } catch (e) {
      return false;
    }
  }
  
  /**
   * Resolve relative URL to absolute URL
   * @param {string} url - Relative or absolute URL
   * @param {string} base - Base URL
   * @return {string} Absolute URL
   */
  resolveUrl(url, base) {
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
      
      // Use URL API for proper resolution
      return new URL(url, base).href;
    } catch (e) {
      console.error(`Error resolving URL: ${e.message}`);
      return url;
    }
  }
  
  /**
   * Capitalize first letter of a string
   * @param {string} string - String to capitalize
   * @return {string} Capitalized string
   */
  capitalizeFirstLetter(string) {
    return string.charAt(0).toUpperCase() + string.slice(1);
  }
}
