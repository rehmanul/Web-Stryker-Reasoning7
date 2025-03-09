# Integrated Web Extraction System

This project combines the reliable extraction framework from Version 1 with the advanced capabilities of Version 2 into a single, robust web extraction system for Google Apps Script.

## Features

- **Reliable Core Extraction**: Built on Version 1's proven foundation
- **Enhanced Capabilities**: Integrated with Version 2's advanced features
- **Automatic Fallback**: Falls back to basic extraction if advanced features fail
- **Configurable**: Easy configuration through the spreadsheet interface
- **Comprehensive Data Collection**: Extracts titles, descriptions, content, links, images, and more
- **Advanced Analysis**: Includes language detection, sentiment analysis, and readability scoring

## Installation

1. Create a new Google Sheets document
2. Open Apps Script editor (Extensions > Apps Script)
3. Replace the default Code.gs file with the provided Code.gs content
4. Create a new file named appsscript.json and add the provided configuration
5. Save the project
6. Reload the spreadsheet

## Usage

1. After installation, you'll see a "Web Extraction" menu in your spreadsheet
2. The script will automatically create three sheets:
   - **Config**: Where you configure extraction settings and add URLs
   - **Data**: Where extracted data will be stored
   - **Log**: Where processing logs will be recorded
3. Add URLs to extract in the Config sheet under the "URLs to Process" section
4. Click on "Web Extraction" > "Run Extraction" to start the process
5. Monitor progress in the Log sheet
6. View extracted data in the Data sheet

## Configuration Options

In the Config sheet, you can modify these settings:

- **Enable Advanced Features**: Set to TRUE to use Version 2's enhanced extraction capabilities
- **Fallback to Basic**: Set to TRUE to fall back to basic extraction if advanced features fail
- **Max Retries**: Number of times to retry fetching a URL before giving up

## Advanced Features

The integrated system includes these advanced capabilities:

1. **Enhanced Metadata Extraction**: Captures Open Graph and Twitter card metadata
2. **Link and Image Extraction**: Collects all links and images with their contexts
3. **Language Detection**: Identifies the probable language of the content
4. **Sentiment Analysis**: Determines if content sentiment is positive, negative, or neutral
5. **Keyword Extraction**: Identifies the most important keywords in the content
6. **Readability Scoring**: Calculates the readability level of the content
7. **Entity Extraction**: Identifies people, organizations, and locations mentioned

## Troubleshooting

- If extraction fails, check the Log sheet for detailed error messages
- Ensure you have proper URL formatting in the Config sheet
- Verify that your Google Sheets has appropriate permissions for external requests
- If advanced features consistently fail, consider setting "Enable Advanced Features" to FALSE

## Performance Considerations

- The system is designed to respect Google Apps Script quotas and limitations
- For large numbers of URLs, consider processing in batches
- Advanced features require more processing time and resources
