# Google Search Query Documents Downloader
 A versatile python based file downloader with a GUI that extracts files of specified types from search queries, manages duplicates, and logs download sources
 
## Overview

The **Google Search Query Document Downloader** is a versatile tool designed to streamline the process of downloading documents from Google search results. With a user-friendly GUI, this tool allows you to specify search queries and select file types (PDF, DOCX, DOC, PPTX, PPT, CSV, XLSX, XLS, RTF) for extraction. It efficiently handles multiple queries, avoids duplicate downloads, and logs download sources for record-keeping.

## Features

- **Search Query-Based Downloads**: Extract documents based on search queries from Google.
- **Multiple File Extensions**: Download documents in various formats, including PDF, DOCX, DOC, PPTX, PPT, CSV, XLSX, XLS, and RTF.
- **Duplicate Management**: Avoid re-downloading the same files with built-in caching.
- **Logging**: Automatically generates an Excel file listing the download links and filenames for each query.
- **User-Friendly Interface**: Intuitive GUI for easy configuration of search queries, output directory, and file types.
- **Progress and Status Updates**: Displays status messages and success notifications to keep users informed.

## Installation

1. **Clone the Repository**:
   
   ```bash
   git clone https://github.com/rishabhc9/Google-Search-Query-Documents-Downloader.git
   cd google-search-query-document-downloader
   ```
   
3. **Install Required Packages:**:
Ensure you have Python installed, then install the required packages:

   ```bash
   pip install -r requirements.txt
   ```
   
5. **Run the Tool:**:
Execute the main script using the following command:

   ```bash
   python main.py
   ```

6.**Run the Multi Page Scraping Tool:**:
Execute the multi page scraping script using the following command:

   ```bash
   python multi_page_scrape.py
   ```

## Usage

1. **Select Search Query File**:
   - Click the "Browse" button to select an Excel file containing your search queries. Ensure the queries are listed in a column labeled "queries."

2. **Choose Output Directory**:
   - Click the "Browse" button to select the directory where the downloaded files will be saved. Each search query will be processed into its own subfolder within this directory.

3. **Select File Extension**:
   - Use the dropdown menu to choose the file type(s) you want to download. Options include PDF, DOCX, DOC, PPTX, PPT, CSV, XLSX, XLS, and RTF.

4. **Specify Number of Search Pages**:
   - Enter the number of search result pages to process per query. This determines how many pages of search results will be scraped for each query.

5. **Start Scraping**:
   - Click the "Scrape" button to begin the download process. The tool will display a status message indicating that scraping has started and will update with a success message once the process is complete.

6. **Clear Cache**:
   - Click the "Clear Cache" button to remove the cache of previously downloaded files. This will reset the record of downloaded files and allow for fresh downloads if you rerun the tool.

## Screenshots
<img width="944" alt="Screenshot 2024-08-22 at 11 41 42 AM" src="https://github.com/user-attachments/assets/0b724165-dc6e-4567-a376-45556a4f0469">


