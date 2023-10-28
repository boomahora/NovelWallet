# NovelWallet

## Overview
**NovelWallet** is a tool designed to replace words in a Word document with words provided in an Excel file. The replaced words are distributed across chapters, ensuring each word from the Excel file appears once in every chapter. After the replacement process, an Excel file is generated, detailing the inserted words, their page numbers, and corresponding chapters.

## Key Features

- **Dynamic Word Replacement**: Words are randomly replaced while ensuring punctuation remains consistent.
  
- **PDF Conversion**: Due to Word's dynamic pagination, the document is converted to PDF for accurate page number extraction.

- **Identifier System**: To track replaced words, identifiers are added around them. These identifiers are later removed, and the document is reverted to its clean state.

- **Chapter Range Tracking**: The tool keeps track of chapter ranges and each word's index within chapters. This facilitates accurate page number estimation after removing identifiers.

- **Error Handling**: Robust error handling mechanisms are implemented, including fuzzy matching for word extraction from PDF and chapter/page validation.

## How It Works

1. **PDF Conversion**: Convert the Word document to PDF to bypass Word's dynamic pagination limitations.
  
2. **Word Replacement with Identifiers**: Words from the Excel file are inserted into the PDF with surrounding identifiers to enable precise tracking.
  
3. **PDF to Word Conversion**: After inserting all words, the document is converted back to Word format, and all identifiers are removed.

4. **Chapter Range Analysis**: The tool then analyzes the chapter ranges before and after removing identifiers to estimate the page numbers of the replaced words.

5. **Error Handling**: Fuzzy matching is employed to handle discrepancies introduced during the PDF conversion. Additionally, the system ensures sequential page numbering and chapter tracking.

6. **Logging**: Any deviations or potential issues encountered during the word replacement process are logged for reference.

## Note

The word replacement process respects punctuation. For instance, a phrase like "hello, " will be replaced with "siesta, " to maintain the flow and format of the original document.

## Potential Issues & Resolutions

- **PDF Conversion Discrepancies**: The conversion process can introduce variations in chapter and page numbering. The tool ensures that each new page number is sequential, and chapters progress in order.

- **Fuzzy Matching**: To counteract minor inconsistencies during PDF conversion, the tool uses fuzzy matching against the list of inserted words.

- **Word Count Discrepancies**: If there's a mismatch between the expected and found word count, the information is logged. This discrepancy might arise if the last page of the document is treated as a separate chapter without any words.
