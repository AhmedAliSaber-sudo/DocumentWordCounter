# Document Word Counter

## Description
Document Word Counter is a Windows application that counts words in various document types (.docx, .doc, .pptx, .xlsx) within a specified folder. It generates an Excel file with the document names, word counts, and last modified dates.

## Features
- User-friendly graphical interface
- Supports multiple file types: .docx, .doc, .pptx, .xlsx
- Processes all supported documents in a selected folder
- Displays real-time progress during processing
- Generates an Excel file with results
- Shows completion message with the location of the output file

## Installation

### Prerequisites
- Windows operating system
- Microsoft Office (for processing .doc files)

### Steps
1. Download the `DocumentWordCounter.exe` file from the latest release.
2. Place the executable in your desired location on your computer.

## Usage
1. Run `DocumentWordCounter.exe`.
2. Click the "Browse" button to select the folder containing your documents.
3. Click "Process Files" to start counting words.
4. Wait for the process to complete. You'll see progress updates in the application window.
5. Once finished, a message box will show the location of the created Excel file.

## Output
The application generates an Excel file named "Document_Names_and_Word_Counts.xlsx" in the same folder as the processed documents. This file contains:
- Document Name
- Word Count
- Date Modified

## Troubleshooting
- If the application doesn't start, ensure that you have the necessary permissions to run executables from your chosen location.
- For .doc files, make sure Microsoft Word is installed on your system.
- If your antivirus flags the application, you may need to add an exception for it.

## Known Limitations
- Large numbers of files or very large documents may take a significant time to process.
- Word counts for Excel files include all non-empty cells, which may lead to higher counts than expected if there's a lot of numerical data.
- PowerPoint word counts include all text in shapes and text boxes, which might include hidden or background elements.

## Future Enhancements
- Support for additional file formats
- Option to customize output file name and location
- Improved error handling and logging
- Multithreading for faster processing

## Support
If you encounter any issues or have suggestions for improvements, please open an issue in the GitHub repository or contact the developer.

## License
[Specify your chosen license here]

## Acknowledgments
This application uses the following open-source libraries:
- docx2txt
- openpyxl
- python-pptx
- pywin32
