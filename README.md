# PDF-Page-Splitter
This code splits out different sized PDF pages so it's easier to print.  This is intended for notaries when printing large documents.
The output log file is a summary of how to fold the documents back together after printing the different tray sizes.
The most commonly used option is the split -s.  It will ask for a number to break up a large pdf into multiple pdfs. 
Otherwise, you give it files or filenames or globs and it will generate new pdfs grouped by page sizes.
