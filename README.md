# pdf-and-pptx-to-docx
Extracts text on pages, slides and images in pdf/pptx and adds it to docx file.
Uses TesseractOCR for image text detection.

Input files should be put into "input" directory created after the program is first run. This directory can contain subdirectories - the program cen convert all files in given subdirectory.
Output will be saved to "output" directory. If converting files in subdirectory, the output will be saved to a subdirectory in output with the same name as input directory.
