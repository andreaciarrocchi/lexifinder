# Lexifinder
Lexifinder is a GUI tool to create the analytical index (https://en.wikipedia.org/wiki/Index_(publishing)) of a manuscript.

## How to use
First, you need to convert your manuscript to PDF with your word processor. This is because docx and odt formats do not store information about page numbers: these are calculated at time of rendering. Therefore, the only way to obtain a reliable result is to convert your file to pdf. Then choose the PDF file and the text output file where to put the index. Similarity threshhold is set to 67 by default, but can range from 1 to 100. Insert the keywords you want to be sought in the text. They should be separeted by semicolons (e.g. "myword; another word; lastword"). Finally, press Start to initiate the process, which can be halted by pressing Cancel. The index will be saved in the specified path.

## How it works

## Future development
