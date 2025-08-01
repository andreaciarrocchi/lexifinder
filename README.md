# Lexifinder
Lexifinder is a GUI tool to create the analytical index (https://en.wikipedia.org/wiki/Index_(publishing)) of a manuscript.

## How to use
First, you need to convert your manuscript to PDF with your word processor. This is because docx and odt formats do not store information about page numbers: these are calculated at time of rendering. Therefore, the only way to obtain a reliable result is to convert your file to pdf. Then choose the PDF file and the text output file where to put the index. Similarity threshhold is set to 67 by default, but can range from 1 to 100. Insert the keywords you want to be sought in the text. They should be separeted by semicolons (e.g. "myword; another word; lastword"). Finally, press Start to initiate the process, which can be halted by pressing Cancel. The index will be saved in the specified path.

## How it works
Lexifinder opens the pdf file and extracts all the nouns. Then, it compares each of them to keywords. Words that reach the semantic similarity threshold are selected. To perform such a job, I used the spaCy library (https://pypi.org/project/spacy/) with a pretrained model. Finally, the app searches all occurrences in pages of each matching words and saves the result in a text file.<br>
The GUI has been designed with QT designer and the resulting ui file is loaded at startup. The Python script was converted to an executable file for Windows and Linux with PyInstaller (https://pyinstaller.org/en/stable/), with the following command

```
pyinstaller --onefile --icon=lexifinder.png -w --add-data "lexifinder.ui:." --add-data "en_core_web_md:en_core_web_md" lexifinder.py
```

On Linux, I then created an AppImage (https://appimage.org/) with appimagetool (https://github.com/AppImage/appimagetool).

```
./appimagetool Lexifinder.AppDir
```

## Releases
Two standalone releases are available: one for Windows (an executable file) and one for Linux (an AppImage). To ease the process of making the app available from the launcher, I packed the AppImage as a tar archive containing two BASH scripts to add (install_lexifinder.sh) and remove it (uninstall_lexifinder.sh).

## Future development
I would like to introduce some new features:
- adding support for other languages, as the app now supports only English;
- marking the selected words on the manuscript file (docx or odt), to allow the use of the word processor features related to index creation;
- adding a release for Mac, if I raise enough money to buy a Mac mini and pay the developer's fee
- introducing the ability to organize the index in two layers, but I do not know if that is feasible

## License
Lexifinder is free software licensed under the GNU General Public License v3 (GPLv3). You may redistribute and/or modify it under the terms of the GPLv3 as published by the Free Software Foundation. See the LICENSE file for details.
This software uses PySide6 (Qt for Python), which is released under the GPLv3 license. By using PySide6 under the GPL, this project must also comply with the terms of the GPLv3.
You are free to inspect, modify, and redistribute this software, provided that you preserve the GPL license and include the full source code when distributing.

## How to contribute
Suggestions, bug reporting, sponsorships and code contributions are welcomed. Feel free to share your ideas.
