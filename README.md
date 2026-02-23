# LexiFinder

**LexiFinder** is a tool to generate [analytic indexes](https://en.wikipedia.org/wiki/Index_(publishing)) from documents automatically. Given one or more source documents and a set of keywords, it extracts all nouns, compares them semantically to the keywords using a pretrained NLP model, and produces a structured, hierarchical index ready to be included in a book or manuscript.

LexiFinder works in two ways: as a **command-line tool** for scripting, automation, and batch processing, and as a **graphical application** for a guided, point-and-click experience. Both interfaces share the same underlying engine and support the same features.

Supported input formats are **PDF**, **DOCX**, and **ODT**. The index can be exported as plain text, JSON, CSV, or HTML.

---

## Releases

Two standalone releases are available, requiring no Python installation:

- **Windows**: `lexifinder.exe` — a self-contained executable built with [PyInstaller](https://pyinstaller.org/).
- **Linux**: `LexiFinder-x86_64.AppImage` — a portable AppImage built with PyInstaller and [appimagetool](https://github.com/AppImage/appimagetool).

Both releases bundle the default English NLP model (`en_core_web_md`). Additional language models can be downloaded from within the application.

---

## How It Works

LexiFinder uses [spaCy](https://spacy.io/) pretrained language models to analyse the document. The processing pipeline works as follows:

1. **Text extraction** — the document is read and its text is extracted, either page by page (PDF) or paragraph by paragraph (DOCX/ODT).
2. **Noun extraction** — all nouns are identified using spaCy's part-of-speech tagger.
3. **Semantic matching** — each noun is compared to your keywords (or grouped automatically, depending on the chosen strategy) using word-vector similarity. Only nouns that reach or exceed the similarity threshold are retained.
4. **Index assembly** — matched nouns are organised into a hierarchical structure and their locations (page numbers or paragraph references) are recorded.
5. **Export** — the result is saved to a file in the format of your choice.

### Indexing Strategies

LexiFinder supports four strategies for building the index hierarchy:

| Strategy | Levels | Description |
|---|---|---|
| `keywords` | 2 | Your keywords become the main entries; semantically related nouns become sub-entries. This is the default. |
| `hybrid` | 3 | Like `keywords`, but sub-entries are further grouped into automatic sub-clusters. |
| `auto` | 2 | Fully automatic: nouns are clustered with K-means and cluster names are inferred from the data. No keywords needed. |
| `frequent` | 2 | The most frequent nouns in the document become the main entries. No keywords needed. |

### Language Support

LexiFinder supports **English**, **Italian**, **German**, **French**, and **Spanish** through the corresponding spaCy models. The default bundled model is the English medium model (`en_core_web_md`). Other models can be downloaded directly from the application.

---

## Command-Line Interface

The CLI is activated whenever LexiFinder is launched with at least one argument. Running it without arguments (or with `--gui`) opens the graphical interface instead.

### Basic Usage

```
lexifinder -i <input_file> -o <output_file> [options]
```

### Arguments

| Argument | Description |
|---|---|
| `-i`, `--input FILE` | Path to the input document (`.pdf`, `.docx`, or `.odt`). |
| `-o`, `--output FILE` | Path to the output index file. |
| `-k`, `--keywords "word1; word2"` | Keywords separated by semicolons. Required for `keywords` and `hybrid` strategies. |
| `-t`, `--threshold FLOAT` | Similarity threshold from `0.0` to `1.0` (default: `0.5`). Lower values cast a wider net. |
| `-m`, `--mode {page,paragraph}` | Whether locations are reported as page numbers or paragraph numbers (default: `page` for PDF, `paragraph` for DOCX/ODT). |
| `-f`, `--export-format {txt,json,csv,html}` | Output format (default: `txt`). |
| `--strategy {keywords,auto,hybrid,frequent}` | Indexing strategy (default: `keywords`). |
| `--max-per-category N` | Maximum number of sub-entries per category (default: `30`). |
| `--clusters N` | Number of clusters for the `auto` strategy (default: `8`). |
| `--subclusters N` | Number of sub-clusters for the `hybrid` strategy (default: `3`). |
| `--top N` | Number of top frequent terms for the `frequent` strategy (default: `15`). |
| `--exclude-generic` | Remove common filler nouns such as *thing*, *aspect*, *factor*, etc. |
| `--min-occurrences N` | Only include terms that appear at least N times in the document (default: `1`). |
| `--model MODEL` | Name of the spaCy model to use (default: `en_core_web_md`). |
| `-x`, `--mark` | Instead of a text index, insert index entry fields directly into a DOCX or ODT file so that the word processor can generate its own native index. |
| `--save-config FILE` | Save the current configuration to a JSON file for later reuse. |
| `--load-config FILE` | Load a previously saved configuration. Command-line arguments override loaded values. |
| `--gui` | Launch the graphical interface. |
| `--version` | Show version information and exit. |

### Practical Examples

**Basic keyword index from a PDF, reported by page:**
```bash
lexifinder -i thesis.pdf -o index.txt -k "climate change; biodiversity; ecosystem"
```

**Keyword index from a Word document, reported by paragraph:**
```bash
lexifinder -i manuscript.docx -o index.txt -k "justice; democracy; rights"
```

**Stricter matching with filtering — only well-attested, meaningful terms:**
```bash
lexifinder -i book.pdf -o index.txt -k "AI; machine learning" --exclude-generic --min-occurrences 3
```

**Three-level hierarchical index using the hybrid strategy:**
```bash
lexifinder -i document.pdf -o index.txt -k "physics; chemistry" --strategy hybrid --subclusters 3
```
Output structure:
```
PHYSICS
  [Sub-group 1]
    quantum: p.12; p.45; p.67
    particle: p.15; p.48
  [Sub-group 2]
    radiation: p.23; p.31
```

**Fully automatic index — no keywords required:**
```bash
lexifinder -i report.pdf -o index.txt --strategy auto --clusters 10
```

**Export to HTML for a styled, web-ready index:**
```bash
lexifinder -i paper.pdf -o index.html -k "neuroscience; cognition" -f html
```

**Index in Italian using the Italian medium model:**
```bash
lexifinder -i romanzo.pdf -o indice.txt -k "amore; guerra; identità" --model it_core_news_md
```

**Mark a DOCX file so that Word can generate its own native index:**
```bash
lexifinder -i thesis.docx -o /dev/null -k "science; technology" --mark
```
This creates a `thesis_marked.docx` file. Open it in Word, place the cursor where you want the index, and go to *References → Insert Index → OK*. For LibreOffice, go to *Insert → Table of Contents and Index → Alphabetical Index → OK*.

### Batch Processing

LexiFinder can process an entire folder of documents in a single run.

**Index all supported files in a directory:**
```bash
lexifinder --batch-dir ./chapters --output-dir ./indexes -k "philosophy; ethics"
```

**Process only PDF files matching a pattern:**
```bash
lexifinder --batch-dir ./chapters --output-dir ./indexes --pattern "chapter*.pdf" -k "history"
```

**Batch processing using a saved configuration:**
```bash
lexifinder --batch-dir ./docs --output-dir ./output --load-config my-config.json
```

Each input file produces a separate output file named `<original_name>-index.<format>` in the output directory.

### Configuration Files

You can save a set of options to a JSON file and reuse it across multiple runs, which is convenient for processing a multi-chapter book.

**Save the current configuration:**
```bash
lexifinder -i ch1.pdf -o ch1-index.txt -k "science; technology" --strategy hybrid --save-config book-config.json
```

**Reuse it for subsequent chapters, overriding only the file paths:**
```bash
lexifinder -i ch2.pdf -o ch2-index.txt --load-config book-config.json
lexifinder -i ch3.pdf -o ch3-index.txt --load-config book-config.json
```

### Model Management

```bash
lexifinder --list-models              # Show all supported spaCy models
lexifinder --list-installed           # Show installed models
lexifinder --download-model it_core_news_md   # Download the Italian medium model
lexifinder --delete-model it_core_news_md     # Delete an installed model
```

---

## Graphical Interface

Launching LexiFinder without arguments, or with the `--gui` flag, opens the graphical interface. The GUI adapts to your system's dark or light theme automatically.

The interface is organised into four tabs accessible from the bottom navigation bar:

### Single File

This tab processes one document at a time. Fill in the fields as follows:

- **Input file** — select your PDF, DOCX, or ODT file.
- **Output file** — choose where to save the index.
- **Keywords** — enter your keywords separated by semicolons, for example: `war; peace; identity`. Leave empty if using the `auto` or `frequent` strategy.
- **Similarity threshold** — adjust the slider to set how strictly nouns must match your keywords. A value of `0.5` is a good starting point; lower values include more loosely related terms, higher values are more selective.
- **Mode** — choose *page* (PDF only) or *paragraph*.
- **Strategy** — select one of the four indexing strategies described above.
- **Export format** — choose between plain text, JSON, CSV, or HTML.
- **Additional options** — enable *Exclude generic words* and/or set a minimum occurrence count to refine the index.
- **Mark document** — for DOCX and ODT files, tick this option to insert index entry fields directly into the document instead of producing a separate index file.

Click **▶ Run** to start. Progress and log messages appear in the log panel on the right. Click **■ Stop** to cancel a running job.

### Batch Mode

This tab mirrors the Single File options but applies them to an entire folder. Specify an input directory, an output directory, and optionally a file name pattern (e.g. `chapter*.pdf`). All matched files are processed sequentially and a summary is shown at the end.

### Models

This tab lists the spaCy models currently installed on your system and lets you download or delete them with a single click. The bundled default model (`en_core_web_md`) cannot be deleted.

### About

Shows version information, links to the project repository and author's website, and a donation link.

---

## How to Contribute

Contributions of any kind are welcome:

- **Bug reports** — if something does not work as expected, please open an issue on the [GitHub repository](https://github.com/andreaciarrocchi/lexifinder) with a description of the problem, the document format used, and the command or settings that triggered it.
- **Feature requests** — suggestions for new features or improvements can be submitted as issues. Ideas currently under consideration include support for additional languages, macOS release, and deeper integration with word processor index workflows.
- **Code contributions** — pull requests are welcome. Please keep changes focused and describe what they do and why.
- **Sponsorship** — if LexiFinder saves you time, consider supporting its development via [PayPal](https://paypal.me/ciarro85).

---

## License

LexiFinder is free software licensed under the **GNU General Public License v3** (GPLv3). You may redistribute and/or modify it under the terms of the GPLv3 as published by the Free Software Foundation. See the `LICENSE` file for details.

You are free to inspect, modify, and redistribute this software, provided that you preserve the GPL license and include the full source code when distributing.
