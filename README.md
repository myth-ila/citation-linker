# Citation Linker

A Python tool that automatically finds academic citations in Microsoft Word 
documents and converts them into clickable hyperlinks that jump to the 
corresponding bibliography entry.

## What It Does

If your Word document contains something like:

> ...early forensic methods significantly advanced criminal investigation 
> techniques (Murdoch and Ogden 1900), reshaping how detectives 
> approached evidence collection...

This tool will turn `Murdoch and Ogden 1900` into a **clickable hyperlink** 
that jumps to the full reference in your bibliography section.

> **Note:** All citations and author names used in examples are fictitious. 
> They are named after characters from the Canadian TV series 
> [*Murdoch Mysteries*](https://en.wikipedia.org/wiki/Murdoch_Mysteries) 
> and do not represent real academic publications.

## Features

- Automatically extracts bibliography entries (author names + year)
- Matches in-text citations to bibliography entries
- Handles multiple citation formats:
  - Single author: `(Crabtree 1912)`
  - Two authors: `(Murdoch and Ogden 1900)`
  - Three+ authors: `(Murdoch, Ogden, and Crabtree 1914)`
  - Multi-word last names: `(Conan Doyle, Holmes, and Watson 1895)`
- Handles semicolon-separated citations: `(Murdoch and Ogden 1900; Brackenreid 1910)`
- Cleans prefixes like `e.g.,` and `see`
- Creates internal bookmarks and hyperlinks within the document

## Requirements

- Python 3.7+
- `python-docx` library

## Installation

This tool requires the `python-docx` library to read and write Word documents. 
Open your terminal (Command Prompt on Windows, Terminal on Mac) and run:

    ```bash
    pip install python-docx
    ```
If you're using a virtual environment or Anaconda, make sure it's activated first.

## Usage
1. Place your Word document (`.docx`) in the same folder as the script
2. Edit the file paths at the bottom of `add_hyperlinks.py`:

    ```python
    input_file = r"Your_Document.docx"
    output_file = r"Your_Document_linked.docx"
    ```
3. Run the script:

    ```bash
    python add_hyperlinks.py
    ```
4. Open the output file — your citations are now hyperlinked!

## Assumptions

- Your document has a References or Bibliography section heading
- Bibliography entries include the year in parentheses, e.g., (2024)
- In-text citations are in parentheses using author last names and year
- Citation style follows APA-like conventions (e.g., Author, Author, and Author YEAR)

## Example Output

```
Found bibliography entry: ['Murdoch', 'Ogden'] (1900)
Found bibliography entry: ['Crabtree'] (1912)
Found bibliography entry: ['Conan Doyle', 'Holmes', 'Watson'] (1895)
Found citation: Murdoch and Ogden 1900
Found citation: Crabtree 1912
Found citation: Conan Doyle, Holmes, and Watson 1895
Document saved to Your_Document_linked.docx
```
## Development Process

This tool was developed with AI assistance from Anthropic's Claude (Haiku 4.5 and Opus 4.6). 
The code was built through an iterative process — I described the goal, tested 
outputs on my document, identified edge cases and failures, and worked with 
Claude to diagnose and fix them over multiple rounds.

## License

MIT License — feel free to use, modify, and distribute.
