# Directory to ENEX Exporter

This PowerShell script (`LogseqToJoplin-Convert.ps1`) exports notes from any directory containing Markdown, Org-mode, DOCX, or ODT files into an [Evernote ENEX](https://help.evernote.com/hc/en-us/articles/209005557-Export-notes-and-notebooks) file, which can be imported into [Joplin](https://joplinapp.org/), Evernote, or any other application that supports the ENEX format.

## Features
- **Universal Directory Export:** Converts all supported files in a directory (and subdirectories) to ENEX for Joplin/Evernote import.
- **Logseq Support:** Optionally indexes Logseq content, resolving block references, embeds, and page links.
- **Image Handling:** Extracts and embeds images from Markdown, Org-mode, DOCX, and ODT files, including those extracted by Pandoc.
- **Pandoc Integration:** Uses [Pandoc](https://pandoc.org/) to convert Markdown, Org-mode, DOCX, and ODT files to HTML for ENEX.
- **Tag Extraction:** Derives tags from file naming conventions (triple underscores).
- **Metadata Preservation:** Preserves creation and modification dates in the ENEX export.
- **Batch Processing:** Processes all supported files in the directory tree.

## Usage
1. **Requirements:**
   - PowerShell (Windows, macOS, or Linux)
   - [Pandoc](https://pandoc.org/) installed and available in your PATH

2. **Install Pandoc:**
   - Visit [Pandoc Install](https://pandoc.org/installing.html) for detailed instructions.
   - **Windows:** Download the installer from the [Pandoc releases page](https://github.com/jgm/pandoc/releases) and run it.
   - **macOS:** Use [Homebrew](https://brew.sh/):  
     ```sh
     brew install pandoc
     ```
   - **Linux:** Use your package manager, e.g.:  
     ```sh
     sudo apt-get install pandoc
     ```
   - After installation, ensure `pandoc` is available in your terminal by running:
     ```sh
     pandoc --version
     ```

3. **Run the script:**
   ```powershell
   .\LogseqToJoplin-Convert.ps1 -InputDir "<path-to-your-notes-directory>" [-DirectoryType Logseq|Other] [-CreateFolderStructure]
   ```
   - Replace `<path-to-your-notes-directory>` with the path to your notes.
   - Use `-DirectoryType Logseq` for Logseq exports (default), or `Other` for generic directories.
   - Optional: add `-CreateFolderStructure` to create a physical folder structure instead of an ENEX export.

   Modes when using `-CreateFolderStructure`:
   - Logseq mode (`-DirectoryType Logseq`): files are organized into folders derived from their tags. Hierarchical tags like `[[Project/Subtask/Item]]` will create nested folders (`Project\Subtask\Item\note.md`). Logseq block references and embeds are resolved.
   - Other mode (`-DirectoryType Other`): the script preserves the input directory's folder hierarchy and converts supported files in-place to Markdown files under the output folder. For example, `InputDir\FolderA\Note.docx` becomes `OutputFolder\FolderA\Note.md` after conversion.

4. **Import into Joplin or Evernote:**
   - Open Joplin or Evernote
   - Go to `File > Import > ENEX`
   - Select the generated ENEX file (e.g., `<InputDir>_evernote_export.enex`)

## Output
- **ENEX File (default):** A single `.enex` file named `<InputDir>_evernote_export.enex` containing all notes and embedded resources, ready for import into any ENEX-compatible application.

- **Folder Structure (when using `-CreateFolderStructure`):**
   - The script creates an output folder named `FolderStructure` next to the generated ENEX file (or next to your `InputDir` when used). Inside it:
      - In Logseq mode, files are placed in folders derived from their tags (hierarchical tags create nested folders).
      - In Other mode, the original folder hierarchy from the input directory is recreated and files are converted to Markdown while preserving relative paths.
   - Filenames are sanitized for filesystem compatibility and all converted files are UTF-8 encoded.

- ## Limitations
- Only Markdown (`.md`), Org-mode (`.org`), DOCX (`.docx`), and ODT (`.odt`) files are supported.
- Requires Pandoc for conversion.
- Some advanced Logseq or document features may not be fully supported.

Notes:
- When using `-CreateFolderStructure` in Other mode, office documents (`.docx`, `.odt`) are converted to Markdown using Pandoc and the converted Markdown is written in the matching output location.
- The `FolderStructure` output directory will be removed and recreated when the script runs with `-CreateFolderStructure` to ensure a clean output. Back up any existing folder if needed.

## License
MIT License
