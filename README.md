# PPTX-Text-Parser

This program is an adaptation of the [PPTX PowerPoint Translations](https://github.com/thormander/PPTX-Translator-OpenAI). The script is designed to extract text from PowerPoint (`.pptx`) files, clean the extracted text, and save it to a CSV file that can be used for further analysis, such as creating vector stores for semantic search.

## Features

- **Extract Text from PowerPoint Files**: The script processes individual `.pptx` files or entire folders of PowerPoint presentations.
- **Save Output to CSV**: The extracted text is saved to a CSV file in the same directory as the script.
- **Unique Identification**: Each slide's content is labeled with a `FileName_SlideNumber` identifier to ensure traceability.
- **Seamless KNIME Integration**: The CSV file can be easily loaded into KNIME for further text processing, embedding, and vector storage.

## Usage

1. **Place the Script**:
   - Download or clone the script to your local machine.

2. **Run the Script**:
   - Execute the script using Python, specifying the path to the PowerPoint file or folder containing `.pptx` files.
   - Example usage:
     ```bash
     python3 extractPPTXText.py /path/to/pptx/folder
     ```

3. **Output**:
   - The script will generate a `extracted_slide_texts.csv` file in the same directory where the script is located.

## Example CSV Output

The output CSV file will contain two columns:

- **FileName_SlideNumber**: Combines the base name of the PowerPoint file with the slide number (e.g., `presentation_Slide1`).
- **Slide Text**: Contains the cleaned and concatenated text from each slide.

Example:

| FileName_SlideNumber | Slide Text                                                          |
|----------------------|---------------------------------------------------------------------|
| presentation_Slide1  | "Text from the first slide."                                        |
| presentation_Slide2  | "Text from the second slide."                                       |
| presentation_Slide3  | "Text from the third slide."                                        |

## Cleaning Process

You can further clean the extracted text in KNIME by:

- **Removing excess whitespace**
- **Filtering out non-ASCII characters**
- **Excluding short documents based on word or character count**

## License

This project is based on [PPTX PowerPoint Translations](https://github.com/thormander/PPTX-Translator-OpenAI) and adheres to its licensing terms.
