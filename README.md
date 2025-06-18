# PDF and DOCX Merger

A FastAPI web application that merges PDF or DOCX files in a specific order based on their filenames.

## Features

- Upload multiple PDF or DOCX files
- Automatically sort files based on part numbers in filenames (e.g., file_part1.pdf, file_part2.pdf)
- Merge files into a single document
- Download the merged document

## Requirements

- Python 3.7+
- Dependencies listed in `requirements.txt`

## Installation

1. Clone this repository or download the files
2. Install the required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

1. Start the application:

```bash
python main.py
```

2. Open your web browser and navigate to `http://localhost:8000`
3. Upload your files (all must be of the same type - either all PDF or all DOCX)
4. Enter a name for the output file (optional)
5. Click "Merge Files"
6. The merged document will be downloaded automatically

## File Naming Convention

The application sorts files based on "part" numbers in their filenames. For example:
- document_part1.pdf
- document_part2.pdf
- document_part3.pdf

Files will be merged in ascending order of their part numbers.

## Notes

- All files must be of the same type (either all PDF or all DOCX)
- The application creates `uploads` and `static` directories if they don't exist
- Merged files are stored in the `uploads` directory
