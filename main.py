import os
import re
import shutil
import tempfile
from typing import List, Optional
from fastapi import FastAPI, File, UploadFile, HTTPException, Form
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
import uvicorn
from pypdf import PdfWriter, PdfReader
import docx

app = FastAPI(title="PDF and DOCX Merger")

# Create directories if they don't exist
os.makedirs("uploads", exist_ok=True)
os.makedirs("static", exist_ok=True)

# Mount static files
app.mount("/static", StaticFiles(directory="static"), name="static")

# Helper function to extract part number from filename
def extract_part_number(filename: str) -> int:
    """Extract part number from filename (e.g., 'file_part1.pdf' -> 1)"""
    match = re.search(r'part(\d+)', filename.lower())
    if match:
        return int(match.group(1))
    return 0  # Default to 0 if no part number found

# Helper function to sort files by part number
def sort_files_by_part(files: List[UploadFile]) -> List[UploadFile]:
    """Sort files by their part number in the filename"""
    return sorted(files, key=lambda file: extract_part_number(file.filename))

# Helper function to merge PDF files
def merge_pdf_files(file_paths: List[str], output_path: str) -> None:
    """Merge multiple PDF files into a single PDF"""
    merger = PdfWriter()
    
    for file_path in file_paths:
        reader = PdfReader(file_path)
        for page in reader.pages:
            merger.add_page(page)
    
    with open(output_path, "wb") as output_file:
        merger.write(output_file)

# Helper function to merge DOCX files
def merge_docx_files_custom(file_paths: List[str], output_path: str) -> None:
    """Merge multiple DOCX files into a single DOCX"""
    combined_doc = docx.Document()
    
    for i, file_path in enumerate(file_paths):
        doc = docx.Document(file_path)
        
        # Skip adding a page break before the first document
        if i > 0:
            combined_doc.add_page_break()
        
        # Copy content from each document
        for element in doc.element.body:
            combined_doc.element.body.append(element)
    
    combined_doc.save(output_path)

@app.get("/", response_class=HTMLResponse)
async def get_upload_page():
    """Return the HTML upload page"""
    html_content = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>PDF and DOCX Merger</title>
        <style>
            body {
                font-family: Arial, sans-serif;
                max-width: 800px;
                margin: 0 auto;
                padding: 20px;
                line-height: 1.6;
            }
            h1 {
                color: #333;
                text-align: center;
            }
            .container {
                background-color: #f9f9f9;
                border-radius: 8px;
                padding: 20px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            }
            .form-group {
                margin-bottom: 15px;
            }
            label {
                display: block;
                margin-bottom: 5px;
                font-weight: bold;
            }
            input[type="file"] {
                width: 100%;
                padding: 10px;
                border: 1px solid #ddd;
                border-radius: 4px;
            }
            button {
                background-color: #4CAF50;
                color: white;
                padding: 10px 15px;
                border: none;
                border-radius: 4px;
                cursor: pointer;
                font-size: 16px;
            }
            button:hover {
                background-color: #45a049;
            }
            .note {
                font-size: 0.9em;
                color: #666;
                margin-top: 5px;
            }
            .loading {
                display: none;
                text-align: center;
                margin-top: 20px;
            }
            .spinner {
                border: 4px solid rgba(0, 0, 0, 0.1);
                width: 36px;
                height: 36px;
                border-radius: 50%;
                border-left-color: #09f;
                animation: spin 1s linear infinite;
                display: inline-block;
            }
            @keyframes spin {
                0% { transform: rotate(0deg); }
                100% { transform: rotate(360deg); }
            }
        </style>
    </head>
    <body>
        <h1>PDF and DOCX Merger</h1>
        <div class="container">
            <form action="/merge/" enctype="multipart/form-data" method="post" onsubmit="showLoading()">
                <div class="form-group">
                    <label for="files">Select files to merge:</label>
                    <input type="file" name="files" id="files" multiple required accept=".pdf,.docx,.doc">
                    <p class="note">Files will be merged in order based on "part" numbers in their names (e.g., file_part1.pdf, file_part2.pdf)</p>
                </div>
                <div class="form-group">
                    <label for="output_filename">Output filename (without extension):</label>
                    <input type="text" name="output_filename" id="output_filename" value="merged_document">
                </div>
                <button type="submit">Merge Files</button>
            </form>
            <div id="loading" class="loading">
                <div class="spinner"></div>
                <p>Processing files, please wait...</p>
            </div>
        </div>

        <script>
            function showLoading() {
                document.getElementById('loading').style.display = 'block';
            }
        </script>
    </body>
    </html>
    """
    return html_content

@app.post("/merge/")
async def merge_files(
    files: List[UploadFile] = File(...),
    output_filename: str = Form("merged_document")
):
    """Merge uploaded files in order based on part numbers in their filenames"""
    if not files:
        raise HTTPException(status_code=400, detail="No files uploaded")
    
    # Check if all files are of the same type
    file_extensions = [os.path.splitext(file.filename)[1].lower() for file in files]
    
    if not all(ext in ['.pdf', '.docx', '.doc'] for ext in file_extensions):
        raise HTTPException(status_code=400, detail="Only PDF and DOCX/DOC files are supported")
    
    if len(set(file_extensions)) > 1:
        raise HTTPException(status_code=400, detail="All files must be of the same type (either all PDF or all DOCX/DOC)")
    
    # Sort files by part number
    sorted_files = sort_files_by_part(files)
    
    # Create a temporary directory to store uploaded files
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_file_paths = []
        
        # Save uploaded files to temporary directory
        for file in sorted_files:
            temp_file_path = os.path.join(temp_dir, file.filename)
            with open(temp_file_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            temp_file_paths.append(temp_file_path)
        
        # Determine file type and merge accordingly
        file_ext = os.path.splitext(sorted_files[0].filename)[1].lower()
        output_path = f"uploads/{output_filename}{file_ext}"
        
        if file_ext == '.pdf':
            merge_pdf_files(temp_file_paths, output_path)
        else:  # .docx or .doc
            merge_docx_files_custom(temp_file_paths, output_path)
    
    return FileResponse(
        path=output_path,
        filename=f"{output_filename}{file_ext}",
        media_type="application/octet-stream"
    )

if __name__ == "__main__":
    # Get port from environment variable or default to 8000
    port = int(os.environ.get("PORT", 8000))
    # Use 0.0.0.0 for production, 127.0.0.1 for local development
    host = os.environ.get("HOST", "127.0.0.1")
    uvicorn.run("main:app", host=host, port=port, reload=True)
