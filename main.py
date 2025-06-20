import os
import re
import shutil
import tempfile
import zipfile
import rarfile
import patoolib
from typing import List, Optional, Tuple
from fastapi import FastAPI, File, UploadFile, HTTPException, Form, Request
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

# Add CORS middleware to allow cross-origin requests
from fastapi.middleware.cors import CORSMiddleware

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Allows all origins
    allow_credentials=True,
    allow_methods=["*"],  # Allows all methods
    allow_headers=["*"],  # Allows all headers
)

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




# Helper function to extract files from compressed archives
def extract_compressed_file(file_path: str, extract_dir: str) -> List[str]:
    """Extract files from ZIP or RAR archive and return paths to extracted files"""
    file_ext = os.path.splitext(file_path)[1].lower()
    extracted_files = []
    
    try:
        if file_ext == '.zip':
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
                extracted_files = [os.path.join(extract_dir, name) for name in zip_ref.namelist() 
                                if not name.endswith('/')]  # Skip directories
        elif file_ext == '.rar':
            rarfile.UNRAR_TOOL = 'unrar'  # Make sure unrar is installed on the system
            with rarfile.RarFile(file_path) as rar_ref:
                rar_ref.extractall(extract_dir)
                extracted_files = [os.path.join(extract_dir, name) for name in rar_ref.namelist() 
                                if not rar_ref.getinfo(name).isdir()]
        else:
            # Use patool for other archive formats
            patoolib.extract_archive(file_path, outdir=extract_dir)
            # Get all files in the extraction directory
            for root, _, files in os.walk(extract_dir):
                for file in files:
                    extracted_files.append(os.path.join(root, file))
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error extracting archive: {str(e)}")
    
    return extracted_files

# Helper function to filter files by extension
def filter_files_by_extension(file_paths: List[str], extensions: List[str]) -> List[str]:
    """Filter files by their extensions"""
    return [path for path in file_paths if os.path.splitext(path)[1].lower() in extensions]

# Helper function to merge DOCX files
def merge_docx_files_custom(file_paths: List[str], output_path: str) -> None:
    """Merge multiple DOCX files into a single DOCX"""
    # Start with the first document as the base
    if not file_paths:
        return
    
    # Use the first document as a base and then append others with page breaks
    first_doc = docx.Document(file_paths[0])
    
    # Process each additional document
    for i, file_path in enumerate(file_paths[1:], 1):  # Start from the second document
        # Add a section break to ensure new document starts on a new page
        first_doc.add_section()
        section = first_doc.sections[-1]
        section.start_type = 2  # New page section break
        
        # Load the document to append
        doc = docx.Document(file_path)
        
        # Append all paragraphs from the document
        for para in doc.paragraphs:
            # Create a new paragraph with the same style
            p = first_doc.add_paragraph()
            p.style = para.style
            
            # Copy all runs with their formatting
            for run in para.runs:
                r = p.add_run(run.text)
                r.bold = run.bold
                r.italic = run.italic
                r.underline = run.underline
                if run.font.size:
                    r.font.size = run.font.size
                if run.font.color.rgb:
                    r.font.color.rgb = run.font.color.rgb
        
        # Copy tables
        for table in doc.tables:
            # Create a new table with the same dimensions
            tbl = first_doc.add_table(rows=len(table.rows), cols=len(table.columns))
            
            # Copy cell contents
            for i, row in enumerate(table.rows):
                for j, cell in enumerate(row.cells):
                    if i < len(tbl.rows) and j < len(tbl.rows[i].cells):
                        # Copy cell text
                        tbl.rows[i].cells[j].text = cell.text
    
    # Save the combined document
    first_doc.save(output_path)

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
                    <label for="files">Select files to merge or upload a compressed archive:</label>
                    <input type="file" name="files" id="files" multiple required accept=".pdf,.docx,.doc,.zip,.rar">
                    <p class="note">Files will be merged in order based on "part" numbers in their names (e.g., file_part1.pdf, file_part2.pdf)</p>
                    <p class="note">You can also upload a ZIP or RAR file containing PDF or DOCX files to be merged</p>
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
    
    # Create a temporary directory to store uploaded files
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_file_paths = []
        extracted_files = []
        is_archive = False
        
        # Save uploaded files to temporary directory
        for file in files:
            file_ext = os.path.splitext(file.filename)[1].lower()
            temp_file_path = os.path.join(temp_dir, file.filename)
            
            with open(temp_file_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            
            # Check if the file is a compressed archive
            if file_ext in ['.zip', '.rar']:
                is_archive = True
                # Create a subdirectory for extraction
                extract_dir = os.path.join(temp_dir, f"extracted_{os.path.basename(file.filename)}")
                os.makedirs(extract_dir, exist_ok=True)
                
                # Extract files from the archive
                extracted = extract_compressed_file(temp_file_path, extract_dir)
                extracted_files.extend(extracted)
            else:
                temp_file_paths.append(temp_file_path)
        
        # If we extracted files from archives, use those instead
        if is_archive:
            # Filter for PDF and DOCX files only
            valid_extensions = ['.pdf', '.docx', '.doc']
            filtered_files = filter_files_by_extension(extracted_files, valid_extensions)
            
            if not filtered_files:
                raise HTTPException(status_code=400, detail="No PDF or DOCX/DOC files found in the archive")
            
            # Check if all files are of the same type
            file_extensions = [os.path.splitext(path)[1].lower() for path in filtered_files]
            unique_extensions = set(file_extensions)
            
            if len(unique_extensions) > 1:
                raise HTTPException(status_code=400, 
                                  detail="Files in the archive must be of the same type (either all PDF or all DOCX/DOC)")
            
            # Sort extracted files by part number in filename
            sorted_paths = sorted(filtered_files, key=lambda path: extract_part_number(os.path.basename(path)))
            temp_file_paths = sorted_paths
        else:
            # For direct uploads, check file types
            file_extensions = [os.path.splitext(file.filename)[1].lower() for file in files]
            
            if not all(ext in ['.pdf', '.docx', '.doc'] for ext in file_extensions):
                raise HTTPException(status_code=400, detail="Only PDF and DOCX/DOC files are supported")
            
            if len(set(file_extensions)) > 1:
                raise HTTPException(status_code=400, 
                                  detail="All files must be of the same type (either all PDF or all DOCX/DOC)")
            
            # Sort files by part number
            sorted_files = sort_files_by_part(files)
            temp_file_paths = [os.path.join(temp_dir, file.filename) for file in sorted_files]
        
        # Determine file type and set output path
        if not temp_file_paths:
            raise HTTPException(status_code=400, detail="No valid files found to merge")
            
        file_ext = os.path.splitext(temp_file_paths[0])[1].lower()
        output_path = f"uploads/{output_filename}{file_ext}"
        
        if file_ext == '.pdf':
            # For PDF files, merge them directly
            merge_pdf_files(temp_file_paths, output_path)
        else:  # .docx or .doc
            # For DOCX files, merge them
            merge_docx_files_custom(temp_file_paths, output_path)
    
    # Set the appropriate media type based on file extension
    if output_path.endswith(".pdf"):
        media_type = "application/pdf"
    else:  # .docx or .doc
        media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    
    return FileResponse(
        path=output_path,
        filename=os.path.basename(output_path),
        media_type=media_type
    )

@app.post("/api/merge/")
async def api_merge_files(
    file: Optional[UploadFile] = File(None),
    data: Optional[UploadFile] = File(None),
    archive: Optional[UploadFile] = File(None),
    output_filename: Optional[str] = Form("merged_document"),
    request: Request = None
):
    """API endpoint to merge files from a ZIP or RAR archive"""
    # Get the actual file from any of the possible parameter names
    actual_file = file or data or archive
    
    # Debug information
    debug_info = {
        "file_provided": file is not None,
        "data_provided": data is not None,
        "archive_provided": archive is not None,
        "output_filename": output_filename,
    }
    
    if not actual_file:
        # Try to get the file from form data directly
        form = await request.form()
        for key, value in form.items():
            if isinstance(value, UploadFile):
                actual_file = value
                debug_info["found_in_form"] = key
                break
    
    if not actual_file:
        raise HTTPException(status_code=400, detail=f"No file provided. Debug info: {debug_info}")
    
    # Check if the file is a compressed archive
    file_ext = os.path.splitext(actual_file.filename)[1].lower()
    if file_ext not in ['.zip', '.rar']:
        raise HTTPException(status_code=400, detail=f"Only ZIP or RAR archives are supported. Received file with extension: {file_ext}. Debug info: {debug_info}")
    
    # Create a temporary directory to store and extract files
    with tempfile.TemporaryDirectory() as temp_dir:
        # Save the uploaded archive
        temp_file_path = os.path.join(temp_dir, actual_file.filename)
        with open(temp_file_path, "wb") as buffer:
            shutil.copyfileobj(actual_file.file, buffer)
        
        # Create extraction directory
        extract_dir = os.path.join(temp_dir, "extracted")
        os.makedirs(extract_dir, exist_ok=True)
        
        # Extract files from the archive
        try:
            extracted_files = extract_compressed_file(temp_file_path, extract_dir)
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Error extracting archive: {str(e)}")
        
        # Filter for PDF and DOCX files only
        valid_extensions = ['.pdf', '.docx', '.doc']
        filtered_files = filter_files_by_extension(extracted_files, valid_extensions)
        
        if not filtered_files:
            raise HTTPException(status_code=400, detail="No PDF or DOCX/DOC files found in the archive")
        
        # Check if all files are of the same type
        file_extensions = [os.path.splitext(path)[1].lower() for path in filtered_files]
        unique_extensions = set(file_extensions)
        
        if len(unique_extensions) > 1:
            raise HTTPException(status_code=400, 
                              detail="Files in the archive must be of the same type (either all PDF or all DOCX/DOC)")
        
        # Sort extracted files by part number in filename
        sorted_paths = sorted(filtered_files, key=lambda path: extract_part_number(os.path.basename(path)))
        
        # Determine file type and set output path
        file_ext = os.path.splitext(sorted_paths[0])[1].lower()
        output_path = f"uploads/{output_filename}{file_ext}"
        
        if file_ext == '.pdf':
            # For PDF files, merge them directly
            merge_pdf_files(sorted_paths, output_path)
        else:  # .docx or .doc
            # For DOCX files, merge them
            merge_docx_files_custom(sorted_paths, output_path)
        
        # Set the appropriate media type based on file extension
        if output_path.endswith(".pdf"):
            media_type = "application/pdf"
        else:  # .docx or .doc
            media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        
        return FileResponse(
            path=output_path,
            filename=os.path.basename(output_path),
            media_type=media_type
        )

if __name__ == "__main__":
    # Get port from environment variable or default to 8000
    port = int(os.environ.get("PORT", 8000))
    # Use 0.0.0.0 for production, 127.0.0.1 for local development
    host = os.environ.get("HOST", "127.0.0.1")
    uvicorn.run("main:app", host=host, port=port, reload=True)
