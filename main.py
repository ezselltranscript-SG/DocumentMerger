import os
import re
import shutil
import tempfile
import zipfile
import rarfile
import patoolib
from typing import List, Optional, Tuple
from fastapi import FastAPI, File, UploadFile, HTTPException, Form, Request
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import uvicorn
from pypdf import PdfWriter, PdfReader
import docx

app = FastAPI(title="PDF and DOCX Merger API")

# Configurar CORS para permitir solicitudes desde cualquier origen
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

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
    extracted_files = []
    try:
        # Primero intentamos con zipfile
        try:
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
                extracted_files = [os.path.join(extract_dir, name) for name in zip_ref.namelist() if not name.endswith('/')]
                return extracted_files
        except zipfile.BadZipFile:
            # No es un archivo ZIP, intentamos con RAR
            pass
        
        # Intentamos con rarfile
        try:
            rarfile.UNRAR_TOOL = 'unrar'
            with rarfile.RarFile(file_path) as rar_ref:
                rar_ref.extractall(extract_dir)
                extracted_files = [os.path.join(extract_dir, name) for name in rar_ref.namelist() if not rar_ref.getinfo(name).isdir()]
                return extracted_files
        except rarfile.NotRarFile:
            # No es un archivo RAR, intentamos con patoolib como último recurso
            pass
        
        # Último recurso: intentamos con patoolib que maneja varios formatos
        patoolib.extract_archive(file_path, outdir=extract_dir)
        for root, _, files in os.walk(extract_dir):
            for file in files:
                extracted_files.append(os.path.join(root, file))
        
        if not extracted_files:
            raise ValueError("No se pudieron extraer archivos del archivo comprimido")
            
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
    
    # Asumimos que el archivo es un archivo comprimido válido y lo procesamos directamente
    # sin verificar la extensión, ya que n8n puede estar enviando archivos sin extensión
    
    # Create a temporary directory to store and extract files
    with tempfile.TemporaryDirectory() as temp_dir:
        # Save the uploaded archive with a nombre genérico para evitar problemas con extensiones
        temp_file_path = os.path.join(temp_dir, "uploaded_archive")
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
    port = int(os.environ.get("PORT", 10000))
    uvicorn.run(app, host="0.0.0.0", port=port)
