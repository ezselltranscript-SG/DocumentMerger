from fastapi import FastAPI, File, UploadFile, HTTPException, Form, Request
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from typing import List, Optional
from pypdf import PdfWriter, PdfReader
from docx import Document
from docx.enum.text import WD_BREAK
from docxcompose.composer import Composer

import os
import re
import shutil
import tempfile
import zipfile
import rarfile
import patoolib
import logging

# Configurar logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(title="PDF and DOCX Merger API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

os.makedirs("uploads", exist_ok=True)

# Utilidades
def extract_part_number(filename: str) -> int:
    match = re.search(r'part(\d+)', filename.lower())
    return int(match.group(1)) if match else 0

def sort_files_by_part(files: List[str]) -> List[str]:
    return sorted(files, key=lambda f: extract_part_number(os.path.basename(f)))

def extract_compressed_file(file_path: str, extract_dir: str) -> List[str]:
    extracted_files = []
    try:
        try:
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
                return [os.path.join(extract_dir, name) for name in zip_ref.namelist() if not name.endswith('/')]
        except zipfile.BadZipFile:
            pass

        try:
            rarfile.UNRAR_TOOL = 'unrar'
            with rarfile.RarFile(file_path) as rar_ref:
                rar_ref.extractall(extract_dir)
                return [os.path.join(extract_dir, name) for name in rar_ref.namelist() if not rar_ref.getinfo(name).isdir()]
        except rarfile.NotRarFile:
            pass

        patoolib.extract_archive(file_path, outdir=extract_dir)
        for root, _, files in os.walk(extract_dir):
            for file in files:
                extracted_files.append(os.path.join(root, file))

        return extracted_files
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error extracting archive: {str(e)}")

def filter_files_by_extension(file_paths: List[str], extensions: List[str]) -> List[str]:
    return [p for p in file_paths if os.path.splitext(p)[1].lower() in extensions]

def merge_pdf_files(file_paths: List[str], output_path: str) -> None:
    merger = PdfWriter()
    for path in file_paths:
        reader = PdfReader(path)
        for page in reader.pages:
            merger.add_page(page)
    with open(output_path, "wb") as out:
        merger.write(out)

def merge_docx_files_raw(file_paths: List[str], output_path: str) -> None:
    """
    Fusiona archivos DOCX sin modificar su contenido interno.
    Cada documento se inserta completo con su propio formato, estructura y encabezados.
    """
    if not file_paths:
        raise HTTPException(status_code=400, detail="No DOCX files provided")
    
    if len(file_paths) == 1:
        # Si solo hay un archivo, simplemente copiarlo
        shutil.copy(file_paths[0], output_path)
        return
    
    try:
        # Crear un directorio temporal para trabajar
        with tempfile.TemporaryDirectory() as temp_dir:
            # Extraer el primer documento como base
            base_extract_dir = os.path.join(temp_dir, "base_doc")
            os.makedirs(base_extract_dir, exist_ok=True)
            
            with zipfile.ZipFile(file_paths[0], 'r') as zip_ref:
                zip_ref.extractall(base_extract_dir)
            
            # Leer el contenido del document.xml
            document_xml_path = os.path.join(base_extract_dir, "word", "document.xml")
            with open(document_xml_path, 'r', encoding='utf-8') as f:
                base_content = f.read()
            
            # Encontrar la posición donde insertar el contenido de los otros documentos
            # (justo antes del cierre del cuerpo del documento)
            insert_pos = base_content.rfind("</w:body>")
            if insert_pos == -1:
                raise ValueError("No se pudo encontrar el final del cuerpo del documento base")
            
            # Preparar el contenido combinado
            combined_content = base_content[:insert_pos]
            
            # Para cada documento adicional
            for i, file_path in enumerate(file_paths[1:], 1):
                logger.info(f"Procesando documento {i}: {os.path.basename(file_path)}")
                
                # Extraer el documento actual
                doc_extract_dir = os.path.join(temp_dir, f"doc_{i}")
                os.makedirs(doc_extract_dir, exist_ok=True)
                
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    zip_ref.extractall(doc_extract_dir)
                
                # Leer el contenido del document.xml
                doc_xml_path = os.path.join(doc_extract_dir, "word", "document.xml")
                with open(doc_xml_path, 'r', encoding='utf-8') as f:
                    doc_content = f.read()
                
                # Extraer solo el contenido del cuerpo (entre <w:body> y </w:body>)
                body_start = doc_content.find("<w:body>") + len("<w:body>")
                body_end = doc_content.rfind("</w:body>")
                
                if body_start == -1 or body_end == -1:
                    raise ValueError(f"No se pudo extraer el cuerpo del documento {i}")
                
                body_content = doc_content[body_start:body_end]
                
                # Agregar un salto de página antes del contenido
                page_break = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
                
                # Agregar el contenido al documento combinado
                combined_content += page_break + body_content
                
                # Copiar todos los archivos de word/ excepto document.xml
                for item in os.listdir(os.path.join(doc_extract_dir, "word")):
                    if item != "document.xml":
                        src = os.path.join(doc_extract_dir, "word", item)
                        dst = os.path.join(base_extract_dir, "word", f"{i}_{item}")
                        
                        # Copiar el archivo
                        if os.path.isfile(src):
                            shutil.copy(src, dst)
                        elif os.path.isdir(src):
                            shutil.copytree(src, dst)
                
                # Actualizar las relaciones
                rels_dir = os.path.join(base_extract_dir, "word", "_rels")
                os.makedirs(rels_dir, exist_ok=True)
                
                base_rels_path = os.path.join(rels_dir, "document.xml.rels")
                doc_rels_path = os.path.join(doc_extract_dir, "word", "_rels", "document.xml.rels")
                
                if os.path.exists(doc_rels_path) and os.path.exists(base_rels_path):
                    with open(base_rels_path, 'r', encoding='utf-8') as f:
                        base_rels = f.read()
                    
                    with open(doc_rels_path, 'r', encoding='utf-8') as f:
                        doc_rels = f.read()
                    
                    # Extraer todas las relaciones del documento actual
                    import re
                    rel_pattern = r'<Relationship [^>]+>'
                    doc_rel_matches = re.findall(rel_pattern, doc_rels)
                    
                    # Modificar los IDs y targets para evitar conflictos
                    modified_rels = []
                    for rel in doc_rel_matches:
                        # Cambiar el ID para evitar conflictos
                        rel = rel.replace('Id="rId', f'Id="rId{i*1000}')
                        
                        # Cambiar el Target si apunta a un archivo que hemos renombrado
                        if 'Target="word/' in rel:
                            rel = rel.replace('Target="word/', f'Target="word/{i}_')
                        
                        modified_rels.append(rel)
                    
                    # Agregar las relaciones modificadas al documento base
                    base_rels = base_rels.replace('</Relationships>', 
                                                 ''.join(modified_rels) + '</Relationships>')
                    
                    # Guardar las relaciones actualizadas
                    with open(base_rels_path, 'w', encoding='utf-8') as f:
                        f.write(base_rels)
            
            # Completar el documento combinado
            combined_content += "</w:body></w:document>"
            
            # Guardar el documento combinado
            with open(document_xml_path, 'w', encoding='utf-8') as f:
                f.write(combined_content)
            
            # Crear un nuevo archivo DOCX con el contenido combinado
            with zipfile.ZipFile(output_path, 'w') as zip_out:
                for root, _, files in os.walk(base_extract_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arc_name = os.path.relpath(file_path, base_extract_dir)
                        zip_out.write(file_path, arc_name)
            
            logger.info(f"Documento combinado guardado en {output_path}")
    
    except Exception as e:
        logger.error(f"Error al fusionar documentos: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error al fusionar documentos: {str(e)}")

@app.post("/api/merge/")
async def api_merge_files(
    file: Optional[UploadFile] = File(None),
    data: Optional[UploadFile] = File(None),
    archive: Optional[UploadFile] = File(None),
    output_filename: Optional[str] = Form("merged_document"),
    request: Request = None
):
    actual_file = file or data or archive
    if not actual_file:
        form = await request.form()
        for value in form.values():
            if isinstance(value, UploadFile):
                actual_file = value
                break
    if not actual_file:
        raise HTTPException(status_code=400, detail="No file provided.")

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = os.path.join(temp_dir, "archive_input")
        with open(temp_path, "wb") as f:
            shutil.copyfileobj(actual_file.file, f)

        extract_dir = os.path.join(temp_dir, "extracted")
        os.makedirs(extract_dir, exist_ok=True)
        extracted = extract_compressed_file(temp_path, extract_dir)

        valid_exts = ['.pdf', '.docx']
        filtered = filter_files_by_extension(extracted, valid_exts)
        if not filtered:
            raise HTTPException(status_code=400, detail="No valid PDF or DOCX files found.")

        sorted_files = sort_files_by_part(filtered)
        ext = os.path.splitext(sorted_files[0])[1].lower()
        output_path = f"uploads/{output_filename}{ext}"

        if ext == ".pdf":
            merge_pdf_files(sorted_files, output_path)
            media_type = "application/pdf"
        elif ext == ".docx":
            merge_docx_files_raw(sorted_files, output_path)
            media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        else:
            raise HTTPException(status_code=400, detail="Unsupported file type.")

        return FileResponse(path=output_path, filename=os.path.basename(output_path), media_type=media_type)

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 10000))
    uvicorn.run(app, host="0.0.0.0", port=port)
