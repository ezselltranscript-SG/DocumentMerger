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

def merge_docx_with_sections(file_paths: List[str], output_path: str) -> None:
    """
    Fusiona archivos DOCX manteniendo cada documento como una sección independiente
    con su propio encabezado, formato y estructura.
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
            # Directorio para el documento combinado
            combined_dir = os.path.join(temp_dir, "combined")
            os.makedirs(combined_dir, exist_ok=True)
            os.makedirs(os.path.join(combined_dir, "word"), exist_ok=True)
            os.makedirs(os.path.join(combined_dir, "word", "_rels"), exist_ok=True)
            
            # Inicializar el contenido del documento combinado
            combined_document = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            combined_document += '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            combined_document += 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'
            combined_document += '<w:body>\n'
            
            # Inicializar las relaciones
            combined_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            combined_rels += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
            
            # Contador para IDs de relación
            rel_id_counter = 1
            
            # Procesar cada documento
            for doc_index, file_path in enumerate(file_paths):
                logger.info(f"Procesando documento {doc_index}: {os.path.basename(file_path)}")
                
                # Directorio para extraer el documento actual
                doc_dir = os.path.join(temp_dir, f"doc_{doc_index}")
                os.makedirs(doc_dir, exist_ok=True)
                
                # Extraer el documento
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    zip_ref.extractall(doc_dir)
                
                # Agregar salto de página entre documentos (excepto el primero)
                if doc_index > 0:
                    combined_document += '<w:p>\n<w:r>\n<w:br w:type="page"/>\n</w:r>\n</w:p>\n'
                
                # Agregar salto de sección para cada documento
                if doc_index > 0:
                    # Agregar un salto de sección que inicia una nueva sección
                    combined_document += '<w:sectPr>\n'
                    combined_document += '  <w:type w:val="nextPage"/>\n'
                    combined_document += '</w:sectPr>\n'
                
                # Leer el contenido del documento
                doc_xml_path = os.path.join(doc_dir, "word", "document.xml")
                with open(doc_xml_path, 'r', encoding='utf-8') as f:
                    doc_content = f.read()
                
                # Extraer el cuerpo del documento
                body_start = doc_content.find("<w:body>") + len("<w:body>")
                body_end = doc_content.rfind("</w:body>")
                
                if body_start == -1 or body_end == -1:
                    raise ValueError(f"No se pudo extraer el cuerpo del documento {doc_index}")
                
                body_content = doc_content[body_start:body_end]
                
                # Encontrar y extraer la configuración de sección (sectPr)
                sect_pr_start = body_content.rfind("<w:sectPr")
                sect_pr_end = body_content.rfind("</w:sectPr>") + len("</w:sectPr>")
                
                if sect_pr_start != -1 and sect_pr_end != -1:
                    sect_pr = body_content[sect_pr_start:sect_pr_end]
                    # Eliminar la configuración de sección del cuerpo para agregarla después
                    body_content = body_content[:sect_pr_start] + body_content[sect_pr_end:]
                else:
                    sect_pr = "<w:sectPr></w:sectPr>"
                
                # Copiar los encabezados y pies de página
                word_dir = os.path.join(doc_dir, "word")
                for item in os.listdir(word_dir):
                    # Copiar encabezados, pies de página y estilos
                    if item.startswith("header") or item.startswith("footer") or item == "styles.xml":
                        src_path = os.path.join(word_dir, item)
                        dst_name = f"{doc_index}_{item}"
                        dst_path = os.path.join(combined_dir, "word", dst_name)
                        
                        # Copiar el archivo
                        shutil.copy(src_path, dst_path)
                        
                        # Actualizar las referencias en la configuración de sección
                        if item.startswith("header") or item.startswith("footer"):
                            # Crear una nueva relación para este encabezado/pie de página
                            rel_id = f"rId{rel_id_counter}"
                            rel_id_counter += 1
                            
                            # Determinar el tipo de relación
                            rel_type = "header" if item.startswith("header") else "footer"
                            
                            # Agregar la relación al archivo de relaciones
                            combined_rels += f'<Relationship Id="{rel_id}" '
                            combined_rels += f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{rel_type}" '
                            combined_rels += f'Target="word/{dst_name}"/>\n'
                            
                            # Actualizar la referencia en sectPr
                            old_ref = f'r:id="rId[0-9]+"'
                            new_ref = f'r:id="{rel_id}"'
                            
                            # Buscar referencias al encabezado/pie de página en la configuración de sección
                            import re
                            if re.search(f'<w:{rel_type}[^>]*{old_ref}', sect_pr):
                                sect_pr = re.sub(f'(<w:{rel_type}[^>]*){old_ref}', f'\\1{new_ref}', sect_pr)
                
                # Copiar las relaciones del documento
                doc_rels_path = os.path.join(doc_dir, "word", "_rels")
                if os.path.exists(doc_rels_path):
                    for item in os.listdir(doc_rels_path):
                        if item != "document.xml.rels":
                            src_path = os.path.join(doc_rels_path, item)
                            dst_path = os.path.join(combined_dir, "word", "_rels", f"{doc_index}_{item}")
                            shutil.copy(src_path, dst_path)
                
                # Copiar otros archivos necesarios (imágenes, etc.)
                for root, _, files in os.walk(word_dir):
                    for file in files:
                        if not (file.startswith("header") or file.startswith("footer") or 
                                file == "document.xml" or file == "styles.xml"):
                            rel_path = os.path.relpath(os.path.join(root, file), word_dir)
                            src_path = os.path.join(root, file)
                            dst_path = os.path.join(combined_dir, "word", f"{doc_index}_{rel_path}")
                            
                            # Crear directorios si es necesario
                            os.makedirs(os.path.dirname(dst_path), exist_ok=True)
                            
                            # Copiar el archivo
                            if os.path.isfile(src_path):
                                shutil.copy(src_path, dst_path)
                
                # Agregar el contenido del cuerpo al documento combinado
                combined_document += body_content
                
                # Agregar la configuración de sección al final de cada documento
                # excepto para el último documento que tendrá su propia configuración al final
                if doc_index < len(file_paths) - 1:
                    combined_document += sect_pr + '\n'
                else:
                    # Para el último documento, guardar la configuración para agregarla al final
                    last_sect_pr = sect_pr
            
            # Finalizar el documento combinado
            combined_document += last_sect_pr + '\n'
            combined_document += '</w:body>\n</w:document>'
            
            # Finalizar el archivo de relaciones
            combined_rels += '</Relationships>'
            
            # Guardar el documento combinado
            with open(os.path.join(combined_dir, "word", "document.xml"), 'w', encoding='utf-8') as f:
                f.write(combined_document)
            
            # Guardar el archivo de relaciones
            with open(os.path.join(combined_dir, "word", "_rels", "document.xml.rels"), 'w', encoding='utf-8') as f:
                f.write(combined_rels)
            
            # Copiar archivos necesarios del primer documento para la estructura básica
            base_files = ["[Content_Types].xml", "_rels/.rels"]
            for file_path in file_paths:
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    for base_file in base_files:
                        try:
                            zip_ref.extract(base_file, combined_dir)
                            break  # Si se extrajo correctamente, salir del bucle
                        except KeyError:
                            continue  # Si no existe, probar con el siguiente documento
            
            # Crear el archivo DOCX final
            with zipfile.ZipFile(output_path, 'w') as zip_out:
                for root, _, files in os.walk(combined_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arc_name = os.path.relpath(file_path, combined_dir)
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
            merge_docx_with_sections(sorted_files, output_path)
            media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        else:
            raise HTTPException(status_code=400, detail="Unsupported file type.")

        return FileResponse(path=output_path, filename=os.path.basename(output_path), media_type=media_type)

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 10000))
    uvicorn.run(app, host="0.0.0.0", port=port)
