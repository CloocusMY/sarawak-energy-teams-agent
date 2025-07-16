import os
import requests
import msal
from dotenv import load_dotenv
from PyPDF2 import PdfReader
import io
from typing import List, Dict
import tiktoken
 
# Load environment variables
load_dotenv()

async def get_doc_data_for_folder(folder_name: str, embeddings) -> List[Dict]:
    print(f"📂 Fetching documents from SharePoint folder: {folder_name}")

    site_id = os.getenv("SHAREPOINT_SITE_ID")
    drive_id = os.getenv("SHAREPOINT_DRIVE_ID")
    graph_api_endpoint = "https://graph.microsoft.com/v1.0"
    access_token = get_sharepoint_access_token()
    headers = {"Authorization": f"Bearer {access_token}"}

    documents = []

    folder_url = f"{graph_api_endpoint}/sites/{site_id}/drives/{drive_id}/root:/{folder_name}:/children"
    response = requests.get(folder_url, headers=headers)

    if response.status_code != 200:
        print(f"❌ Failed to access folder {folder_name}: {response.status_code}")
        return documents

    files = response.json().get("value", [])
    if not files:
        print(f"⚠️ No files found in folder {folder_name}")
        return documents

    for idx, file in enumerate(files, start=1):
        if file.get("folder"):
            continue  # skip subfolders

        file_id = file.get("id")
        file_name = file.get("name")
        download_url = f"{graph_api_endpoint}/sites/{site_id}/drives/{drive_id}/items/{file_id}/content"
        file_response = requests.get(download_url, headers=headers)

        if file_response.status_code != 200:
            print(f"❌ Failed to download {file_name}")
            continue

        # Handle content
        if file_name.lower().endswith(".pdf"):
            content = extract_text_from_pdf(file_response.content, file_name)
        else:
            try:
                content = file_response.content.decode("utf-8")
            except UnicodeDecodeError:
                print(f"⚠️ Could not decode {file_name} as text.")
                content = f"{file_name}: unable to decode."

        max_tokens = 8000
        overlap_tokens = 400
        chunks = split_text_with_overlap(content, max_tokens, overlap_tokens)

        for chunk_idx, chunk in enumerate(chunks, start=1):
            try:
                embedding = await get_embedding_vector(chunk, embeddings)
            except Exception as e:
                print(f"❌ Embedding failed for {file_name} chunk {chunk_idx}: {e}")
                continue

            documents.append({
                "docId": f"{folder_name}_{idx}_{chunk_idx}",
                "docTitle": f"{folder_name}_{file_name}",
                "description": chunk,
                "descriptionVector": embedding
            })

    print(f"✅ Fetched {len(documents)} docs from {folder_name}")
    return documents

tokenizer = tiktoken.get_encoding("cl100k_base")

def split_text_with_overlap(text: str, max_tokens: int, overlap_tokens: int):
    tokens = tokenizer.encode(text)
    chunks = []
    start_index = 0

    while start_index < len(tokens):
        end_index = min(start_index + max_tokens, len(tokens))
        chunk = tokens[start_index:end_index]
        chunks.append(tokenizer.decode(chunk))
        # move forward with overlap
        start_index += max_tokens - overlap_tokens
    return chunks


def get_sharepoint_access_token():
    """Get access token for SharePoint"""
    tenant_id = os.getenv("TENANT_ID")
    client_id = os.getenv("CLIENT_ID")
    client_secret = os.getenv("CLIENT_SECRET")
   
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    scopes = ["https://graph.microsoft.com/.default"]
   
    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        authority=authority,
        client_credential=client_secret
    )
   
    result = app.acquire_token_for_client(scopes=scopes)
   
    if "access_token" not in result:
        raise Exception(f"❌ SharePoint authentication failed: {result.get('error_description', result.get('error'))}")
   
    return result["access_token"]
 
 
def get_file_from_sharepoint(headers, graph_api_endpoint, site_id, drive_id, folder_name):
    """Get file content from SharePoint folder"""
    try:
        folder_url = f"{graph_api_endpoint}/sites/{site_id}/drives/{drive_id}/root:/{folder_name}:/children"
        response = requests.get(folder_url, headers=headers)
       
        if response.status_code != 200:
            print(f"❌ Failed to access folder {folder_name}: {response.status_code}")
            return None
           
        files = response.json().get("value", [])
       
        for file in files:
            if not file.get("folder"):
                file_id = file["id"]
                file_name = file["name"]
               
                # Download file content
                download_url = f"{graph_api_endpoint}/sites/{site_id}/drives/{drive_id}/items/{file_id}/content"
                file_response = requests.get(download_url, headers=headers)
               
                if file_response.status_code == 200:
                    if file_name.lower().endswith('.pdf'):
                        try:
                            content = extract_text_from_pdf(file_response.content, file_name)

                            if content:
                                print(f"✅ Successfully extracted text from PDF: {file_name} from {folder_name}")
                                return content
                            else:
                                print(f"⚠️ Could not extract text from PDF: {file_name}")
                                # Fallback: return filename and basic info
                                return f"Document: {file_name} from {folder_name} folder. This is a PDF document that requires text extraction."
                        except Exception as e:
                            print(f"❌ Error processing PDF {file_name}: {str(e)}")
                            return f"Document: {file_name} from {folder_name} folder. PDF processing failed."
                    else:
                        try:
                            content = file_response.content.decode('utf-8')
                            print(f"✅ Successfully loaded: {file_name} from {folder_name}")
                            return content
                        except UnicodeDecodeError:
                            print(f"⚠️ Could not decode {file_name} as text")
                            return f"Document: {file_name} from {folder_name} folder. File could not be decoded as text."
                       
        print(f"⚠️ No suitable files found in {folder_name}")
        return None
       
    except Exception as e:
        print(f"❌ Error accessing {folder_name}: {str(e)}")
        return None
 
 
def extract_text_from_pdf(pdf_content, file_name):
    """Extract text from PDF"""
    try:
        reader = PdfReader(io.BytesIO(pdf_content))
        text = ""
        for page in reader.pages:
            text += page.extract_text() or ""
        if not text.strip():
            raise ValueError("No text found in PDF")
        return text.strip()
    except Exception as e:
        print(f"❌ Error extracting text from PDF {file_name}: {str(e)}")
        return None
     
 
async def get_embedding_vector(text: str, embeddings):
    """Generate embedding vector - keeping original functionality"""
    result = await embeddings.create_embeddings(text)
    if (result.status != 'success' or not result.output):
        if result.status == 'error':
            raise Exception(f"Failed to generate embeddings for description: <{text[:200]+'...'}>\n\nError: {result.output}")
        raise Exception(f"Failed to generate embeddings for description: <{text[:200]+'...'}>")
   
    return result.output[0]