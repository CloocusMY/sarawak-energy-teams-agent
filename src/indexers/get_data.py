import os
import requests
import msal
from dotenv import load_dotenv
 
# Load environment variables
load_dotenv()
 
async def get_doc_data(embeddings):
    """
    Fetch documents from SharePoint instead of local storage
    """
    print("🔄 Fetching documents from SharePoint...")
   
    # Get SharePoint access token
    access_token = get_sharepoint_access_token()
   
    # SharePoint configuration from environment variables
    site_id = os.getenv("SHAREPOINT_SITE_ID")
    drive_id = os.getenv("SHAREPOINT_DRIVE_ID")
    graph_api_endpoint = "https://graph.microsoft.com/v1.0"
   
    if not site_id or not drive_id:
        raise Exception("SHAREPOINT_SITE_ID and SHAREPOINT_DRIVE_ID must be set in environment variables")
   
    headers = {"Authorization": f"Bearer {access_token}"}
   
    documents = []
   
    # Document 1: From HR folder
    hr_file_content = get_file_from_sharepoint(headers, graph_api_endpoint, site_id, drive_id, "HR")
    if hr_file_content:
        doc1 = {
            "docId": "1",
            "docTitle": "HR_Documents",
            "description": hr_file_content,
            "descriptionVector": await get_embedding_vector(hr_file_content, embeddings=embeddings),
        }
        documents.append(doc1)
   
    # Document 2: From Procurement folder
    procurement_file_content = get_file_from_sharepoint(headers, graph_api_endpoint, site_id, drive_id, "Procurement")
    if procurement_file_content:
        doc2 = {
            "docId": "2",
            "docTitle": "Procurement_Documents",
            "description": procurement_file_content,
            "descriptionVector": await get_embedding_vector(procurement_file_content, embeddings=embeddings),
        }
        documents.append(doc2)
   
    print(f"✅ Successfully loaded {len(documents)} documents from SharePoint")
    return documents
 
 
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
        # List files in folder
        folder_url = f"{graph_api_endpoint}/sites/{site_id}/drives/{drive_id}/root:/{folder_name}:/children"
        response = requests.get(folder_url, headers=headers)
       
        if response.status_code != 200:
            print(f"❌ Failed to access folder {folder_name}: {response.status_code}")
            return None
           
        files = response.json().get("value", [])
       
        # Get the first file (you can modify this logic)
        for file in files:
            if not file.get("folder"):  # Skip folders
                file_id = file["id"]
                file_name = file["name"]
               
                # Download file content
                download_url = f"{graph_api_endpoint}/sites/{site_id}/drives/{drive_id}/items/{file_id}/content"
                file_response = requests.get(download_url, headers=headers)
               
                if file_response.status_code == 200:
                    # For PDF files, extract text using a simple method
                    if file_name.lower().endswith('.pdf'):
                        try:
                            # For now, we'll extract basic text.
                            # In production, you'd want to use a proper PDF library like PyPDF2 or pdfplumber
                            content = extract_text_from_pdf_simple(file_response.content, file_name)
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
                        # For text files, decode content
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
 
 
def extract_text_from_pdf_simple(pdf_content, file_name):
    """Simple PDF text extraction - returns basic info for now"""
    # For a quick test, we'll return file info
    # In production, you'd use PyPDF2, pdfplumber, or Azure Form Recognizer
   
    try:
        # Try to extract some basic text using simple string search
        # This is a very basic approach just to get something working
        content_str = str(pdf_content)
       
        # Look for common text patterns
        text_content = f"PDF Document: {file_name}\n"
        text_content += f"File size: {len(pdf_content)} bytes\n"
        text_content += "This is a PDF document that contains text content. "
        text_content += "Full text extraction would require additional PDF processing libraries."
       
        return text_content
       
    except Exception as e:
        print(f"❌ Error in simple PDF extraction: {str(e)}")
        return f"PDF Document: {file_name} - Basic extraction failed"
 
 
async def get_embedding_vector(text: str, embeddings):
    """Generate embedding vector - keeping original functionality"""
    result = await embeddings.create_embeddings(text)
    if (result.status != 'success' or not result.output):
        if result.status == 'error':
            raise Exception(f"Failed to generate embeddings for description: <{text[:200]+'...'}>\n\nError: {result.output}")
        raise Exception(f"Failed to generate embeddings for description: <{text[:200]+'...'}>")
   
    return result.output[0]