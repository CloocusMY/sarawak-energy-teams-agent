import asyncio, os, argparse
from dataclasses import dataclass
from typing import List, Optional

from azure.core.credentials import AzureKeyCredential
from azure.core.exceptions import ServiceRequestError
from azure.search.documents import SearchClient
from azure.search.documents.indexes import SearchIndexClient
from azure.search.documents.indexes.models import (
    SearchIndex,
    SimpleField,
    SearchableField,
    SearchField,
    SearchFieldDataType,
    ComplexField,
    CorsOptions,
    VectorSearch,
    VectorSearchProfile,
    HnswAlgorithmConfiguration
)
from teams.ai.embeddings import AzureOpenAIEmbeddings, AzureOpenAIEmbeddingsOptions

from get_data import get_doc_data, get_doc_data_for_folder

from dotenv import load_dotenv

# FIXED: Get the correct path to the project root and env file
script_dir = os.path.dirname(os.path.abspath(__file__))  # src/indexers/
src_dir = os.path.dirname(script_dir)                   # src/
project_root = os.path.dirname(src_dir)                 # SE-BoK/
env_file_path = os.path.join(project_root, 'env', '.env.local.user')

print(f"Script location: {script_dir}")
print(f"Project root: {project_root}")
print(f"Loading environment from: {env_file_path}")
print(f"Environment file exists: {os.path.exists(env_file_path)}")

load_dotenv(env_file_path, override=True)

@dataclass
class Doc:
    docId: Optional[str] = None
    docTitle: Optional[str] = None
    description: Optional[str] = None
    descriptionVector: Optional[List[float]] = None

async def upsert_documents(client: SearchClient, documents: list[Doc]):
    return client.merge_or_upload_documents(documents)

async def create_index_if_not_exists(client: SearchIndexClient, name: str):
    doc_index = SearchIndex(
        name=name,
        fields = [
            SimpleField(name="docId", type=SearchFieldDataType.String, key=True),
            SimpleField(name="docTitle", type=SearchFieldDataType.String),
            SearchableField(name="description", type=SearchFieldDataType.String, searchable=True),
            SearchField(name="descriptionVector", type=SearchFieldDataType.Collection(SearchFieldDataType.Single), hidden=False, searchable=True, vector_search_dimensions=1536, vector_search_profile_name='my-vector-config'),
        ],
        scoring_profiles=[],
        cors_options=CorsOptions(allowed_origins=["*"]),
        vector_search = VectorSearch(
            profiles=[VectorSearchProfile(name="my-vector-config", algorithm_configuration_name="my-algorithms-config")],
            algorithms=[HnswAlgorithmConfiguration(name="my-algorithms-config")],
        )
    )

    client.create_or_update_index(doc_index)

def load_keys_from_args():
    parser = argparse.ArgumentParser(description='Load keys from command input parameters.')
    parser.add_argument('--api-key', type=str, required=True, help='Azure OpenAI API key for authentication')
    parser.add_argument('--ai-search-key', type=str, required=True, help='AI Search key for authentication')
    args = parser.parse_args()
    return args

async def setup(search_api_key, search_api_endpoint, args):
    # Two indexes (adjust names as you like)
    folder_index_map = {
        "HR": "saeshav-hr-index",
        "Procurement": "saeshav-procurement-index"
    }

    credentials = AzureKeyCredential(search_api_key)
    search_index_client = SearchIndexClient(search_api_endpoint, credentials)

    # Create indexes if they don't exist
    for folder, index_name in folder_index_map.items():
        print(f"🔧 Creating or updating index: {index_name}")
        await create_index_if_not_exists(search_index_client, index_name)

    await asyncio.sleep(5)  # Give some time after creation

    # Initialize embeddings once
    embeddings = AzureOpenAIEmbeddings(AzureOpenAIEmbeddingsOptions(
        azure_api_key=args.api_key,
        azure_endpoint=os.getenv('AZURE_OPENAI_ENDPOINT'),
        azure_deployment=os.getenv('AZURE_OPENAI_EMBEDDING_DEPLOYMENT')
    ))

    # Fetch documents folder-by-folder and upsert to respective index
    for folder, index_name in folder_index_map.items():
        print(f"📥 Fetching documents for folder: {folder}")
        data = await get_doc_data_for_folder(folder_name=folder, embeddings=embeddings)

        search_client = SearchClient(search_api_endpoint, index_name, credentials)
        if data:
            print(f"📤 Uploading {len(data)} documents to index: {index_name}")
            await upsert_documents(search_client, data)
        else:
            print(f"⚠️ No documents found for folder: {folder}")

    print("✅ All folders processed successfully!")
    
args = load_keys_from_args()
search_api_key = args.ai_search_key
search_api_endpoint = os.getenv('AZURE_SEARCH_ENDPOINT')

# Debug: Print what we got from environment
print(f"\nEnvironment variables loaded:")
print(f"AZURE_SEARCH_ENDPOINT: {search_api_endpoint}")
print(f"AZURE_OPENAI_ENDPOINT: {os.getenv('AZURE_OPENAI_ENDPOINT')}")
print(f"AZURE_OPENAI_EMBEDDING_DEPLOYMENT: {os.getenv('AZURE_OPENAI_EMBEDDING_DEPLOYMENT')}")

if not search_api_endpoint:
    print("❌ ERROR: AZURE_SEARCH_ENDPOINT is not set!")
    print(f"Check your environment file at: {env_file_path}")
    exit(1)

try:
    asyncio.run(setup(search_api_key, search_api_endpoint, args))
    print("✅ Setup finished successfully!")
except ServiceRequestError as e:
    print(f"❌ Setup index failed due to ServiceRequestError: {e.message}")
    print(f"Please check your keys, models and endpoints in {env_file_path}")
except Exception as e:
    print(f"❌ Setup failed with error: {e}")