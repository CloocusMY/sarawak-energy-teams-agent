from dataclasses import dataclass
from typing import Optional, List
from azure.search.documents.indexes.models import _edm as EDM
from azure.search.documents.models import VectorQuery, VectorizedQuery
from teams.ai.embeddings import AzureOpenAIEmbeddings, AzureOpenAIEmbeddingsOptions
from teams.state.memory import Memory
from teams.state.state import TurnContext
from teams.ai.tokenizers import Tokenizer
from teams.ai.data_sources import DataSource
from state_store import storage

from config import Config

async def get_embedding_vector(text: str):
    embeddings = AzureOpenAIEmbeddings(AzureOpenAIEmbeddingsOptions(
        azure_api_key=Config.AZURE_OPENAI_API_KEY,
        azure_endpoint=Config.AZURE_OPENAI_ENDPOINT,
        azure_deployment=Config.AZURE_OPENAI_EMBEDDING_DEPLOYMENT
    ))
    try:
        result = await embeddings.create_embeddings(text)
    except Exception as e:
        raise Exception(f"Error generating embeddings: {str(e)}")
    if (result.status != 'success' or not result.output):
        raise Exception(f"Failed to generate embeddings for description: {text}")
    
    return result.output[0]

@dataclass
class Doc:
    docId: Optional[str] = None
    docTitle: Optional[str] = None
    description: Optional[str] = None
    descriptionVector: Optional[List[float]] = None

@dataclass
class AzureAISearchDataSourceOptions:
    name: str
    indexName: str
    azureAISearchApiKey: str
    azureAISearchEndpoint: str

from azure.core.credentials import AzureKeyCredential
from azure.search.documents import SearchClient
import json

@dataclass
class Result:
    def __init__(self, output, length, too_long):
        self.output = output
        self.length = length
        self.too_long = too_long

class AzureAISearchDataSource(DataSource):
    def __init__(self, options: AzureAISearchDataSourceOptions):
        self.name = options.name
        self.options = options
        self.searchClient = SearchClient(
            options.azureAISearchEndpoint,
            options.indexName,
            AzureKeyCredential(options.azureAISearchApiKey)
        )
        
    def name(self):
        return self.name

    async def render_data(self, _context: TurnContext, memory: Memory, tokenizer: Tokenizer, maxTokens: int):
        query = memory.get('temp.input')

        user_id = _context.activity.from_property.id
        conv_id = _context.activity.conversation.id
        key = f"{conv_id}:{user_id}"

        stored = await storage.read([key])
        selection = None
        if stored and key in stored:
            selection = stored[key].get("selected_sources")  # "hr" or "procurement"

        # ✅ if nothing stored, allow both
        if selection is None:
            selection = "both"

        # ✅ now filter
        if self.name == "azure-ai-search-hr" and selection != "hr":
            return Result('', 0, False)
        if self.name == "azure-ai-search-procurement" and selection != "procurement":
            return Result('', 0, False)

        embedding = await get_embedding_vector(query)
        print(f"🔎 DataSource {self.name} called with query: {query}")
        vector_query = VectorizedQuery(vector=embedding, k_nearest_neighbors=2, fields="descriptionVector")

        if not query:
            return Result('', 0, False)

        selectedFields = [
            'docTitle',
            'description',
            'location',
            'descriptionVector',
        ]

        searchResults = self.searchClient.search(
            search_text=query,
            select=selectedFields,
            vector_queries=[vector_query],
            query_type="semantic",
            semantic_configuration_name="my-semantic-config"
        )

        if not searchResults:
            print(f"⚠️ No results found for query: {query}")
            return Result('', 0, False)
        print(f"✅ Found {(searchResults)} results for query: {query}") 
        usedTokens = 0
        doc = ''
        for result in searchResults:
            tokens = len(tokenizer.encode(json.dumps(result["description"])))

            if usedTokens + tokens > maxTokens:
                break

            doc += json.dumps(result["description"])
            usedTokens += tokens

        return Result(doc, usedTokens, usedTokens > maxTokens)