import os
from dotenv import load_dotenv
from fastapi import APIRouter
from fastapi import FastAPI
from pydantic import BaseModel
from langchain_ollama import OllamaEmbeddings
from pymongo import MongoClient
import numpy as np

load_dotenv()

router = APIRouter()


# --- Configuration ---
MONGODB_URI = os.getenv("MONGO_URI")  # Use your MongoDB connection string if needed
DATABASE_NAME = "AutoSlider"
COLLECTION_NAME = "LayoutRAG"


class Item(BaseModel):
    jsonString : str 
    layoutContentDesc : str
    description: str

class SearchQuery(BaseModel):
    query : str 

# --- MongoDB setup ---
client = MongoClient(MONGODB_URI)
db = client[DATABASE_NAME]
collection = db[COLLECTION_NAME]

embeddings = OllamaEmbeddings(model="llama3.2:3b")

def get_embeddings():
    return embeddings

@router.post("/layouts/")
async def create_item(item: Item):
    """
    Endpoint to receive an item, generate embeddings for its description, and store it in MongoDB.
    """
    # Generate embeddings for the description
    vector = embeddings.embed_query(item.description)

    # Create the document to insert into MongoDB
    document = {
        "jsonString": item.jsonString,
        "layoutContentDesc" : item.layoutContentDesc,
        "description": item.description,
        "embedding": vector  # Store the vector as a field in the document
    }

    # Insert the document into MongoDB
    result = collection.insert_one(document)

    return {"message": "Item added successfully", "item_id": str(result.inserted_id)}

@router.post("/search/")
async def search_items(query: SearchQuery ):
    """
    Endpoint to search for items based on a query string.
    """
    # Generate embeddings for the query
    query_vector = embeddings.embed_query(query.query)

    # Search for similar items in MongoDB (using cosine similarity)
    results = collection.find({})
    # Extract the relevant information from the results
    items = []
    base_similarity = -1
    for result in results:
        db_vector = result["embedding"]
        similarity = np.dot(query_vector, db_vector)

        if similarity > base_similarity: 
            items.append({
                "jsonString": result["jsonString"],
                "description": result["description"],
                "layoutContentDesc": result["layoutContentDesc"]
            })

    return {"results": items}