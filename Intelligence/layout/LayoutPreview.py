import os 
from dotenv import load_dotenv
from fastapi import APIRouter, File, UploadFile
from fastapi.responses import JSONResponse
from motor.motor_asyncio import AsyncIOMotorClient
from gridfs import GridFS
from bson.objectid import ObjectId
from pymongo import MongoClient

load_dotenv()
# Initialize FastAPI app
router = APIRouter()

# MongoDB configuration
MONGO_URI = os.getenv("MONGO_URI")
DB_NAME = "slides_database"
COLLECTION_NAME = "slides_fs"

# Connect to MongoDB
client = MongoClient(MONGO_URI)
db = client[DB_NAME]
fs = GridFS(db)

@router.post("/upload-slide-preview")
async def upload_slide_preview(file: UploadFile = File(...)):
    try:
        # Read the file content
        file_content = await file.read()

        # Save the file in MongoDB using GridFS
        file_id = fs.put(file_content, filename=file.filename, content_type=file.content_type)

        # Return the file ID as a response
        return JSONResponse(content={"message": "File uploaded successfully!", "file_id": str(file_id)}, status_code=200)

    except Exception as e:
        return JSONResponse(content={"message": "Error uploading file", "error": str(e)}, status_code=500)

@router.get("/download-slide-preview/{file_id}")
async def download_slide_preview(file_id: str):
    try:
        # Retrieve the file from GridFS
        file_data = fs.get(ObjectId(file_id))
        file_content = file_data.read()

        # Return the file content as a response
        return JSONResponse(content={"filename": file_data.filename, "content": file_content.decode("latin1")}, status_code=200)

    except Exception as e:
        return JSONResponse(content={"message": "Error retrieving file", "error": str(e)}, status_code=500)