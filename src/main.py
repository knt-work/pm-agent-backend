"""
Main module for the AI application.
This serves as the entry point for the application.
"""

from fastapi import FastAPI
from mangum import Mangum
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Initialize FastAPI app
app = FastAPI(
    title="PM Agent Backend",
    description="Backend API for PM Agent",
    version="1.0.0"
)

@app.get("/health")
async def health_check():
    """
    Simple health check endpoint to verify the API is running.
    """
    return {
        "status": "healthy",
        "message": "Service is running"
    }

# Create Lambda handler
handler = Mangum(app)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)