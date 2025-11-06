"""
Main module for the simple API application.
"""

from fastapi import FastAPI

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

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)