# PM Agent Backend

This is a Python-based backend service with API endpoints, designed to run on AWS Lambda.

## API Documentation

The service provides the following endpoints:

### Health Check
- **URL**: `/health`
- **Method**: `GET`
- **Response**:
  ```json
  {
    "status": "healthy",
    "message": "Service is running"
  }
  ```

## Project Structure

```
├── src/                    # Source code
│   ├── models/            # ML/DL model implementations
│   ├── data/              # Data processing and loading
│   ├── utils/             # Utility functions and helpers
│   └── main.py           # Main application entry point
├── tests/                 # Unit tests
├── requirements.txt       # Project dependencies
├── .env.example          # Example environment variables
└── README.md             # Project documentation
```

## Setup

1. Create a virtual environment:
```bash
python -m venv venv
```

2. Activate the virtual environment:
- Windows:
```bash
.\venv\Scripts\activate
```
- Unix/MacOS:
```bash
source venv/bin/activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Copy `.env.example` to `.env` and update with your configurations:
```bash
cp .env.example .env
```

## Usage

### Local Development
Run the FastAPI application locally:
```bash
cd src
uvicorn main:app --reload
```

The API will be available at:
- API Endpoint: http://localhost:8000
- Swagger Documentation: http://localhost:8000/docs
- ReDoc Documentation: http://localhost:8000/redoc

### AWS Lambda Deployment

1. Create a deployment package:
```bash
# Create a deployment directory
mkdir deployment
cd deployment

# Create a Python virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: .\venv\Scripts\activate

# Install dependencies
pip install -r ../requirements.txt

# Copy source files
cp -r ../src .

# Create deployment package
zip -r ../deployment.zip ./src
cd venv/lib/python3.10/site-packages
zip -r ../../../../deployment.zip .
cd ../../../../
```

2. Configure AWS Lambda:
   - Create a new Lambda function
   - Runtime: Python 3.10
   - Handler: `src.main.handler`
   - Upload the `deployment.zip` file
   - Configure environment variables if needed
   - Set up API Gateway as trigger
   - Configure memory and timeout settings as needed

3. After deployment, your API will be available at the API Gateway URL provided by AWS.

## Development

- Add your models in the `src/models/` directory
- Process and prepare data in the `src/data/` directory
- Add utility functions in `src/utils/`
- Write tests in the `tests/` directory

## License

[Add your license here]