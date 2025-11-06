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
├── lambda_function.py    # AWS Lambda handler
├── src/                  # Source code
│   ├── models/          # ML/DL model implementations
│   ├── data/            # Data processing and loading
│   └── utils/           # Utility functions and helpers
├── tests/               # Unit tests
├── requirements.txt     # Project dependencies
├── .env.example        # Example environment variables
└── README.md           # Project documentation
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

### Automated Deployment Process

This project uses GitHub Actions for automated building and packaging. The pipeline automatically:
1. Runs tests
2. Creates a deployment package
3. Makes the package available as an artifact

To use the automated deployment:

1. Push your changes to the `main` branch or create a pull request
2. GitHub Actions will automatically run the build pipeline
3. Download the deployment package from the Actions tab in GitHub

### Manual Deployment

If you need to create the deployment package locally:

#### Windows PowerShell
```powershell
# Create deployment directory and virtual environment
mkdir deployment
cd deployment
python -m venv venv
.\venv\Scripts\activate

# Install dependencies
pip install -r ..\requirements.txt

# Copy source files and create zip
Copy-Item -Path ..\src -Destination . -Recurse
Compress-Archive -Path src\*, venv\Lib\site-packages\* -DestinationPath deployment.zip -Force
```

#### Unix/Linux/MacOS
```bash
# Create deployment directory and virtual environment
mkdir deployment
cd deployment
python -m venv venv
source venv/bin/activate

# Install dependencies
pip install -r ../requirements.txt

# Copy source files and create zip
cp -r ../src .
zip -r deployment.zip ./src
cd venv/lib/python3.10/site-packages
zip -r ../../../../deployment.zip .
```

### AWS Lambda Configuration

1. Create/Update Lambda Function:
   - Runtime: Python 3.10 or Python 3.12
   - Handler: `lambda_function.handler`
   - Memory: 256MB (recommended)
   - Timeout: 30 seconds
   - Upload the `deployment.zip` file

2. Configure API Gateway:
   - Create REST API Gateway
   - Create resources for each endpoint:
     - `/health` (GET)
   - Deploy API to a stage (e.g., 'prod', 'dev')
   - Note the API Gateway URL

3. Environment Variables (if needed):
   - Configure in Lambda console
   - Add variables defined in your .env file

4. Monitoring:
   - CloudWatch Logs automatically enabled
   - Monitor API Gateway metrics
   - Set up CloudWatch alarms if needed

## Development

- Add your models in the `src/models/` directory
- Process and prepare data in the `src/data/` directory
- Add utility functions in `src/utils/`
- Write tests in the `tests/` directory

## License

[Add your license here]