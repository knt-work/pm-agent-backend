"""
Build script for creating AWS Lambda deployment package
"""
import os
import shutil
from pathlib import Path

def build_package():
    """Create deployment package for AWS Lambda."""
    print("Creating deployment package...")
    
    # Clean up old deployment files
    if os.path.exists('deployment'):
        shutil.rmtree('deployment')
    if os.path.exists('deployment.zip'):
        os.remove('deployment.zip')
        
    # Create fresh deployment directory
    os.makedirs('deployment')
    
    # Copy required files
    shutil.copy2('lambda_function.py', 'deployment/')
    shutil.copytree('src', 'deployment/src')
    
    # Create virtual environment
    os.system('python -m venv deployment/venv')
    
    # Install dependencies in a clean virtual environment
    if os.name == 'nt':  # Windows
        # Create and activate virtual environment
        os.system('deployment\\venv\\Scripts\\pip install --no-cache-dir -r requirements.txt')
        
        # Create a fresh zip file
        if os.path.exists('deployment.zip'):
            os.remove('deployment.zip')
        
        # First create the zip with the source files
        os.system('powershell Compress-Archive -Path lambda_function.py,src -DestinationPath deployment.zip -Force')
        
        # Change to the site-packages directory and add everything to the zip
        site_packages = 'deployment\\venv\\Lib\\site-packages'
        os.system(f'cd "{site_packages}" && powershell Compress-Archive -Path * -DestinationPath ..\\..\\..\\..\\deployment.zip -Update')
    else:  # Unix/Linux/MacOS
        os.system('deployment/venv/bin/pip install --no-cache-dir -r requirements.txt')
        if os.path.exists('deployment.zip'):
            os.remove('deployment.zip')
        os.system('zip -r deployment.zip lambda_function.py src/')
        os.system('cd deployment/venv/lib/python*/site-packages && zip -r ../../../../deployment.zip * -x "*/__pycache__/*" "*.dist-info/*"')
    
    # Clean up deployment directory
    shutil.rmtree('deployment')
    print("Deployment package created: deployment.zip")

if __name__ == "__main__":
    build_package()