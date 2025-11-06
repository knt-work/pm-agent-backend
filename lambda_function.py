"""
AWS Lambda handler wrapper for the PM Agent Backend.
This file serves as the entry point for AWS Lambda and imports the main application.
"""

import os
import sys

# Add the current directory to Python path to find the src module
sys.path.append(os.path.dirname(os.path.realpath(__file__)))

from mangum import Mangum
from src.main import app

# Create Lambda handler
handler = Mangum(app)