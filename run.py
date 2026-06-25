# run.py
import sys
import os

# Add the current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app.main import app
import uvicorn

if __name__ == "__main__":
    # Check if running as .exe
    if getattr(sys, 'frozen', False):
        uvicorn.run(app, host="0.0.0.0", port=8000)
    else:
        uvicorn.run("app.main:app", host="0.0.0.0", port=8000, reload=True)