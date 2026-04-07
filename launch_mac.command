#!/bin/bash
# XLSForm Quality Reviewer — macOS Launcher
# Double-click this file to start the app.

# Change to the directory where this script lives
cd "$(dirname "$0")"

echo "================================================"
echo "  XLSForm Quality Reviewer"
echo "================================================"
echo ""

# Check Python
if ! command -v python3 &>/dev/null; then
    echo "ERROR: Python 3 is not installed."
    echo "Please follow the instructions in SETUP.md (Step 1) and try again."
    read -p "Press Enter to close..."
    exit 1
fi

# Create virtual environment if it does not exist
if [ ! -d "venv" ]; then
    echo "Setting up the virtual environment for the first time..."
    python3 -m venv venv
    if [ $? -ne 0 ]; then
        echo "ERROR: Failed to create virtual environment."
        read -p "Press Enter to close..."
        exit 1
    fi
fi

# Activate the virtual environment
source venv/bin/activate

# Install / update dependencies
echo "Checking dependencies..."
pip install -q -r requirements.txt
if [ $? -ne 0 ]; then
    echo "ERROR: Failed to install dependencies."
    echo "Please check your internet connection and try again."
    read -p "Press Enter to close..."
    exit 1
fi

echo ""
echo "Starting the application..."
echo "The app will open in your browser at http://localhost:8501"
echo "To stop the app, close this window or press Ctrl + C."
echo ""

streamlit run app.py
