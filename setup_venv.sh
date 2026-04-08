#!/bin/bash
# Script to set up a Python virtual environment for xMLTree

set -e

echo "Setting up Python virtual environment for xMLTree..."
echo ""

# Check if python3 is available
if ! command -v python3 &> /dev/null; then
    echo "Error: python3 not found. Please install Python 3.8 or higher."
    exit 1
fi

# Create virtual environment
if [ ! -d ".venv" ]; then
    echo "Creating virtual environment in .venv/"
    python3 -m venv .venv
    echo "Virtual environment created."
else
    echo "Virtual environment already exists at .venv/"
fi

# Activate virtual environment
echo "Activating virtual environment..."
source .venv/bin/activate

# Upgrade pip
echo "Upgrading pip..."
pip install --upgrade pip

# Install requirements
echo "Installing required packages..."
pip install -r requirements.txt

echo ""
echo "Setup complete! Virtual environment is ready."
echo ""
echo "To run xMLTree with the virtual environment:"
echo "  source .venv/bin/activate"
echo "  python src/xMLTree.py"
echo ""
echo "Or use the run.sh script:"
echo "  ./run.sh"
echo ""