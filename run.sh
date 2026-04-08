#!/bin/bash
# Script to run xMLTree with virtual environment

set -e

# Check if virtual environment exists
if [ ! -d ".venv" ]; then
    echo "Virtual environment not found."
    echo "Please run ./setup_venv.sh first to set up the environment."
    exit 1
fi

# Activate virtual environment
source .venv/bin/activate

# Run the application
echo "Running xMLTree..."
python src/xMLTree.py