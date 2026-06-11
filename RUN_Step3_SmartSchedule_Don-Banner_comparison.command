#!/bin/zsh

echo "========================================="
echo "Starting Smart Schedule Step 3..."
echo "========================================="
echo ""

# Navigate to the directory where this script is located
cd "$(dirname "$0")"

# Try to find Python in common Anaconda locations
PYTHON_PATH=""

if [ -f "$HOME/anaconda3/bin/python" ]; then
    PYTHON_PATH="$HOME/anaconda3/bin/python"
elif [ -f "/opt/anaconda3/bin/python" ]; then
    PYTHON_PATH="/opt/anaconda3/bin/python"
elif [ -f "$HOME/opt/anaconda3/bin/python" ]; then
    PYTHON_PATH="$HOME/opt/anaconda3/bin/python"
elif [ -f "/usr/local/anaconda3/bin/python" ]; then
    PYTHON_PATH="/usr/local/anaconda3/bin/python"
fi

if [ -z "$PYTHON_PATH" ]; then
    echo "ERROR: Could not find Anaconda Python installation"
    echo "Please check your Anaconda installation"
    echo "Press Enter to exit..."
    read
    exit 1
fi

echo "Using Python at: $PYTHON_PATH"
echo ""

# Check and install required packages
echo "Checking required packages..."
echo ""

"$PYTHON_PATH" -c "import openpyxl" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Installing openpyxl..."
    "$PYTHON_PATH" -m pip install openpyxl
fi

"$PYTHON_PATH" -c "import openpyxl" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Installing openpyxl..."
    "$PYTHON_PATH" -m pip install display
fi

"$PYTHON_PATH" -c "import xlrd" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Installing xlrd..."
    "$PYTHON_PATH" -m pip install xlrd
fi

echo ""
echo "All required packages are installed."
echo ""

# Run the Python script
"$PYTHON_PATH" source_Step3_SmartSchedule_v20260324_Don-Banner_comparison.py

# Capture exit status
EXIT_CODE=$?

echo ""
echo "========================================="
if [ $EXIT_CODE -eq 0 ]; then
    echo "Script completed successfully!"
else
    echo "Script failed with error code: $EXIT_CODE"
fi
echo "========================================="
echo ""
read -p "Press Enter to exit..."