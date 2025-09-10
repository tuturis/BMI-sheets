#!/bin/bash

# Business Model Innovation Sheet Generator
# This script sets up the environment and generates the BMI Excel sheet

echo "=== Business Model Innovation Sheet Generator ==="
echo

# Check if Python 3 is available
if ! command -v python3 &> /dev/null; then
    echo "Error: Python 3 is not installed or not in PATH"
    exit 1
fi

echo "✓ Python 3 found: $(python3 --version)"

# Install requirements if they don't exist
if [ ! -d "$HOME/.local/lib/python*/site-packages/openpyxl" ] && [ ! -d "venv" ]; then
    echo "📦 Installing dependencies..."
    python3 -m pip install -r requirements.txt
    echo "✓ Dependencies installed"
else
    echo "✓ Dependencies already installed"
fi

# Generate the Excel sheet
echo "🏗️  Generating BMI Excel sheet..."
python3 generate_bmi_sheet.py

# Check if the file was created successfully
if [ -f "business-model-innovation.xlsx" ]; then
    echo "✅ Success! Generated business-model-innovation.xlsx"
    echo "📊 File size: $(du -h business-model-innovation.xlsx | cut -f1)"
    echo "📅 Created: $(date)"
    echo
    echo "The Excel file contains:"
    echo "  • BMI Overview (Introduction and instructions)"
    echo "  • Business Model Canvas (9-block framework)"
    echo "  • Innovation Framework (Opportunity analysis)"
    echo "  • Analysis & Metrics (KPIs and risk assessment)"
    echo
    echo "Ready for use in business model innovation planning! 🚀"
else
    echo "❌ Error: Failed to generate Excel file"
    exit 1
fi