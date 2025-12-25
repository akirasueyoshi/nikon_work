#!/bin/bash
# Setup script for Linux/Mac

set -e

echo "================================"
echo "Document Relevance Matrix Setup"
echo "================================"
echo ""

# Check if uv is installed
if ! command -v uv &> /dev/null; then
    echo "Installing uv..."
    curl -LsSf https://astral.sh/uv/install.sh | sh
    export PATH="$HOME/.cargo/bin:$PATH"
else
    echo "✓ uv is already installed"
fi

# Sync dependencies
echo ""
echo "Installing dependencies..."
uv sync

echo ""
echo "================================"
echo "✓ Setup completed!"
echo "================================"
echo ""
echo "Try running:"
echo "  uv run extract-links examples/test_files"
echo "  uv run build-matrix extraction_results/document_graph_*.json"
