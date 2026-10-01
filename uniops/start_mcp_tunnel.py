"""
UniOps Live MCP Tunnel & Local Server Runner
Provides an instant HTTPS URL for Claude Custom Connectors.
"""

import sys
import os
import subprocess
import time
from pathlib import Path

# Add project root to sys.path
BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

from xd_processor import create_app

def main():
    print("=" * 60)
    print("🚀 UNIOPS LIVE MCP SERVER FOR CLAUDE CUSTOM CONNECTOR")
    print("=" * 60)
    print("\nStarting local server on port 5000...")
    
    app = create_app()
    app.run(host="0.0.0.0", port=5000, debug=False)

if __name__ == "__main__":
    main()
