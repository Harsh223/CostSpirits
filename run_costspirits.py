#!/usr/bin/env python3
"""
Quick start script for CostSpirits
Run this file to launch the application
"""

import subprocess
import sys
import os

def main():
    """Launch CostSpirits application"""
    try:
        # Check if streamlit is installed
        subprocess.run([sys.executable, "-c", "import streamlit"], check=True, capture_output=True)
        
        # Get the directory where this script is located
        script_dir = os.path.dirname(os.path.abspath(__file__))
        costspirits_path = os.path.join(script_dir, "CostSpirits.py")
        
        # Launch the Streamlit app
        print("🚀 Launching CostSpirits...")
        print("📊 Opening spacecraft cost estimation tool...")
        print("🌐 The application will open in your default web browser")
        print("⏹️  Press Ctrl+C to stop the application")
        print("-" * 50)
        
        subprocess.run([sys.executable, "-m", "streamlit", "run", costspirits_path])
        
    except subprocess.CalledProcessError:
        print("❌ Error: Streamlit is not installed.")
        print("📦 Please install requirements first:")
        print("   pip install -r requirements.txt")
        sys.exit(1)
    except FileNotFoundError:
        print("❌ Error: CostSpirits.py not found.")
        print("📁 Make sure you're running this script from the CostSpirits directory.")
        sys.exit(1)
    except KeyboardInterrupt:
        print("\n👋 CostSpirits application stopped.")
        sys.exit(0)

if __name__ == "__main__":
    main()