#!/usr/bin/env python3
"""
Quick start script for AMCM Calculator
Provides dependency checking and helpful error messages for easy application launch.
"""

import sys
import subprocess
import os

def check_python_version():
    """Check if Python version is compatible"""
    if sys.version_info < (3, 8):
        print("❌ Error: Python 3.8 or higher is required")
        print(f"   Current version: {sys.version}")
        print("   Please upgrade Python and try again")
        return False
    print(f"✅ Python version: {sys.version.split()[0]}")
    return True

def check_streamlit():
    """Check if Streamlit is installed and working"""
    try:
        import streamlit
        print(f"✅ Streamlit version: {streamlit.__version__}")
        return True
    except ImportError:
        print("❌ Streamlit not found")
        return False

def check_dependencies():
    """Check if all required dependencies are installed"""
    required_packages = [
        'pandas',
        'numpy', 
        'openpyxl'
    ]
    
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package)
            print(f"✅ {package} is installed")
        except ImportError:
            missing_packages.append(package)
            print(f"❌ {package} is missing")
    
    return len(missing_packages) == 0, missing_packages

def install_dependencies():
    """Install missing dependencies"""
    print("\n🔧 Installing dependencies...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        print("✅ Dependencies installed successfully")
        return True
    except subprocess.CalledProcessError:
        print("❌ Failed to install dependencies")
        return False

def launch_app():
    """Launch the AMCM Calculator application"""
    print("\n🚀 Launching AMCM Calculator...")
    print("   Opening in your default web browser...")
    print("   Press Ctrl+C to stop the application")
    print("-" * 50)
    
    try:
        subprocess.run([sys.executable, "-m", "streamlit", "run", "amcm_calculator.py"])
    except KeyboardInterrupt:
        print("\n👋 Application stopped by user")
    except FileNotFoundError:
        print("❌ Streamlit command not found")
        print("   Try: python -m streamlit run amcm_calculator.py")
        return False
    except Exception as e:
        print(f"❌ Error launching application: {e}")
        return False
    
    return True

def main():
    """Main function to run all checks and launch the application"""
    print("🚀 AMCM Calculator - Quick Start")
    print("=" * 40)
    
    # Check Python version
    if not check_python_version():
        sys.exit(1)
    
    # Check if amcm_calculator.py exists
    if not os.path.exists("amcm_calculator.py"):
        print("❌ Error: amcm_calculator.py not found")
        print("   Make sure you're running this script from the correct directory")
        sys.exit(1)
    
    # Check Streamlit
    streamlit_ok = check_streamlit()
    
    # Check other dependencies
    deps_ok, missing = check_dependencies()
    
    # Install dependencies if needed
    if not streamlit_ok or not deps_ok:
        print(f"\n📦 Missing dependencies detected")
        if os.path.exists("requirements.txt"):
            install_choice = input("   Install missing dependencies? (y/n): ").lower().strip()
            if install_choice in ['y', 'yes']:
                if install_dependencies():
                    print("✅ All dependencies installed")
                else:
                    print("❌ Installation failed. Please install manually:")
                    print("   pip install -r requirements.txt")
                    sys.exit(1)
            else:
                print("❌ Cannot proceed without dependencies")
                sys.exit(1)
        else:
            print("❌ requirements.txt not found")
            print("   Please install dependencies manually:")
            print("   pip install streamlit pandas numpy openpyxl")
            sys.exit(1)
    
    # Launch the application
    print("\n✅ All checks passed!")
    launch_app()

if __name__ == "__main__":
    main()