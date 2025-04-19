import os
import sys
import subprocess
import time

def check_requirements():
    """Check if all required files exist"""
    required_files = [
        'sender.py',
        'email-list.xlsx',
        'subject.txt',
        'message.txt'
    ]
    
    missing_files = []
    for file in required_files:
        if not os.path.exists(file):
            missing_files.append(file)
    
    return missing_files

def check_python_version():
    """Check if Python version is compatible"""
    if sys.version_info < (3, 7, 1) or sys.version_info >= (3, 7, 2):
        return False
    return True

def install_dependencies():
    """Install required packages"""
    print("\nInstalling required packages...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pandas", "openpyxl"])
        print("Packages installed successfully!")
        return True
    except subprocess.CalledProcessError:
        print("Failed to install packages. Please run 'python install.py' manually.")
        return False

def main():
    print("=== Email Sender Launcher ===")
    print("Checking requirements...")
    
    # Check Python version
    if not check_python_version():
        print("\n⚠️ Warning: This project requires Python 3.7.1 for proper TLS encryption support.")
        print("If you're using Anaconda, please run:")
        print("conda install python=3.7.1=h33f27b4_4")
        print("\nDo you want to continue anyway? (y/n): ", end='')
        if input().lower() not in ['y', 'yes']:
            sys.exit(1)
    
    # Check for required files
    missing_files = check_requirements()
    if missing_files:
        print("\n❌ Missing required files:")
        for file in missing_files:
            print(f"- {file}")
        print("\nPlease create these files before running the sender.")
        sys.exit(1)
    
    # Check for optional files
    optional_files = {
        'sender_name.txt': 'sender name',
        'document.pdf': 'PDF attachment'
    }
    
    for file, description in optional_files.items():
        if not os.path.exists(file):
            print(f"\nℹ️ Note: {description} file ({file}) not found. This is optional.")
    
    # Install dependencies if needed
    try:
        import pandas
        import openpyxl
    except ImportError:
        if not install_dependencies():
            sys.exit(1)
    
    # Run the sender script
    print("\n✅ All requirements met!")
    print("\nStarting email sender...")
    time.sleep(1)  # Give user time to read messages
    
    try:
        subprocess.check_call([sys.executable, "sender.py"])
    except subprocess.CalledProcessError:
        print("\n❌ Error running sender.py")
        sys.exit(1)

if __name__ == "__main__":
    main() 