import subprocess
import sys

def install_packages():
    # List of required packages
    required_packages = [
        'pandas',
        'openpyxl'
    ]
    
    print("Installing required packages...")
    
    for package in required_packages:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            print(f"Successfully installed {package}")
        except subprocess.CalledProcessError:
            print(f"Failed to install {package}")
            sys.exit(1)
    
    print("\nAll packages installed successfully!")
    print("\nNote: This project requires Python 3.7.1 for proper TLS encryption support.")
    print("If you're using Anaconda, you may need to run:")
    print("conda install python=3.7.1=h33f27b4_4")

if __name__ == "__main__":
    install_packages() 