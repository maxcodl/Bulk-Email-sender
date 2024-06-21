import subprocess
import sys

# List of required packages
required_packages = [
    "undetected-chromedriver",
    "selenium",
    "pandas",
    "pyautogui"
]

# Function to check if a package is installed
def check_package(package):
    try:
        __import__(package)
        return True
    except ImportError:
        return False

# Install missing packages
def install_package(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# Check and install required packages
for package in required_packages:
    package_name = package.split('==')[0]  # Get the package name without version
    if not check_package(package_name):
        print(f"{package} is not installed. Installing...")
        install_package(package)
    else:
        print(f"{package} is already installed.")

# Run the main script
subprocess.check_call([sys.executable, "send.py"])
