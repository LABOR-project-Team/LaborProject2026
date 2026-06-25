import subprocess
import sys
import importlib.util

# List of dependencies with version constraints
dependencies = [
    "pandas>=2.0.0",
    "openpyxl>=3.1.0",
    "xlrd>=2.0.1",
    "pillow>=9.0.0",
]

# Conditional dependency for Windows
if sys.platform == "win32":
    dependencies.append("pywin32>=306")

def install_dependencies():
    for package in dependencies:
        try:
            # Clean package name for checking (removing version specifiers)
            package_name = package.split(">")[0].split("=")[0]
            
            # Check if package is installed
            importlib.util.find_spec(package_name.replace("-", "_"))
            print(f"[OK] {package} is already installed.")
        except ImportError:
            print(f"[INSTALLING] {package}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            print(f"[SUCCESS] {package} installed.")

if __name__ == "__main__":
    install_dependencies()
    print("\nAll dependencies processed.")