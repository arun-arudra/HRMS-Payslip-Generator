import subprocess
import sys
import os
import logging
from pathlib import Path

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s: %(message)s")
logger = logging.getLogger("installer")

def install_requirements():
    """Install dependencies from requirements.txt."""
    req_file = Path("requirements.txt")
    if not req_file.exists():
        logger.error(f"'{req_file}' not found. Please create it with the required packages.")
        return False
    
    logger.info("Installing required Python packages...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", str(req_file)])
        logger.info("Installation complete. All dependencies are met.")
        return True
    except subprocess.CalledProcessError as e:
        logger.error(f"Failed to install packages: {e}")
        return False
    except Exception as e:
        logger.error(f"An unexpected error occurred during installation: {e}")
        return False

def main():
    """Main function to run the installer and then the main script."""
    # Check if requirements.txt exists
    if not Path("requirements.txt").exists():
        logger.info("Creating a default requirements.txt file...")
        with open("requirements.txt", "w") as f:
            f.write("pandas\n")
            f.write("openpyxl\n")
            f.write("reportlab\n")
            f.write("svglib\n")
            f.write("pydyf\n") # pydyf is a dependency of svglib
            f.write("requests\n") # This is often useful for general Python projects; can be removed if not needed.
        logger.info("requirements.txt created.")
    
    if not install_requirements():
        logger.error("Could not install dependencies. Please fix the issue and run again.")
        sys.exit(1)
        
    logger.info("Dependencies installed. Now starting the main application...")
    
    # Run the main script
    try:
        subprocess.check_call([sys.executable, "hrms.py"])
    except FileNotFoundError:
        logger.error("hrms.py not found. Please ensure it is in the same directory.")
        sys.exit(1)
    except subprocess.CalledProcessError as e:
        logger.error(f"An error occurred while running hrms.py: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
