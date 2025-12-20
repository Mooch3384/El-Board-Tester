import subprocess
import sys
import os

def main():
    # make sure pyinstaller is there
    subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    # build it
    subprocess.run([
        "pyinstaller",
        "--name=BoardTesterPro",
        "--onefile",
        "--windowed",
        "--clean",
        "main.py"
    ])
    
    print("\nDone! Check the dist folder.")

if __name__ == "__main__":
    main()