import os
import time
import subprocess
import sys

while True:
    print("Starting Streamlit...")
    p = subprocess.Popen([
        sys.executable, "start.py"
    ])
    p.wait()
    print(f"Streamlit exited with code {p.returncode}. Restarting in 2 seconds...")
    time.sleep(2)
