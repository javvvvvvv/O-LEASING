import sys
import traceback
import asyncio
import os

if sys.platform == 'win32':
    try:
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    except AttributeError:
        pass

import streamlit.web.cli as stcli

if __name__ == '__main__':
    sys.argv = [
        "streamlit", 
        "run", 
        "app.py", 
        "--server.port", "8501", 
        "--server.sslCertFile", "certs/cert.pem", 
        "--server.sslKeyFile", "certs/key.pem", 
        "--server.headless", "true"
    ]
    try:
        code = stcli.main()
        with open("crash.log", "w") as f:
            f.write(f"Exited cleanly with code {code}\n")
        sys.exit(code)
    except SystemExit as e:
        with open("crash.log", "w") as f:
            f.write(f"SystemExit: {e.code}\n")
        sys.exit(e.code)
    except Exception as e:
        with open("crash.log", "w") as f:
            f.write("Crash:\n")
            traceback.print_exc(file=f)
        sys.exit(1)
