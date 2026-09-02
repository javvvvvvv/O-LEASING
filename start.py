import sys
import asyncio

# Fix for asyncio ProactorEventLoop crashing on WinError 10054 (ConnectionResetError)
# when the user refreshes the page or browser drops the websocket connection.
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
    sys.exit(stcli.main())
