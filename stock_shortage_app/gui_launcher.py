import os
import sys
import threading
import time
import webview
from availability_app import app

def run_flask():
    app.run(host="127.0.0.1", port=5001, debug=False, threaded=True)

if __name__ == '__main__':
    # Start Flask server in background thread
    t = threading.Thread(target=run_flask, daemon=True)
    t.start()
    
    # Wait briefly for Flask to initialize
    time.sleep(1.0)

    # Open native desktop window GUI loading the Flask Web UI
    webview.create_window(
        'SAP Product Availability Predictor',
        'http://127.0.0.1:5001',
        width=1280,
        height=850,
        resizable=True
    )
    webview.start()
