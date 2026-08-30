import os
import sys

# Ensure current directory is in python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from availability_app import app

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5001))
    print(f"Starting SAP Product Availability Predictor on http://127.0.0.1:{port}")
    app.run(host="0.0.0.0", port=port, debug=True)
