import os
<<<<<<< HEAD
import sys

# Ensure current directory is in python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from availability_app import app

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5001))
    print(f"Starting SAP Product Availability Predictor on http://127.0.0.1:{port}")
    app.run(host="0.0.0.0", port=port, debug=True)
=======
from stock_shortage_app.availability_app import app

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
>>>>>>> 32b3bcf3109b6a7029a23c39993996c77ca32e4b
