import os
import sys

# Ensure current directory is in path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from availability_app import app

if __name__ == '__main__':
    print("==================================================")
    print(" SAP Product Availability Predictor Application")
    print("==================================================")
    print("Running on: http://127.0.0.1:5001")
    print("Upload MB52 & COHV files to predict stockouts and delivery fulfillment.")
    print("==================================================")
    app.run(debug=True, port=5001, threaded=True)
