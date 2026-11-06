#!/bin/bash
echo "Installing dependencies..."
pip3 install -q streamlit openpyxl xlrd
echo "Starting Excel Trimmer..."
streamlit run app.py --server.maxUploadSize=400
