#!/bin/bash
cd "$(dirname "$0")"
streamlit run app.py --server.address 127.0.0.1 --server.port 8501
