#!/bin/bash
source venv/bin/activate
( sleep 2 && open http://localhost:5001 2>/dev/null || xdg-open http://localhost:5001 2>/dev/null ) &
python app.py
