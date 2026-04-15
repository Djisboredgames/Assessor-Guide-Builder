#!/bin/bash
PROJECT_DIR="$HOME/Desktop/assessor-guide-builder"

if [ ! -d "$PROJECT_DIR" ]; then
    osascript -e 'display alert "Error" message "Project folder not found at ~/Desktop/assessor-guide-builder" as critical'
    exit 1
fi

cd "$PROJECT_DIR"

STREAMLIT="$PROJECT_DIR/venv/bin/streamlit"

if [ ! -f "$STREAMLIT" ]; then
    osascript -e 'display alert "Error" message "Streamlit not installed. Open Terminal and run: cd ~/Desktop/assessor-guide-builder && source venv/bin/activate && pip install -r requirements.txt" as critical'
    exit 1
fi

lsof -ti:8501 | xargs kill -9 2>/dev/null

(sleep 2 && open "http://localhost:8501") &

"$STREAMLIT" run app.py --server.headless true --browser.gatherUsageStats false
