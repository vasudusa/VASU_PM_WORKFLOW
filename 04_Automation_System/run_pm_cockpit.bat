@echo off
setlocal
cd /d "%~dp0"
python -m pip install -r requirements.txt
python publish_guides.py
python pm_cockpit.py --mode all
endlocal
