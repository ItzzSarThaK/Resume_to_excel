@echo off
echo Starting Resume Scanner...
if exist ".venv\Scripts\python.exe" (
    echo Using virtual environment...
    .venv\Scripts\python.exe resume_scanner.py
) else (
    echo Using system Python...
    python resume_scanner.py
)
pause


