@echo off

REM Clear the output_tickets folder
del /Q "output_tickets\*"

REM Run pdf splitter
python3 pdf_split.py

REM Run email sender
python3 sender_windows.py