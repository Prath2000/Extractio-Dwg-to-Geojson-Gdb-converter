@echo off
:: ============================================================
:: launcher.bat — extractio launcher
:: Place in same folder as launcher.py and extractio.py
:: Add this folder to Windows PATH to use 'extractio' from anywhere
:: ============================================================
python "%~dp0launcher.py" %*
