@echo off
rem Started by TQConnectorSetup.exe from the folder it unpacked into.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" -Payload "%~dp0payload.zip"
