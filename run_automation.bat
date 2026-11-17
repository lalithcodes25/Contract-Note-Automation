@echo off
cd /d "D:\Lalith\contract_note_automation"
echo ================================= >> cron.log
echo Execution started at %DATE  %TIME% >> cron.log
python contract_note_automation.py %* >> cron.log 2>&1
echo Execution finished with exit code %ERRORLEVEL% >> cron.log