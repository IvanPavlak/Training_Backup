@echo off

:: Check hostname to determine which Python script to run
if "%COMPUTERNAME%" == "PC-WIN10" ( 
    set "python_script=E:\VSCode\GitHub\Training_Backup\TrainingBackup.py"
) else (
    set "python_script=C:\Users\Ivan\VSCode\GitHub\Training_Backup\TrainingBackup.py"
)

:: Activate the conda environment and run the Python script
call "C:\Users\Ivan\miniconda3\Scripts\activate.bat" "C:\Users\Ivan\miniconda3"
call conda activate TrainingBackup
python "%python_script%"