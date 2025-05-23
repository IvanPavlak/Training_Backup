@echo off

:: Check hostname to determine which Python script to run
if "%COMPUTERNAME%" == "PC-WIN11" ( 
    set "python_script=E:\Development\GitHub\Training_Backup\TrainingBackup.py"
) else (
    set "python_script=C:\Users\Ivan\Development\GitHub\Training_Backup\TrainingBackup.py"
)

:: Activate the conda environment and run the Python script
call "C:\Users\Ivan\miniconda3\Scripts\activate.bat" "C:\Users\Ivan\miniconda3"
call conda activate TrainingBackup
python "%python_script%"