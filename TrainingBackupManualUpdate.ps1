# Training Backup Manual Update
function Training-Backup {
    # Store the current directory
    $currentDirectory = Get-Location
    # Set default backup directory
    $backupDirectory = "E:\VSCode\GitHub\Training_Backup"
    
    # Check hostname and set backup directory accordingly
    $hostname = (Get-CimInstance Win32_ComputerSystem).Name
    if ($hostname -eq "Laptop-Win10") {
        $backupDirectory = "C:\Users\Ivan\VSCode\GitHub\Training_Backup"
    }
    
    # Change to the backup directory
    Set-Location -Path $backupDirectory
    
    try {
        # Execute the training-backup.bat file
        & ".\TrainingBackup.bat"
        # Output success message in green
        Write-Host "=> Training Backup Completed!" -ForegroundColor Green
    }
    catch {
        # Output error message in red
        Write-Host "=> Error during Training Backup!" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
    finally {
        # Return to the default directory
        Set-Location -Path $currentDirectory
    }
}