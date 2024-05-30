# Training Backup Manual Update
function Training-Backup {
    # Store the current directory
    $currentDirectory = Get-Location
    # Set the directory path
    $backupDirectory = "E:\VSCode\GitHub\Training_Backup"
    # Change to the backup directory
    Set-Location -Path $backupDirectory
    try {
        # Execute the training-backup.bat file
        & ".\TrainingBackup_PC.bat"
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