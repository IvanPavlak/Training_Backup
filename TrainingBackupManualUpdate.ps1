function Training-Backup {
	$currentDirectory = Get-Location
	$backupDirectory = "E:\Development\GitHub\Training_Backup"
    
	$hostname = (Get-CimInstance Win32_ComputerSystem).Name
	if ($hostname -eq "Laptop-Win11") {
		$backupDirectory = "C:\Users\Ivan\Development\GitHub\Training_Backup"
	}
    
	Set-Location -Path $backupDirectory
    
	try {
		& ".\TrainingBackup.bat"
		Write-Host "`n=> Training Backup Completed!`n" -ForegroundColor Green
	}
	catch {
		Write-Host "`n=> Error during Training Backup!`n" -ForegroundColor Red
		Write-Host $_.Exception.Message -ForegroundColor Red
	}
	finally {
		Set-Location -Path $currentDirectory
	}
}