function Training-Backup {
	$currentDirectory = Get-Location
    
	Set-Location -Path $MachineSpecificPaths.TrainingBackupDirectory
    
	try {
		& ".\TrainingBackup.bat"
		Write-Host -ForegroundColor Green "`n=> Training Backup Completed!"
	}
	catch {
		Write-Host -ForegroundColor Red "`n=> Error during Training Backup!"
		Write-Host -ForegroundColor Red $_.Exception.Message
	}
	finally {
		Set-Location -Path $currentDirectory
	}
}