


 function Stop-ProcessByName {
    param (
        [string]$ProcessName
    )

    # Run the taskkill command to terminate the process
    $output = & taskkill /F /IM $ProcessName 2>&1

    # Check the output for success or failure
    if ($output -like "*SUCCESS*") {
        Write-Host "The process '$ProcessName' has been terminated."
        return 0
    }
    else {
        Write-Host "Failed to terminate the process '$ProcessName'."
        Write-Host "Error message: $output"
        return 1
    }
}

Stop-ProcessByName -ProcessName "EXCEL.exe"

Stop-ProcessByName -ProcessName "Powershell.exe"