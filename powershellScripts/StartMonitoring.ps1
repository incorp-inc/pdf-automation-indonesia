# Set folder to watch + files to watch + subfolders yes/no
$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = "C:\FORM1770S\Unprocessed"
$watcher.Filter = "*.pdf"  # Filter for PDF files only
$watcher.IncludeSubdirectories = $true
$watcher.EnableRaisingEvents = $true  

# Define actions after an event is detected
$action = {
    $path = $Event.SourceEventArgs.FullPath
    $changeType = $Event.SourceEventArgs.ChangeType
    $logline = "$(Get-Date), $changeType, $path"
    Add-Content "C:\FORM1770S\fileChangeLog.log" -Value $logline
    
    # Check if the event is a folder creation
    if ($changeType -eq 'Created' -and (Test-Path -Path $path -PathType Leaf) -and $path -like "*.pdf") {
        # Execute the script
        Send-GoogleChatNotification -Message "Received a PDF File and Excel"
        Send-GoogleChatNotification -Message "File Added - $logline"
        Send-GoogleChatNotification -Message "Automation Started"
        Start-Process powershell -ArgumentList "-File C:\Users\PKDev02\file-api\powershellScripts\1770SFillingFinal.ps1"
    }
}

# Register events
Register-ObjectEvent $watcher "Created" -Action $action

function Send-GoogleChatNotification {
    param (
        [string]$Message
    )

    $apiEndpoint = "https://chat.googleapis.com/v1/spaces/AAAA_0Hn3NM/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=-c1ikFJWQXX_WkWyeFeC_NpLG1w9F03FOaKG9SOQjVk"

    $headers = @{
        "Content-Type" = "application/json"
    }

    $body = @{
        "text" = $Message
    } | ConvertTo-Json

    try {
        Invoke-RestMethod -Uri $apiEndpoint -Method Post -Headers $headers -Body $body
        Write-Host "Google Chat notification sent successfully."
    }
    catch {
        Write-Host "Error occurred while sending the Google Chat notification: $_"
    }
}

# Allow the script to run indefinitely (or replace with your own exit condition)
while ($true) {
    Start-Sleep 2
}


# # Create a queue to hold script paths
# $scriptQueue = New-Object System.Collections.Queue

# # Set folder to watch + files to watch + subfolders yes/no
# $watcher = New-Object System.IO.FileSystemWatcher
# $watcher.Path = "C:\FORM1770S\Unprocessed"
# $watcher.Filter = "*.pdf"  # Filter for PDF files only
# $watcher.IncludeSubdirectories = $true
# $watcher.EnableRaisingEvents = $true  

# # Define actions after an event is detected
# $action = {
#     $path = $Event.SourceEventArgs.FullPath
#     $changeType = $Event.SourceEventArgs.ChangeType
#     $logline = "$(Get-Date), $changeType, $path"
#     Add-Content "C:\FORM1770S\fileChangeLog.log" -Value $logline
    
#     # Check if the event is a folder creation
#     if ($changeType -eq 'Created' -and (Test-Path -Path $path -PathType Leaf) -and $path -like "*.pdf") {
#         # Enqueue the script path
#         $scriptPath = "C:\Users\PKDev02\file-api\powershellScripts\1770SFillingFinal.ps1"
#         $scriptQueue.Enqueue($scriptPath)

#         # Notify that the script has been added to the queue
#         Send-GoogleChatNotification -Message "Script added to the queue: $scriptPath"
#     }
# }

# # Register events
# Register-ObjectEvent $watcher "Created" -Action $action

# function Send-GoogleChatNotification {
#     param (
#         [string]$Message
#     )

#     $apiEndpoint = "https://chat.googleapis.com/v1/spaces/AAAA_0Hn3NM/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=-c1ikFJWQXX_WkWyeFeC_NpLG1w9F03FOaKG9SOQjVk"

#     $headers = @{
#         "Content-Type" = "application/json"
#     }

#     $body = @{
#         "text" = $Message
#     } | ConvertTo-Json

#     try {
#         Invoke-RestMethod -Uri $apiEndpoint -Method Post -Headers $headers -Body $body
#         Write-Host "Google Chat notification sent successfully."
#     }
#     catch {
#         Write-Host "Error occurred while sending the Google Chat notification: $_"
#     }
# }

# # Process the script queue
# while ($true) {
#     if ($scriptQueue.Count -gt 0) {
#         $scriptToExecute = $scriptQueue.Dequeue()
#         Write-Host "Executing script: $scriptToExecute"
#         Start-Process powershell -ArgumentList "-File $scriptToExecute" -Wait
#         Write-Host "Script execution completed."
#     }

#     Start-Sleep 2
# }

