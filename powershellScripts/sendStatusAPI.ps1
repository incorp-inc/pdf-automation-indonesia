function SendStatus {
    Param (
        [string]$status,
        [string]$percentage,
        [string]$docid,
        [string]$message
    )

    # Set the URL of your API endpoint
    $apiUrlTest = "http://10.0.0.129:8722/api/form1770/status"

        # Set the URL of your API endpoint
    $apiUrlLive = "https://portal.incorp.asia/myportal/api/form1770/status"

    # Create a hashtable for headers
    $body = @{
        "status" = $status
        "percentage" = $percentage
        "docId" = $docid
        "message" = $message
    }

    Write-Host $status $percentage $docid $message

    # Convert the hashtable to JSON
    $jsonBody = $body | ConvertTo-Json

   try {
    $responseTest = Invoke-RestMethod -Uri $apiUrlTest -Method Post -Body $jsonBody -ContentType "application/json"
    $responseTest
    } catch {
    Write-Host "Error calling API: $_"
}
}

 $FileName = '638423923863635695'

 SendStatus -status "pending" -percentage "10" -docid $FileName -message "Completed - Retreving data from provided Excel"