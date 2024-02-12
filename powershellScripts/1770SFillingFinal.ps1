
# Import the module
Import-Module ImportExcel

#Click Function
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class Clicker
{
    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int X, int Y);

    [DllImport("user32.dll")]
    public static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

    public static void Click(int x, int y)
    {
        SetCursorPos(x, y);
        mouse_event(0x0002, 0, 0, 0, 0); // Left mouse button down
        mouse_event(0x0004, 0, 0, 0, 0); // Left mouse button up
    }
}
"@

Add-Type -AssemblyName System.Windows.Forms

#Log file path
$Logfile = "C:\FORM1770S\PDF_logs.log"

# Set the path to the main folder
$mainFolderPath = "C:\FORM1770S\Unprocessed"

# Set the path to the backup folder
$BackupFolderPath = "C:\FORM1770S\Backup"
function log {
    Param ([string]$logstring)
    Add-content $Logfile -value $logstring
    Write-Host $logstring
}

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
function nextPage {     
    
    [Clicker]::Click(1308, 115)
    Start-Sleep .5
    [Clicker]::Click(1308, 115)
    Start-Sleep 1
    [Clicker]::Click(991, 108)
    Start-Sleep 1
}
function previousPage {

    [Clicker]::Click(1308, 115)
    Start-Sleep .5
    [Clicker]::Click(1308, 115)
    Start-Sleep 1
    [Clicker]::Click(308, 103)
    Start-Sleep 1
}

function mapAssetCode {
    Param (
        [string]$code
    )

    $validCodes = @(
        "011", "012", "013", "014", "019", "021", "022", "029",
        "031", "032", "033", "034", "035", "036", "037", "038", "039",
        "041", "042", "043", "049", "051", "052", "053", "054", "055",
        "059", "061", "062", "063", "069", "071", "072", "073", "079"
    )

    $loopCount = [array]::IndexOf($validCodes, $code)

    if ($loopCount -gt 0) {
        for ($i = 0; $i -le $loopCount; $i++) {
            # Simulate pressing the UP arrow key
            [System.Windows.Forms.SendKeys]::SendWait('{DOWN}')

            # Sleep for a short interval between key presses (adjust as needed)
            Start-Sleep -Milliseconds 10
        }
    }
    else {
        [System.Windows.Forms.SendKeys]::SendWait('{DOWN}')
        Write-Host "Invalid code. Please provide a valid code from the list."
    }
}

function mapLiabilityCode {
    Param (
        [string]$code
    )

    $validCodes = @(
        "101", "102", "103", "109"
    )

    $loopCount = [array]::IndexOf($validCodes, $code)

    if ($loopCount -ge 0) {
        for ($i = 1; $i -le $loopCount; $i++) {
            # Simulate pressing the UP arrow key
            [System.Windows.Forms.SendKeys]::SendWait('{DOWN}')

            # Sleep for a short interval between key presses (adjust as needed)
            Start-Sleep -Milliseconds 10
        }
    }
    else {
        Write-Host "Invalid code. Please provide a valid code from the list."
    }
}

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

function openPDF {
    param (
        [string]$FilePath,
        [string]$FolderName,
        [int]$TimeoutSeconds = 5
    )

    $sourcePdf = Get-ChildItem $FilePath
    Start-Process -FilePath $sourcePdf.FullName

    $wshell = New-Object -ComObject WScript.Shell
    $timeoutReached = $false
    $startTime = Get-Date

    while (-not $wshell.AppActivate("$($sourcePdf.Name) (SECURED)")) {
        Start-Sleep -Milliseconds 1000
        $elapsedTime = (Get-Date) - $startTime

        if ($elapsedTime.TotalSeconds -ge $TimeoutSeconds) {
            $FileName = $sourcePdf.Name
            $updatedFileName = $FileName.Replace(".pdf", "")
            Write-Host "PDF Viewer window not displayed within $TimeoutSeconds seconds. Deleting folder: $FolderName"
            SendStatus -status "error" -percentage "0" -emailSent "false" -docid $updatedFileName  -message "Error - Invalid PDF File. Kindly upload the Form-1770S PDF File"
            Send-GoogleChatNotification -Message "Invalid File or Error in Automation"
            Stop-ProcessByName -ProcessName Acrobat.exe
            Stop-ProcessByName -ProcessName Excel.exe
            Stop-ProcessByName -ProcessName EXCEL.exe
            Start-Sleep -Milliseconds 1000
            Remove-Item -Path $FolderName -Recurse -Force
            $timeoutReached = $true
            Exit
            break
           
        }

        Write-Host "PDF Viewer window not displayed yet, waiting..."
    }

    if (-not $timeoutReached) {
        Write-Host "PDF Viewer window displayed within $TimeoutSeconds seconds."
        Send-GoogleChatNotification -Message "PDF Filling Started"
    }
}

function sendPDF {
    Param ([string]$FileName)

    # Set the parameters
    $docId = $FileName
    Write-Host $docId


    # Set the URL of your API endpoint

    $apiUrl = "http://localhost:3000/sendPdf"

    $filePath = "C:\FORM1770S\FilledPDF\$FileName.pdf"

    $url1 = "https://portal.incorp.asia/myportal/api/form1770/documents"

    $url2 = "http://10.0.0.129:8722/api/form1770/documents"

    # Create a hashtable for headers
    $headers = @{
        "docid"    = $docId
        "filepath" = $filePath
        "url1"     = $url1
        "url2"     = $url2
    }

    # Make the API request
    $response = Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $headers

    # Print the API response
    $response

}
 

function SendStatus {
    Param (
        [string]$status,
        [string]$percentage,
        [string]$docid,
        [string]$message,
        [string]$emailSent
    )

    # Set the URL of your API endpoint
    $apiUrlTest = "http://10.0.0.129:8722/api/form1770/status"

    # Set the URL of your API endpoint
    $apiUrlLive = "https://portal.incorp.asia/myportal/api/form1770/status"

    # Create a hashtable for headers
    $body = @{
        "status"     = $status
        "percentage" = $percentage
        "docId"      = $docid
        "message"    = $message
        "emailSent"  = $emailSent
    }

    # Convert the hashtable to JSON
    $jsonBody = $body | ConvertTo-Json

    # Make the API request
    $responseTest = Invoke-RestMethod -Uri $apiUrlTest -Method Post -Body $jsonBody -ContentType "application/json"

    # Print the API response
    $responseTest

    # Make the API request
    $responseLive = Invoke-RestMethod -Uri $apiUrlLive -Method Post -Body $jsonBody -ContentType "application/json"

    # Print the API response
    $responseLive
}

function savePDF {
    Param ([string]$FileName)
   

    [Clicker]::Click(1132, 185)
    Start-Sleep 1
    $wshell = New-Object -ComObject WScript.Shell
    $wshell.SendKeys("^{s}")
    Start-Sleep 1
    $wshell = New-Object -ComObject WScript.Shell
    Start-Sleep 1
    $wshell.SendKeys("C:\FORM1770S\FilledPDF\$($FileName)")
    $wshell.SendKeys("~")
    Start-Sleep .5
    $wshell.SendKeys("{LEFT}")
    $wshell.SendKeys("~")
    Start-Sleep 3
    $wshell.SendKeys("%{F4}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys("~")

    #Stop-ProcessByName -ProcessName "Acrobat.exe"
}

function ProcessAssets {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject[]]$ObjectArray
    )

    $loopCounter = 0

    # Loop through each object in the array
    foreach ($object in $ObjectArray) {

        $loopCounter++


        # Access properties of the current object
        $code = $object.Code
        $description = $object.Description
        $amount = $object.Amount
        $type = $object.Type

        # Perform your processing logic here
        # Example: Print the details of each object
        Write-Host "Code: $code, Description: $description, Amount: $amount, Type: $type"

        if ($loopCounter -eq 1) {
            [Clicker]::Click(360, 135)
        }
        if ($loopCounter -eq 2) {
            [Clicker]::Click(360, 155)
        }
        if ($loopCounter -eq 3) {
            [Clicker]::Click(360, 175)
        }
        if ($loopCounter -eq 4) {
            [Clicker]::Click(360, 195)
        }
        if ($loopCounter -eq 5) {
            [Clicker]::Click(360, 215)
        }
        if ($loopCounter -eq 6) {
            [Clicker]::Click(360, 135)
        }
        if ($loopCounter -eq 7) {
            [Clicker]::Click(360, 155)
        }
        if ($loopCounter -eq 8) {
            [Clicker]::Click(360, 175)
        }
        if ($loopCounter -eq 9) {
            [Clicker]::Click(360, 195)
        }
        if ($loopCounter -eq 10) {
            [Clicker]::Click(360, 215)
        }
        if ($loopCounter -eq 11) {
            [Clicker]::Click(360, 135)
        }
        if ($loopCounter -eq 12) {
            [Clicker]::Click(360, 155)
        }
        if ($loopCounter -eq 13) {
            [Clicker]::Click(360, 175)
        }
        if ($loopCounter -eq 14) {
            [Clicker]::Click(360, 195)
        }
        if ($loopCounter -eq 15) {
            [Clicker]::Click(360, 215)
        }

        Start-Sleep -Milliseconds 300
        
        # Add your additional processing logic here
        # Simulate pressing and releasing the UP arrow key in a loop
        
        if ($loopCounter -eq 1) {
            for ($i = 1; $i -le 35; $i++) {
                # Simulate pressing the UP arrow key
                [System.Windows.Forms.SendKeys]::SendWait('{UP}')
    
                # Sleep for a short interval between key presses (adjust as needed)
                Start-Sleep -Milliseconds 20
            }
        }
        Start-Sleep .5
        $wshell = New-Object -ComObject WScript.Shell
        mapAssetCode -code $code
        Start-Sleep .5
        $wshell.SendKeys("{TAB}")
        Start-Sleep .5
        $wshell.SendKeys($description)
        Start-Sleep .5
        $wshell.SendKeys("{TAB}")
        Start-Sleep .5
        $wshell.SendKeys("2022")
        Start-Sleep .5
        $wshell.SendKeys("{TAB}")
        Start-Sleep .5
        $wshell.SendKeys($amount)
        Start-Sleep .5
        $wshell.SendKeys("{TAB}")
        Start-Sleep .5
        $wshell.SendKeys($description)
        Start-Sleep .5
        if ($object -ne $ObjectArray[-1]) {
            Start-Sleep .5
            $wshell.SendKeys("{TAB}")
            Start-Sleep .5
            $wshell.SendKeys("{TAB}")
            Start-Sleep .5
            $wshell.SendKeys("~")
            Start-Sleep 1
        }

    }
}

function ProcessLiability {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject[]]$ObjectArray
    )

    $loopCounter = 0

    $wshell = New-Object -ComObject WScript.Shell


    # Loop through each object in the array
    foreach ($object in $ObjectArray) {

        $loopCounter++

        # Access properties of the current object
        $code = $object.Code
        $description = $object.Description
        $amount = $object.Amount
        $type = $object.Type

        # Perform your processing logic here
        # Example: Print the details of each object
        Write-Host "Code: $code, Description: $description, Amount: $amount, Type: $type"

       
        
        [Clicker]::Click(300, 472)
        Start-Sleep 1

        if ($loopCounter -eq 1) {
            [Clicker]::Click(890, 386)
        }
        if ($loopCounter -eq 2) {
            [Clicker]::Click(885, 406)
        }
        if ($loopCounter -eq 3) {
            [Clicker]::Click(885, 426)
        }
        if ($loopCounter -eq 4) {
            [Clicker]::Click(885, 386)
        }
        if ($loopCounter -eq 5) {
            [Clicker]::Click(885, 406)
        }
        if ($loopCounter -eq 6) {
            [Clicker]::Click(885, 426)
        }
        if ($loopCounter -eq 7) {
            [Clicker]::Click(885, 386)
        }
        if ($loopCounter -eq 8) {
            [Clicker]::Click(885, 406)
        }
        if ($loopCounter -eq 9) {
            [Clicker]::Click(885, 426)
        }

        Start-Sleep .5
        $wshell.SendKeys("2022")
        Start-Sleep .5
        $wshell.SendKeys("{TAB}")
        Start-Sleep .5
        $wshell.SendKeys($amount)
        Start-Sleep .5
        $wshell.SendKeys("{TAB}")
        Start-Sleep .5
        $wshell.SendKeys("{DOWN}")
        Start-Sleep .5
        mapLiabilityCode -code $code
        Start-Sleep .5
        $wshell.SendKeys("{TAB}")
        Start-Sleep .5
        $wshell.SendKeys($description)
        Start-Sleep .5
        $wshell.SendKeys("{TAB}")
        Start-Sleep .5
        $wshell.SendKeys($description)
        Start-Sleep .5

        if ($object -ne $ObjectArray[-1]) {
            Start-Sleep 1
            $wshell.SendKeys("{TAB}")
            Start-Sleep 1
            $wshell.SendKeys("{TAB}")
        }

    }

    [Clicker]::Click(1097, 494)

    $wshell = New-Object -ComObject WScript.Shell
    Start-Sleep .5 
    $wshell.SendKeys("{TAB}") 


}

function ProcessDependents {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject[]]$ObjectArray
    )

    $loopCounter = 0


    # Loop through each object in the array
    foreach ($object in $ObjectArray) {

        $loopCounter++

        # Access properties of the current object
        $name = $object.Name
        $id = $object.ID
        $relation = $object.Relation
        $occupation = $object.Occupation

        # Perform your processing logic here
        # Example: Print the details of each object
        Write-Host "name: $name, id: $id, relation: $relation, occupation: $occupation"

        $wshell = New-Object -ComObject WScript.Shell
        
        $wshell.SendKeys("~")
        Start-Sleep 1

        
        if ($loopCounter -eq 1) {
            [Clicker]::Click(360, 575)
        }
        if ($loopCounter -eq 2) {
            [Clicker]::Click(360, 595)
        }
        if ($loopCounter -eq 3) {
            [Clicker]::Click(360, 615)
        }
        if ($loopCounter -eq 4) {
            [Clicker]::Click(360, 575)
        }
        if ($loopCounter -eq 5) {
            [Clicker]::Click(360, 595)
        }
        if ($loopCounter -eq 6) {
            [Clicker]::Click(360, 615)
        }
        if ($loopCounter -eq 7) {
            [Clicker]::Click(360, 575)
        }
        if ($loopCounter -eq 8) {
            [Clicker]::Click(360, 595)
        }
        if ($loopCounter -eq 9) {
            [Clicker]::Click(360, 615)
        }



        Start-Sleep 1
        $wshell.SendKeys($name)
        Start-Sleep .5
        $wshell.SendKeys("{TAB}")
        Start-Sleep .5
        #$wshell.SendKeys($id)
        $wshell.SendKeys("1234123412341234")
        Start-Sleep .5
        $wshell.SendKeys("{TAB}")
        Start-Sleep .5
        $wshell.SendKeys($relation)
        Start-Sleep .5
        $wshell.SendKeys("{TAB}")
        Start-Sleep .5
        $wshell.SendKeys($occupation)

        if ($object -ne $ObjectArray[-1]) {
            Start-Sleep .5
            $wshell.SendKeys("{TAB}")
        }

    }
}

function Get-AssetArray {
    param (
        [System.__ComObject]$ExcelSheet
    )

    $startRow = 7
    $codeColumn = 3
    $descriptionColumn = 4
    $amountColumn = 16
    $typeColumn = 5

    $assetArray = @()

    for ($row = $startRow; $true; $row++) {
        $code = $ExcelSheet.Cells.Item($row, $codeColumn).Text -replace '[^\d]'
        if ($code -eq '') {
            break
        }
        $description = $ExcelSheet.Cells.Item($row, $descriptionColumn).Text
        $amount = $ExcelSheet.Cells.Item($row, $amountColumn).Text -replace '[^\d]'
        $type = $ExcelSheet.Cells.Item($row, $typeColumn).Text

        $rowObject = [PSCustomObject]@{
            Code        = $code
            Description = $description
            Amount      = $amount
            Type        = $type
        }

        if ($type -eq 'Asset') {
            $assetArray += $rowObject
        } 
    }

    # Display the list of Asset objects
    log "List of Asset Rows:"
    $assetArray | ForEach-Object { Log $_ }

    return $assetArray
}

function Get-LiabilityArray {
    param (
        [System.__ComObject]$ExcelSheet
    )

    $startRow = 7
    $codeColumn = 3
    $descriptionColumn = 4
    $amountColumn = 16
    $typeColumn = 5

    $liabilityArray = @()

    for ($row = $startRow; $true; $row++) {
        $code = $ExcelSheet.Cells.Item($row, $codeColumn).Text -replace '[^\d]'
        if ($code -eq '') {
            break
        }
        $description = $ExcelSheet.Cells.Item($row, $descriptionColumn).Text
        $amount = $ExcelSheet.Cells.Item($row, $amountColumn).Text -replace '[^\d]'
        $type = $ExcelSheet.Cells.Item($row, $typeColumn).Text

        $rowObject = [PSCustomObject]@{
            Code        = $code
            Description = $description
            Amount      = $amount
            Type        = $type
        }

        if ($type -eq 'Liability') {
            $liabilityArray += $rowObject
        }
    }

    # Display the list of Liability objects
    log "List of Liabilty Rows:"
    $liabilityArray | ForEach-Object { Log $_ }
    return $liabilityArray
}

function Get-DependentArray {
    param (
        [System.__ComObject]$ExcelSheet
    )

    # Starting cell coordinates
    $startRow1 = 25
    $nameColumn = 4
    $idColumn = 5
    $relationColumn = 6
    $occupationColumn = 7


    # Arrays to store objects for each row type
    $dependentArray = @()

    for ($row = $startRow1; $true; $row++) {
        $name = $ExcelSheet.Cells.Item($row, $nameColumn).Text
    
        # Check if the code is empty (assuming the code is the key column)
        if ($name -eq '') {
            break
        }

        $id = $ExcelSheet.Cells.Item($row, $idColumn).Text
        $relation = $ExcelSheet.Cells.Item($row, $relationColumn).Text 
        $occupation = $ExcelSheet.Cells.Item($row, $occupationColumn).Text

        # Create an object for the current row
        $rowObject = [PSCustomObject]@{
            Name       = $name
            ID         = $id
            Relation   = $relation
            Occupation = $occupation
        }

        # Push the object to the appropriate array based on type
   
        $dependentArray += $rowObject
    }

    # Display the list of Dependent objects
    log "List of Dependent Rows:"
    $dependentArray | ForEach-Object { Log $_ }

    return $dependentArray
}

function ClearAssetLiability {

    $wshell = New-Object -ComObject WScript.Shell
    
    [Clicker]::Click(422, 257)

    for ($i = 1; $i -le 15; $i++) {
        # Simulate pressing the UP arrow key
        $wshell.SendKeys("~")
   
        # Sleep for a short interval between key presses (adjust as needed)
        Start-Sleep -Milliseconds 50
    }

    [Clicker]::Click(422, 473)

    for ($i = 1; $i -le 15; $i++) {
        # Simulate pressing the UP arrow key
        $wshell.SendKeys("~")
   
        # Sleep for a short interval between key presses (adjust as needed)
        Start-Sleep -Milliseconds 50
    }

    [Clicker]::Click(422, 636)

    for ($i = 1; $i -le 15; $i++) {
        # Simulate pressing the UP arrow key
        $wshell.SendKeys("~")
   
        # Sleep for a short interval between key presses (adjust as needed)
        Start-Sleep -Milliseconds 50
    }

    [Clicker]::Click(332, 138)
    Start-Sleep .5
}

function Fill1770S {
    param (
        [System.__ComObject]$excelSheet
    )

    $InterestTax = $excelSheet.Cells.Item(7, 6).Text -replace '[^\d]'

    $wshell = New-Object -ComObject WScript.Shell

    Start-Sleep 2

    [Clicker]::Click(1080, 335)
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("2233445566") #Telephone
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("12345678")  #Fax
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")     #Status
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("10")     #Employment Income
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("10")     #Domestic Income
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("20")     #Foreign Income
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("30")     #Donation
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("200")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("40")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("50")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("5555000")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
   
    Start-Sleep 1

    SendStatus -status "pending" -percentage "30" -emailSent "false" -docid $FileName -message "Completed - Form 1770 S-I - Section A"

    [Clicker]::Click(1307, 691)
    Start-Sleep .5
   
    [Clicker]::Click(1307, 691)

    Start-Sleep 1
   



}

function Fill1770S1 {
    param (
        [System.__ComObject]$excelSheet
    )

    $intrestAtt1 = $excelSheet.Cells.Item(7, 5).Text -replace '[^\d]'
    $Royalties = $excelSheet.Cells.Item(8, 5).Text -replace '[^\d]'
    $Rental = $excelSheet.Cells.Item(9, 5).Text -replace '[^\d]'
    $Award = $excelSheet.Cells.Item(10, 5).Text -replace '[^\d]'
    $Gains = $excelSheet.Cells.Item(11, 5).Text -replace '[^\d]'
    $OtherIncome = $excelSheet.Cells.Item(12, 5).Text -replace '[^\d]'

    $DonationGifts = $excelSheet.Cells.Item(16, 5).Text -replace '[^\d]'
    $Inheritance = $excelSheet.Cells.Item(17, 5).Text -replace '[^\d]'
    $Profit = $excelSheet.Cells.Item(18, 5).Text -replace '[^\d]'
    $ClaimsForHealth = $excelSheet.Cells.Item(19, 5).Text -replace '[^\d]'
    $DomesticScholarship = $excelSheet.Cells.Item(20, 5).Text -replace '[^\d]'
    $6A = $excelSheet.Cells.Item(22, 5).Text -replace '[^\d]'
    $6B = $excelSheet.Cells.Item(23, 5).Text -replace '[^\d]'
    $6C = $excelSheet.Cells.Item(24, 5).Text -replace '[^\d]'


    $wshell = New-Object -ComObject WScript.Shell

    Start-Sleep 2

    [Clicker]::Click(1077, 409)

    Start-Sleep .5
    $wshell.SendKeys("$intrestAtt1")
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("$Royalties")
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("$Rental")
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("$Award")
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("$Gains")
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("$OtherIncome")
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("$DonationGifts")
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("$Inheritance")
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("$Profit")
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("$ClaimsForHealth")
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("$DomesticScholarship")
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("$6A")
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("$6B")
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("$6C")
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("1")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("1")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("123451234512345")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("1")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{DOWN}")
    Start-Sleep .5
    $wshell.SendKeys("~")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("{DOWN}")
    Start-Sleep .5
    $wshell.SendKeys("{TAB}")
    Start-Sleep .5
    $wshell.SendKeys("1")

    Start-Sleep 1

    SendStatus -status "pending" -percentage "20" -emailSent "false" -docid $FileName -message "Completed - Form 1770 S-II"





}

function Fill1770S2 {
    param (
        [System.__ComObject]$excelSheet
    )

    $InterestTax = $excelSheet.Cells.Item(7, 6).Text -replace '[^\d]'
    $InterestDiscountTax = $excelSheet.Cells.Item(8, 6).Text -replace '[^\d]'
    $SaleOfShareTax = $excelSheet.Cells.Item(9, 6).Text -replace '[^\d]'
    $LotteryTax = $excelSheet.Cells.Item(10, 6).Text -replace '[^\d]'
    $OneTimeTax = $excelSheet.Cells.Item(11, 6).Text -replace '[^\d]'
    $HonoriumTax = $excelSheet.Cells.Item(12, 6).Text -replace '[^\d]'
    $ValueOfIncomeTax = $excelSheet.Cells.Item(13, 6).Text -replace '[^\d]'
    $RentalIncomeTax = $excelSheet.Cells.Item(14, 6).Text -replace '[^\d]'
    $ValueofBuildingTax = $excelSheet.Cells.Item(15, 6).Text -replace '[^\d]'
    $InterestPaidTax = $excelSheet.Cells.Item(16, 6).Text -replace '[^\d]'
    $IncomeFromTax = $excelSheet.Cells.Item(17, 6).Text -replace '[^\d]'
    $DividenTax = $excelSheet.Cells.Item(18, 6).Text -replace '[^\d]'
    $WifeIncomeTax = $excelSheet.Cells.Item(19, 6).Text -replace '[^\d]'
    $IncomeSubjectTax = $excelSheet.Cells.Item(20, 6).Text -replace '[^\d]'

    $Interest = $excelSheet.Cells.Item(7, 5).Text -replace '[^\d]'
    $InterestDiscount = $excelSheet.Cells.Item(8, 5).Text -replace '[^\d]'
    $SaleOfShare = $excelSheet.Cells.Item(9, 5).Text -replace '[^\d]'
    $Lottery = $excelSheet.Cells.Item(10, 5).Text -replace '[^\d]'
    $OneTime = $excelSheet.Cells.Item(11, 5).Text -replace '[^\d]'
    $Honorium = $excelSheet.Cells.Item(12, 5).Text -replace '[^\d]'
    $ValueOfIncome = $excelSheet.Cells.Item(13, 5).Text -replace '[^\d]'
    $RentalIncome = $excelSheet.Cells.Item(14, 5).Text -replace '[^\d]'
    $ValueofBuilding = $excelSheet.Cells.Item(15, 5).Text -replace '[^\d]'
    $InterestPaid = $excelSheet.Cells.Item(16, 5).Text -replace '[^\d]'
    $IncomeFrom = $excelSheet.Cells.Item(17, 5).Text -replace '[^\d]'
    $Dividen = $excelSheet.Cells.Item(18, 5).Text -replace '[^\d]'
    $WifeIncome = $excelSheet.Cells.Item(19, 5).Text -replace '[^\d]'
    $IncomeSubject = $excelSheet.Cells.Item(20, 5).Text -replace '[^\d]'

    $wshell = New-Object -ComObject WScript.Shell

    Start-Sleep 2


    [Clicker]::Click(688, 266)
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys($Interest)
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys($InterestTax)
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys($InterestDiscount)
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys($InterestDiscountTax)
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys($SaleOfShare)
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys($SaleOfShareTax)
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys($Lottery)
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys($LotteryTax)
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys($OneTime)
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys($OneTimeTax)
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys($Honorium)
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys($HonoriumTax)
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys($ValueOfIncome)
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys($ValueOfIncomeTax)
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys($RentalIncome)
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys($RentalIncomeTax)
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys($ValueofBuilding)
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys($ValueofBuildingTax)
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys($InterestPaid)
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys($InterestPaidTax)
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys($Dividen)
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys($DividenTax)
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys($WifeIncome)
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys($WifeIncomeTax)
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys($IncomeSubject)
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys($IncomeSubjectTax)
   
    Start-Sleep 1

    SendStatus -status "pending" -percentage "30" -emailSent "false" -docid $FileName -message "Completed - Form 1770 S-I - Section A"

    [Clicker]::Click(1307, 691)
    Start-Sleep .5
   
    [Clicker]::Click(1307, 691)

    Start-Sleep 1
   



}

function fillPDF {
    Param (
        [System.__ComObject]$ExcelWorkbook,
        [string]$FileName
    )

    SendStatus -status "pending" -percentage "10" -emailSent "false" -docid $FileName -message "Completed - Retreving data from provided Excel"

    #Parse Excel Sheets
    $attach2 = $ExcelWorkbook.Sheets.Item("Attachment2")
    $attach1 = $ExcelWorkbook.Sheets.Item("Attachment1")
    $al = $ExcelWorkbook.Sheets.Item("A&L")

    nextPage
    Fill1770S1 -excelSheet $attach1

    # nextPage
    # Fill1770S -excelSheet $al
    # previousPage
    # Start-Sleep 1000


    previousPage
    Fill1770S2 -excelSheet $attach2
    ClearAssetLiability
    SendStatus -status "pending" -percentage "40" -emailSent "false" -docid $FileName -message "Completed - Setting the form to default"

    #Process Assets
    $assetArray = Get-AssetArray -ExcelSheet $al
    ProcessAssets -ObjectArray $assetArray
    SendStatus -status "pending" -percentage "60" -emailSent "false" -docid $FileName -message "Completed - Form 1770 S-I - Assets"

    #Process Liabilitys
    $liabilityArray = Get-LiabilityArray -ExcelSheet $al
    ProcessLiability -ObjectArray $liabilityArray  
    SendStatus -status "pending" -percentage "70" -emailSent "false" -docid $FileName -message "Completed - Form 1770 S-I - Liabilities"
   
   
    #Process Dependants
    $dependentArray = Get-DependentArray -ExcelSheet $attach2
    ProcessDependents -ObjectArray $dependentArray
    SendStatus -status "pending" -percentage "80" -emailSent "false" -docid $FileName -message "Completed - Form 1770 S-I - Dependants"

}



#Main method

# Get all subfolders in the main folder
$subFolders = Get-ChildItem -Path $mainFolderPath -Directory

# Loop through each subfolder
foreach ($subFolder in $subFolders) {

    # Backup of the Folder
    Copy-Item -Path $subFolder.FullName -Destination $BackupFolderPath -Recurse -Force

    # Display filenames in the current subfolder
    $files = Get-ChildItem -Path $subFolder.FullName

    foreach ($file in $files) {
        if ($file.Extension -eq '.xlsb') {
            $FileName = $file.Name
            $updatedFileName = $FileName.Replace(".xlsb", "")
            SendStatus -status "pending" -percentage "0" -emailSent "false" -docid $updatedFileName -message "Started PDF Automation Process"
            log "Subfolder $($subFolder.Name) contains Excel: $($file.Name)"
            log "EXCEL File - $($file.FullName)"
            $excel = New-Object -Com Excel.Application
            $importFile = $file.FullName
            # Open the importfile for edit.
            log $importFile
            $excelData = $excel.Workbooks.Open($importFile)
        }
    }

    foreach ($file in $files) {
        if ($file.Extension -eq '.pdf') {
            log "Subfolder $($subFolder.Name) contains PDF: $($file.Name)"
            log "PDF File - $($file.FullName)"
            $FileName = $file.Name
            $updatedFileName = $FileName.Replace(".pdf", "")
            openPDF -FilePath $file.FullName -FolderName $subFolder.FullName
            fillPDF -ExcelWorkbook $excelData -FileName $updatedFileName
            SendStatus -status "completed" -percentage "100" -emailSent "false" -docid $updatedFileName -message "Receiving Filled PDF File "
            Stop-ProcessByName -ProcessName "EXCEL.exe"
            savePDF -FileName $file.Name
            Stop-ProcessByName -ProcessName "EXCEL.exe"
            sendPDF -FileName $updatedFileName
            Send-GoogleChatNotification -Message "Completed PDF Filling for $($FileName)"
            SendStatus -status "completed" -percentage "100" -emailSent "true" -docid $updatedFileName -message "An email containing the filled PDF has been successfully sent to your registered email address"
            #sendEmail
        }
    }


    # Delete the subfolder
    Remove-Item -Path $subFolder.FullName -Force -Recurse

    # Print a separator for better readability
    log ("-" * 40)

   
}

log "Script completed!"

Exit


