# Import the module
Import-Module ImportExcel

$excel = New-Object -Com Excel.Application
$importFile = "C:\FORM1770S\Test Folder\1 - Copy\excel.xlsb"

# Open the import file for edit.
log $importFile
$excelData = $excel.Workbooks.Open($importFile)
$al = $excelData.Sheets.Item("A&L")
$attach2 = $excelData.Sheets.Item("Attachment2")

# Starting cell coordinates
$startRow = 7
$codeColumn = 3
$descriptionColumn = 4
$amountColumn = 16
$typeColumn = 5

# Arrays to store objects for each row type
$assetArray = @()
$liabilityArray = @()

for ($row = $startRow; $true; $row++) {
    $code = $al.Cells.Item($row, $codeColumn).Text -replace '[^\d]'
    
    # Check if the code is empty (assuming the code is the key column)
    if ($code -eq '') {
        break
    }

    $description = $al.Cells.Item($row, $descriptionColumn).Text
    $amount = $al.Cells.Item($row, $amountColumn).Text -replace '[^\d]'
    $type = $al.Cells.Item($row, $typeColumn).Text

    # Create an object for the current row
    $rowObject = [PSCustomObject]@{
        Code = $code
        Description = $description
        Amount = $amount
        Type = $type
    }

    # Push the object to the appropriate array based on type
    if ($type -eq 'Asset') {
        $assetArray += $rowObject
    } elseif ($type -eq 'Liability') {
        $liabilityArray += $rowObject
    }
}

# Display the list of Asset objects
log "List of Asset Rows:"
$assetArray | ForEach-Object { Log $_ }

# Display the list of Liability objects
log "List of Liability Rows:"
$liabilityArray | ForEach-Object { Log $_ }


    # Starting cell coordinates
$startRow1 = 25
$nameColumn = 4
$idColumn = 5
$relationColumn = 6
$occupationColumn = 7


# Arrays to store objects for each row type
$dependentArray = @()

for ($row = $startRow1; $true; $row++) {
    $name = $attach2.Cells.Item($row, $nameColumn).Text
    
    # Check if the code is empty (assuming the code is the key column)
    if ($name -eq '') {
        break
    }

    $id = $attach2.Cells.Item($row, $idColumn).Text
    $relation = $attach2.Cells.Item($row, $relationColumn).Text 
    $occupation = $attach2.Cells.Item($row, $occupationColumn).Text

    # Create an object for the current row
    $rowObject = [PSCustomObject]@{
        Name = $name
        ID = $id
        Relation = $relation
        Occupation = $occupation
    }

    # Push the object to the appropriate array based on type
   
        $dependentArray += $rowObject
}

# Display the list of Asset objects
log "List of Dependent Rows:"
$dependentArray | ForEach-Object { Log $_ }




# Close Excel
$excelData.Close()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
