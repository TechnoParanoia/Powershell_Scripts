$Excel = New-Object -Com Excel.Application 
$Excel.Visible = $True 
#creates a new workbook 
$WorkBook = $Excel.WorkBooks.Add() 
$WorkSheet = $WorkBook.WorkSheets.Item(1)
#$Worksheet.Cells.Item($intRow, $intCol) = $strObj.DisplayName
$intCol = 1
$intRow = 1


#List of Groups
$DGI = Read-Host "What is the Distribution Group Name"


Write-host -ForegroundColor Yellow "Looking for $DGI"

$DG = Get-DistributionGroupMember $DGI

$DG

$Worksheet.Cells.Item($intRow, $intCol) = $DGI
$Worksheet.Cells.Item($intRow, $intCol).Interior.ColorIndex =48
$Worksheet.Cells.Item($intRow, $intCol).Font.Bold=$True

$intCol = 2
$intRow = $intRow + 1 


ForEach ($Group in $DG){
    If ($Group.RecipientType -like "*group*") {
                #Write-host -ForegroundColor Blue "Found a Group"
                $Worksheet.Cells.Item($intRow, $intCol) = $Group.Name
                $Worksheet.Cells.Item($intRow, $intCol).Interior.ColorIndex =48
                $Worksheet.Cells.Item($intRow, $intCol).Font.Bold=$True
                $intRow = $intRow + 1
                $GroupMember = Get-DistributionGroupMember $Group.Name

            ForEach ($Member in $GroupMember) {
                $Worksheet.Cells.Item($intRow, 2) = $Member.DisplayName
                $Worksheet.Cells.Item($intRow, 3) = $Member.PrimarySmtpAddress
                $intRow = $intRow + 1
                             }
                              

          }

    Else {

    #Write-host -ForegroundColor Blue "Found Users"
        
    
        $Worksheet.Cells.Item($intRow, 2) = $Group.DisplayName
        $Worksheet.Cells.Item($intRow, 3) = $Group.PrimarySmtpAddress

        }

        $intRow = $intRow + 1


                        }