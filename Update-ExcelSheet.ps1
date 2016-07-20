#Functions purpose is to get the latest employee number from an excel list. Then to populate the excel sheet with some information regarding the user.
$PathToSites = 'C:\testar\Excel\Anställningsnummer Solutions_Copy.xlsx'


$ImportedExcel = Import-XLSX -Path $PathToSites -Sheet 'NYA Anstnr' -Header Number, UserName, Creator, Company, Ticketnr, Misc
#Counts the number of Companys written. 
$CompanyCount = ($ImportedExcel.Company | Measure-Object).Count
Write-Output "$CompanyCount" 

$EmployeeNumber = ($ImportedExcel | Select-String -Pattern "$CompanyCount" | Select-String -Pattern 'Company=;' -List)
Write-Output "$($EmployeeNumber[0])"
$EmployeeNumber = (($EmployeeNumber)[0] -replace '\D')[0..4] -join ''
Write-Output "$EmployeeNumber"



$ExcelObject = New-Excel -Path 'C:\testar\Excel\Anställningsnummer Solutions_Copy.xlsx'
#The offset is needed because of the header and stuff in the excel sheet. 
$OffsetCompanyCount = $CompanyCount + 4

$CellACoordinates = "a$OffsetCompanyCount`:a$OffsetCompanyCount"
$CellBCoordinates = "b$OffsetCompanyCount`:b$OffsetCompanyCount"
$CellCCoordinates = "c$OffsetCompanyCount`:c$OffsetCompanyCount"	

$ExcelWorkBook = $ExcelObject | Get-Workbook -Verbose
$ExcelWorksheet = $ExcelObject | Get-Worksheet -Name 'NYA anstnr'

$CellValueA = (Get-CellValue -Excel $ExcelObject -Coordinates $CellACoordinates -WorkSheetName $ExcelWorksheet).Anstnr
$CellValueB = (Get-CellValue -Excel $ExcelObject -Coordinates $CellBCoordinates -WorkSheetName $ExcelWorksheet).Användarnamn
$CellvalueC = (Get-CellValue -Excel $ExcelObject -Coordinates $CellCCoordinates -WorkSheetName $ExcelWorksheet).'Registrerat av'

Write-Output "CellValueA: $CellValueA"
Write-Output "CellValueB: $CellValueB"
Write-Output "CellValueC: $CellvalueC"

#Checks to see if the Cell it picks is the correct one and if the rest of the cells are empty. If they arent it means that it is about to overwrite something which we dont want.
if (($CellValueA -eq $EmployeeNumber) -and ($CellValueB -eq $null) -and ($CellvalueC -eq $null))
{
  Write-Host 'hm'
}
else
{
  #This else should log what happened first. The two variables wasnt the same. Suggestion is to log what they both were. 
  #Should also maybe send a mail that the script wasnt successfull.
  exit
}
