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

$hm = (Get-CellValue -Excel $ExcelObject -Coordinates $CellACoordinates -WorkSheetName $ExcelWorksheet).Anstnr

if ($hm -eq $EmployeeNumber)
{
  Write-Host 'hm'
}
else
{
  #This else should log what happened first. The two variables wasnt the same. Suggestion is to log what they both were. 
  #Should also maybe send a mail that the script wasnt successfull.
  exit
}
