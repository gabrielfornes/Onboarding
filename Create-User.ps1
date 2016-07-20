function Format-Ticket
{
  [cmdletbinding()]
    param
    (
      [parameter(mandatory=$true)]
      [string]$Path
    )

  $TextFile = Get-Content C:\testar\Create_users\user.txt
  #Removes all the hashtags and creates an array of the groups.
  $TicketNumber       = (($TextFile | Select-String -Pattern 'TicketNumber') -replace '##Ticketnumber##', '').Trim()
  $UserNumber         = (($TextFile | Select-String -Pattern 'UserNumber') -replace '##UserNumber##', '').Trim()
  $FirstName          = (($TextFile | Select-String -Pattern 'Firstname') -replace '##FirstName##', '').Trim()
  $Lastname           = (($TextFile | Select-String -Pattern 'LastName') -replace '##LastName##', '').Trim()
  $MiddleName         = (($TextFile | Select-String -Pattern 'MiddleName') -replace '##MiddleName##', '').Trim()
  $Company            = (($TextFile | Select-String -Pattern 'Company') -replace '##Company##', '').Trim()
  $Managedby          = (($TextFile | Select-String -Pattern 'ManagedBy') -replace '##managedBy##', '').Trim()
  $HomeFolder         = (($TextFile | Select-String -Pattern 'HomeFolder') -replace '##HomeFolder##', '').Trim()
  $Mail               = (($TextFile | Select-String -Pattern 'Mail') -replace '##Mail##', '').Trim()
  $SecurityGroups     = ((($TextFile | Select-String -Pattern 'SecurityGroups') -replace '##SecurityGroups##', '') -split ',', '').Trim()
  $DistributionGroups = ((($TextFile | Select-String -Pattern 'DistributionGroups') -replace '##DistributionGroups##', '') -split ',', '').Trim()
  
  #Start Verbose
  Write-Verbose $TicketNumber
  Write-Verbose $UserNumber
  Write-Verbose $FirstName
  Write-Verbose $Lastname
  Write-Verbose $MiddleName
  Write-Verbose $Company
  Write-Verbose $Managedby
  Write-Verbose $HomeFolder
  Write-Verbose $Mail
  Write-Verbose (($SecurityGroups | Out-String).Trim())
  Write-Verbose (($DistributionGroups | Out-String).Trim())
  #End Verbose

  New-Employee -TickettNumber $TicketNumber -UserNumber $UserNumber -FirstName $FirstName -LastName $Lastname -MiddleName $MiddleName -Company $Company -Managedby $Managedby -HomeFolder -Mail -DistributionGroups $DistributionGroups -SecurityGroups $SecurityGroups


 }

 function New-Employee
 {
     [CmdletBinding()]     
     
     Param
     (
         # Param1 help description
         [Parameter(Mandatory=$true)]
         [string]$TickettNumber,
 
         [Parameter(Mandatory=$true)]
         [string]$UserNumber,

         [Parameter(Mandatory=$true)]
         [string]$FirstName,

         [Parameter(Mandatory=$true)]
         [string]$LastName,

         [Parameter(Mandatory=$true)]
         [string]$MiddleName,

         [Parameter(Mandatory=$true)]
         [string]$Company,

         [Parameter(Mandatory=$true)]
         [string]$Managedby,
         
         [Parameter(Mandatory=$true)]
         [string[]]$SecurityGroups,

         [Parameter(Mandatory=$true)]
         [string[]]$DistributionGroups,

         [Parameter()]
         [switch]$HomeFolder,

         [Parameter()]
         [switch]$Mail

        
      )

      Write-Output $TickettNumber
      $UserNumber
      $FirstName
      $LastName
      $MiddleName
      $Company
      $Managedby
      $HomeFolder
      $Mail
      $SecurityGroups
      $DistributionGroups

 
     
 }
 