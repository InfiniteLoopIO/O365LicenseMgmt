# Requires -Modules MSOnline

<#
.Synopsis
   Return O365 account assigned licenses
.DESCRIPTION
   Using a supplied UserPrincipalName determine if the O365 account exists and return current license(s) 
   Returned strings will be a semi-colon separated list of O365 Licenses, Unlicensed, or O365 Account Not Found
.EXAMPLE
   Get-365License -UPN sample.name@contoso.com
#>
function Get-365License
{
  [cmdletbinding()]
  PARAM(
    [Parameter(Mandatory=$true,Position=0)]
    [ValidateScript({$_ -match '^([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5})$'})]
    [string]$UPN
  )

  try
  {
    $MSOLuser = Get-MsolUser -UserPrincipalName $UPN -ea Stop
    if($MSOLuser.IsLicensed) 
    {
      ($MSOLuser.Licenses | % {$_.accountskuid}) -join ';'
    } 
    else 
    {
      "Unlicensed"
    }
  }
  catch
  {
    "O365 Account Not Found"
    Write-Debug $_.exception.message
  }
}

<#
.Synopsis
   Add O365 License to an account
.DESCRIPTION
   Using a supplied UserPrincipalName determine if the O365 account exists and attempt to assign valid SKU
   license.  The accounts usage location will be defaulted to US if not set and not specified.
   Informational messages are displayed regarding the license assignment status but nothing is returned 
.EXAMPLE
   Add-365License -UPN sample.name@contoso.com -LicenseType 'tenantname:ENTERPRISEPACK'
#>
function Add-365License
{
  [cmdletbinding()]
  PARAM(
    [Parameter(Mandatory=$true,Position=0)]
    [ValidateScript({$_ -match '^([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5})$'})]
    [string]$UPN,
    [Parameter(Mandatory=$true,Position=1)]
    [ValidateScript({$global:O365LicenseSet -contains $_})]
    [string]$LicenseType,
    [Parameter(Mandatory=$false,Position=2)]
    [string]$Location='US'
  )

  try
  {
    $MSOLuser = Get-MsolUser -UserPrincipalName $UPN -ea Stop

    if(Test-365License $MSOLuser $LicenseType)
    {
      Write-Host "$LicenseType already assigned to $($MSOLuser.UserPrincipalName), no changes made" -f Magenta
    }
    else
    {
      Write-Host "Adding $LicenseType to $($MSOLuser.UserPrincipalName)" -f Green
      if($MSOLuser.UsageLocation -ne $Location)
      {
        Write-Host "Usage location not set, using $location" -f Green
        try{Set-MsolUser -UserPrincipalName $MSOLuser.UserPrincipalName -UsageLocation $location -ea Stop}catch{write-host $_.exception.message -ForegroundColor Red}
      }
      try{Set-MsolUserLicense -user $MSOLuser.UserPrincipalName -AddLicenses $LicenseType -ea Stop}catch{write-host $_.exception.message -ForegroundColor Red}
    }
  }
  catch
  {
    Write-Host $_.exception.message -ForegroundColor Red
  }
}

<#
.Synopsis
   Remove O365 license from an account
.DESCRIPTION
   Using a supplied UserPrincipalName determine if the O365 account exists and attempt to remove a valid SKU
   license.
   Informational messages are displayed regarding the license removal status but nothing is returned 
.EXAMPLE
   Remove-365License -UPN sample.name@contoso.com -LicenseType 'tenantname:ENTERPRISEPACK'
#>
function Remove-365License
{
  [cmdletbinding()]
  PARAM(
    [Parameter(Mandatory=$true,Position=0)]
    [ValidateScript({$_ -match '^([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5})$'})]
    [string]$UPN,
    [Parameter(Mandatory=$true,Position=1)]
    [ValidateScript({$global:O365LicenseSet -contains $_})]
    $LicenseType
  )

  try
  {
    $MSOLuser = Get-MsolUser -UserPrincipalName $UPN -ea Stop

    if($MSOLuser.IsLicensed)
    {
      if(Test-365License $MSOLuser $LicenseType)
      {
        Write-Host "Removing $LicenseType from $($MSOLuser.UserPrincipalName)" -f Green
        try{Set-MsolUserLicense -user $MSOLuser.UserPrincipalName -RemoveLicenses $LicenseType -ea stop}catch{Write-Host $_.exception.message -ForegroundColor Red}
      }
      else
      {
        Write-Host "$LicenseType not assigned to $($MSOLuser.UserPrincipalName), no changes made" -f Magenta
      }
    }
    else 
    {
      Write-Host "$($MSOLuser.UserPrincipalName) was found but currently unlicensed" -f Yellow
    }
  }
  catch
  {
    Write-Host $_.exception.message -ForegroundColor Red
  }
}

<#
.Synopsis
   Check if O365 account is already assigned a license
.DESCRIPTION
   A helper function for the other functions within the module.  Given an O365 user account object and a SKU license
   check to see if the O365 account currently has the license assigned.
   Return TRUE if account is assigned license and FALSE if account is not currently licensed for the supplied value
.EXAMPLE
   Test-365License -MSOLuser $userObj -LicenseType 'tenantname:ENTERPRISEPACK'
#>
function Test-365License
{
[cmdletbinding()]
  PARAM(
    [Parameter(Mandatory=$true,Position=0)]
    [Microsoft.Online.Administration.User]$MSOLuser,
    [Parameter(Mandatory=$true,Position=1)]
    [ValidateScript({$global:O365LicenseSet -contains $_})]
    $LicenseType
  )
    
  $currentLicense += $MSOLuser.Licenses | % {$_.accountSkuId}
  if($currentLicense -contains $LicenseType) {$true} else {$false}
}

# Discover Tenant License SKU and place in global variable to be used with function validate scripts
$global:O365LicenseSet = (Get-MsolAccountSku -ea 0).accountskuid
while($global:O365LicenseSet -eq $null)
{
   Write-Host "Unable to get O365 Tenant Account SKU, cmdlets will not work properly" -ForegroundColor Yellow
   Connect-MsolService
   $global:O365LicenseSet = (Get-MsolAccountSku -ea 0).accountskuid
}
Write-Host "O365 Account SKU discovered" -ForegroundColor Green