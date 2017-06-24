# O365LicenseMgmt
Powershell module to assist with O365 account license discovery, assignment, and removal

## Synopsis
The module cmdlets include Get-365License, Add-365License, and Remove-365License.  The cmdlets will verify the license is valid for your
tenant, verify the MS Online Account exists, perform requested actions, and provide status updates to the conosle.  I initially provided 
validated MS Online Account objects to the cmdlets which allowed for TRUE|FALSE return based on task status but that introduced complications 
when multiple license actions were being performed without refreshing current license status.  This version may be slightly slower than the 
prototype but is more flexible for individual tasks.

## Usage

1) Clone to local repository
2) Create C:\Windows\System32\WindowsPowerShell\v1.0\Modules\O365LicenseMgmt
3) Copy O365LicenseMgmt.psm1 to C:\Windows\System32\WindowsPowerShell\v1.0\Modules\O365LicenseMgmt\
4) Open Powershell or Powershell ISE as administrator
5) Connect-MSOnline
6) Import-Module O365LicenseMgmt
7) Assign array of UserPrincipalName from your tenant to $UPNarr or perform single actions

### List valid license SKU for your tenant
`(Get-MsolAccountSku).AccountSkuId`

### Simple report
`$UPNarr | select @{n='UPN';e={$_}},@{n='365License';e={Get-365License -UPN $_}}`

### Remove any assigned license(s)
`$UPNarr | % {$acct = $_; (Get-365License -UPN $acct ).split(';') | % {if($_ -match ':'){Remove-365License -UPN $acct -LicenseType $_}}}`

### Replace E3 with E1 license type
`$UPNarr | % {Remove-365License -UPN $_ -LicenseType 'tenant:ENTERPRISEPACK'; Add-365License -UPN $_ -LicenseType 'tenant:STANDARDPACK'}`

## License

BSD 2-Clause License

Copyright (c) 2017, https://infiniteloop.io
All rights reserved.
