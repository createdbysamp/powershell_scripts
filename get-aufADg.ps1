# script that loops through all AD Users and returns a list of duplicates and SamAccountName for SSO and comparisons

param (
  [Parameter(Mandatory = $true)]
  [string]$GroupName,

  [string]$OutputPath = ".\AD-Users_.csv",

  [Parameter(Mandatory = $false)]
  [ValidateSet('txt', 'csv', 'json')]
  [string]$Format = 'csv',
  
)
