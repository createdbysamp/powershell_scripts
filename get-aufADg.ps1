# script that loops through all AD Users and returns a list of duplicates and SamAccountName for SSO and comparisons

param (
  [Parameter(Mandatory = $true)]
  [string]$GroupName,

  [string]$OutputPath = ".\AD-Users_.csv",

  [Parameter(Mandatory = $false)]
  [ValidateSet('txt', 'csv', 'json')]
  [string]$Format = 'csv',
  
)

function Get-AllADUsersFromGroup {
  param (
    [string]$GroupName
  )

  $members = Get-ADGroupMember -Identity $GroupName -Recursive

  $seen = @{}

  $users = foreach (member in $members) {
  if ($member.objectClass -eq 'user') {
    if (-not $seen.ContainsKey($member.SamAccountName)) {
      $seen[$member.SamAccountName] = 1
      Get-ADUser -Identity $member.SamAccountName -Properties DisplayName, SamAccountName, Title, EmailAddress, Enabled
    }
    else {
      $seen[$member.SamAccountName]++
    }
  }
  }
  $duplicates = $seen.GetEnumerator() | Where-Object { $_.Value -gt 1 }
  if ($duplicates) {
    Write-Host "`n[WARN] Duplicate member found:"
    foreach ($d in $duplicates) {
      Write-Host " $(d.Key) - seen $($d.Value)x"
    }
  }

  return $users
}

function Print-AllADUsers {
  param (
    [array]$Users
  )

  foreach ($user in $users) {
    Write-Host "$($user.DisplayName) | $($user.SamAccountName) | $($user.Title) | $($user.EmailAddress) | $($user.Enabled)
  }
}

function Export-UsersToFile {

}
