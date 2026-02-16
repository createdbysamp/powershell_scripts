# install if not already
#Install-Module ImportExcel -Scope CurrentUser -Force

# import module
Import-Module ActiveDirectory
Import-Module ImportExcel

# variable declaration
$excelPath = "C:\Users\"
$backupPath = $excelPath -replace '\.xlsx$', '_backup.xlsx'

Copy-Item -Path $backupPath -Destination $excelPath

$sheetName = "Sheet1"

##FUNCTIONS

function Parse-Login {
    param ( 
        [string]$login
    )  

    if ([string]::IsNullOrWhiteSpace($login)) {
        return $null
    }

    $login = $login.Trim()
    
    if ($login -like '*\*') {
        $parts = $login -split '\\'
        $domain = $parts[0] # "domain_name"
        $username = $parts[1] # "E911833"
    }
    else {
        $domain = $null
        $username = $login
    }

    if ([string]::IsNullOrWhiteSpace($username)) {
        return $null
    }

    # returns hashtable with properties 
    return @{
        Domain   = $domain
        Username = $username
    }
    # Write-Host "Domain: $domain"
    # Write-Host "Username: $username"
}

function Check-Dups {
    param (
        [string]$username,
        [hashtable]$seenUsers
    )

    return $seenUsers.ContainsKey($username)
}

function Add-Rows {
    param (
        [PSCustomObject]$row
    )
    
    # Add new columns to this row if they don't exist
    if (-not ($row.PSObject.Properties.Name -contains 'FirstName')) {
        $row | Add-Member -NotePropertyName 'FirstName' -NotePropertyValue $null -Force
    }
    if (-not ($row.PSObject.Properties.Name -contains 'LastName')) {
        $row | Add-Member -NotePropertyName 'LastName' -NotePropertyValue $null -Force
    }
    if (-not ($row.PSObject.Properties.Name -contains 'UPN')) {
        $row | Add-Member -NotePropertyName 'UPN' -NotePropertyValue $null -Force
    }
}

##SCRIPT LOGIC

# hash table init
$seenUsers = @{} 

# read the excel file
$rows = Import-Excel -Path $excelPath -WorksheetName $sheetName

##LOOP LOGIC
# loop through each row
Write-Host "Starting Loop through Rows ..."

foreach ($row in $rows) {

    $login = $row.LoginName

    Add-Rows -row $row

    $userInfo = Parse-Login -login $login

    if ($null -eq $userInfo) {
        Write-Host "...no user available, continuing"
        continue
    }

    #checking for duplicates
    $isDuplicate = Check-Dups -username $userInfo.Username -seenUsers $seenUsers

    if ($isDuplicate) {
        Write-Host "...previously logged user, skipping"
        continue
    }
    # mark user as seen
    $seenUsers[$userInfo.Username] = $true 

    # TODO: ad query here
    try {
        $user = Get-AdUser -Identity $userInfo.Username -Properties GivenName, Surname, UserPrincipalName -ErrorAction Stop
        Write-Host "User found: $($user.SamAccountName)"

        # add results back to spreadsheet
        $row.FirstName = $user.GivenName
        $row.LastName = $user.Surname
        $row.UPN = $user.UserPrincipalName

        Write-Host "UserInfo added to spreadsheet."
    }

    catch {
        $row.UPN = $userInfo.Username
        Write-Host "Error: $($_.Exception.Message)"
    }
    
    Write-Host "Processed: $login"
}

# TODO: export back to excel
$rows | Export-Excel -Path $excelPath -WorksheetName $sheetName -FreezeTopRow -BoldTopRow
