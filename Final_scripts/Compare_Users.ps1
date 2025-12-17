# Load both sheets -- modify Path as per your liking
$cloudData = Import-Excel -Path "C:\Users\sbondugula\UiPath_Folder_Roles_FINANCE_ENY_PROD.xlsx" -WorksheetName "Cloud Folder Roles"
$suiteData = Import-Excel -Path "C:\Users\sbondugula\UiPath_Folder_Roles_FINANCE_ENY_PROD.xlsx" -WorksheetName "Suite Folder Roles"

# Create a lookup map for Suite roles
$suiteMap = @{}
foreach ($row in $suiteData) {
    $key = "$($row.'Folder Name')|$($row.Username)"
    $roles = ($row.Roles -split ',') | ForEach-Object { $_.Trim() }
    $suiteMap[$key] = $roles
}

# Compare and compute symmetric differences
foreach ($row in $cloudData) {
    $key = "$($row.'Folder Name')|$($row.Username)"
    $cloudRoles = ($row.Roles -split ',') | ForEach-Object { $_.Trim() }
    $suiteRoles = $suiteMap[$key]

    if ($suiteRoles) {
        $diff = ($cloudRoles + $suiteRoles | Sort-Object | Get-Unique) | Where-Object { ($_ -notin $cloudRoles) -or ($_ -notin $suiteRoles) }
        $row | Add-Member -NotePropertyName "Role Difference" -NotePropertyValue ($diff -join ", ")
    } else {
        $row | Add-Member -NotePropertyName "Role Difference" -NotePropertyValue "Missing in Suite"
    }
}

# Export updated Cloud sheet
$cloudData | Export-Excel -Path "UiPath_Folder_Roles_FINANCE_ENY_PROD.xlsx" -WorksheetName "Cloud Folder Roles" -AutoSize -BoldTopRow

Write-Output "Role differences written to 'Cloud Folder Roles' sheet."
