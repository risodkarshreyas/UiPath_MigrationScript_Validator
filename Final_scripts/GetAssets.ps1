# Install ImportExcel if not already installed
Install-Module ImportExcel -Scope CurrentUser -Force

# === Authentication ===
$tokenheaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$tokenheaders.Add("Content-Type", "application/json")

$body = @"
{
  `"tenancyName`": `"<yourTenant>`",
  `"usernameOrEmailAddress`": `"apiuser`",
  `"password`": `"<yourPassword>`"
}
"@

$response = Invoke-RestMethod 'https://<yourASFQDN>/default/<yourTenant>/orchestrator_/api/Account/Authenticate' -Method 'POST' -Headers $tokenheaders -Body $body
$token = $response.result

# === Headers and Base URL ===
$headers = @{ Authorization = "Bearer $token" }
$headers["X-UIPATH-TenantName"] = "<yourTenant>"
$accountLogicalName = "default"
$logicalTenant = "17531af8-128b-415a-9bad-c56bd4b211bc"
$baseUrl = "https://<yourASFQDN>/$accountLogicalName/$logicalTenant/odata"

# === Get Folders ===
$FoldersUrl = "$baseUrl/Folders"
$FoldersResponse = Invoke-RestMethod -Uri $FoldersUrl -Headers $headers
$Folders = $FoldersResponse.value

# === Initialize Results Array ===
$AssetResults = @()

foreach ($Folder in $Folders) {
    $Folderheaders = @{ Authorization = "Bearer $token" }
    $Folderheaders["X-UIPATH-OrganizationUnitId"] = $Folder.Id

    $assetsUrl = "https://<yourASFQDN>/$accountLogicalName/<yourTenant>/orchestrator_/odata/Assets"

    try {
        $resp = Invoke-RestMethod -Uri $assetsUrl -Headers $Folderheaders
        foreach ($asset in $resp.value) {
            $valueCode = switch ($asset.ValueType) {
                "Credential" { "username: $($asset.CredentialUsername)" }
                default      { $asset.Value }
            }

            $AssetResults += [PSCustomObject]@{
                FolderName = $Folder.FullyQualifiedName
                AssetName  = $asset.Name
                Description = $asset.Description
                Scope      = $asset.ValueScope
                Type       = $asset.ValueType
                ValueCode  = $valueCode
            }
        }
    } catch {
        Write-Warning "Failed to fetch assets for folder $($Folder.FullyQualifiedName): $_"
    }
}

# === Export to Excel ===
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$excelPath = "AssetsFolderWise_$timestamp.xlsx"
$AssetResults | Export-Excel -Path $excelPath -WorksheetName "Assets" -AutoSize -BoldTopRow

Write-Output "Excel report exported successfully to $excelPath."
