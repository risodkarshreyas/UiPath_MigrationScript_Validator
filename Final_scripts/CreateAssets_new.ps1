# Install ImportExcel if not already installed
#Install-Module ImportExcel -Scope CurrentUser -Force

# === Authentication ===
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Content-Type", "application/x-www-form-urlencoded")

$body = "grant_type=client_credentials&client_id=9fe53278-7de5-4c89-a9fb-3a2ef18cda92&client_secret=zDoybr0S0N6~7AWp&scope=OR.Default"

$response = Invoke-RestMethod 'https://cloud.uipath.com/<yourOrg>/identity_/connect/token' -Method 'POST' -Headers $headers -Body $body
$token = $response.access_token

# === Setup Headers and Base URL ===
$accountLogicalName = "<yourOrg>"
$tenantName = "<yourTenant>"
$logicalTenant = "f9168dbc-d471-4c3f-960f-bfe13a61993e"
$baseUrl = "https://cloud.uipath.com/$accountLogicalName/$logicalTenant/odata"
$headers = @{ Authorization = "Bearer $token" }
$headers["X-UIPATH-TenantName"] = $tenantName

# === Load Excel ===
$excelPath = "C:\Users\sbondugula\AssetsFolderWise_20250726_0426.xlsx"  # Update with your actual file name
$assetsFromExcel = Import-Excel -Path $excelPath -WorksheetName "Assets"

# Add Status column if not present
if (-not ($assetsFromExcel | Get-Member -Name "Status")) {
    $assetsFromExcel | ForEach-Object { $_ | Add-Member -NotePropertyName "Status" -NotePropertyValue "" }
}

# === Get All Folders ===
$foldersUrl = "$baseUrl/Folders"
$folders = (Invoke-RestMethod -Uri $foldersUrl -Headers $headers).value

foreach ($folder in $folders) {
    $orgUnitId = $folder.Id
    $folderName = $folder.FullyQualifiedName
    Write-Output "Processing folder: $folderName"

    $folderHeaders = @{ Authorization = "Bearer $token"; "X-UIPATH-OrganizationUnitId" = $orgUnitId }

    # Get existing assets in this folder
    $existingAssetsUrl = "https://cloud.uipath.com/$accountLogicalName/$tenantName/orchestrator_/odata/Assets"
    try {
        $existingAssets = (Invoke-RestMethod -Uri $existingAssetsUrl -Headers $folderHeaders).value
        $existingAssetNames = $existingAssets.Name
    } catch {
        Write-Warning "Failed to fetch existing assets for folder '$folderName': $_"
        continue
    }

    # Filter assets for this folder
    $folderAssets = $assetsFromExcel | Where-Object { $_.FolderName -eq $folderName }

    foreach ($asset in $folderAssets) {
        if ($existingAssetNames -contains $asset.AssetName) {
            $asset.Status = "Already Exists"
            continue
        }

        # Build asset payload
        $payload = @{
            Name = $asset.AssetName
            ValueScope = $asset.Scope
            ValueType = $asset.Type
            Description = $asset.Description
            CanBeDeleted = $true
        }

        switch ($asset.Type) {
            "Text" {
                $payload.StringValue = $asset.ValueCode
            }
            "Bool" {
                try {
                    $payload.BoolValue = [bool]::Parse($asset.ValueCode)
                } catch {
                    $asset.Status = "Error: Invalid Bool value"
                    continue
                }
            }
            "Integer" {
                try {
                    $payload.IntValue = [int]$asset.ValueCode
                } catch {
                    $asset.Status = "Error: Invalid Integer value"
                    continue
                }
            }
            "Credential" {
                $payload.CredentialUsername = $asset.ValueCode -replace "^username:\s*", ""
                $payload.CredentialPassword = ""
                $payload.CredentialStoreId = 38  # Update if needed
            }
            "Secret" {
                $payload.SecretValue = $asset.ValueCode
            }
        }

        # Create asset
        $createUrl = "https://cloud.uipath.com/$accountLogicalName/$tenantName/orchestrator_/odata/Assets"
        try {
            Invoke-RestMethod -Uri $createUrl -Method Post -Headers $folderHeaders -Body ($payload | ConvertTo-Json -Depth 10) -ContentType "application/json"
            $asset.Status = "Created"
            Write-Output "Created asset '$($asset.AssetName)' in folder '$folderName'"
        } catch {
            $errorMessage = $_.Exception.Message
            $asset.Status = "Error: $errorMessage"
            Write-Warning "Failed to create asset '$($asset.AssetName)' in folder '$folderName': $errorMessage"
        }
    }
}

# === Export Updated Excel ===
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$updatedPath = "CreateAssets_Status_$timestamp.xlsx"
$assetsFromExcel | Export-Excel -Path $updatedPath -WorksheetName "Assets" -AutoSize -BoldTopRow

Write-Output "Updated Excel with status exported to $updatedPath"
