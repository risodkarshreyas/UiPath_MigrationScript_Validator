# Install ImportExcel if not already installed
#Install-Module ImportExcel -Scope CurrentUser -Force

# === Authentication ===
$tokenheaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$tokenheaders.Add("Content-Type", "application/json")

$body = @"
{
  "tenancyName": "<yourTenant>",
  "usernameOrEmailAddress": "apiuser",
  "password": "<yourpwd>"
}
"@

$response = Invoke-RestMethod 'https://<AS_SourceFQDN>/default/<yourTenant>/orchestrator_/api/Account/Authenticate' -Method 'POST' -Headers $tokenheaders -Body $body
$token = $response.result

# === Headers and Base URL ===
$headers = @{ Authorization = "Bearer $token" }
$headers["X-UIPATH-TenantName"] = "<yourTenant>"
$accountLogicalName = "default"
$logicalTenant = "17531af8-128b-415a-9bad-c56bd4b211bc"
$baseUrl = "https://<AS_SourceFQDN>/$accountLogicalName/$logicalTenant/odata"

# === Get Folders ===
$FoldersUrl = "$baseUrl/Folders"
$FoldersResponse = Invoke-RestMethod -Uri $FoldersUrl -Headers $headers
$Folders = $FoldersResponse.value

# === Initialize Results Array ===
$AssetResults = @()

foreach ($Folder in $Folders) {
    $Folderheaders = @{ Authorization = "Bearer $token" }
    $Folderheaders["X-UIPATH-OrganizationUnitId"] = $Folder.Id

    $assetsUrl = "https://<AS_SourceFQDN>/$accountLogicalName/<yourTenant>/orchestrator_/odata/ProcessSchedules"

    try {
        $resp = Invoke-RestMethod -Uri $assetsUrl -Headers $Folderheaders
        foreach ($asset in $resp.value) {
            $SpecificpriorityValue = $asset.SpecificPriorityValue
            


            $AssetResults += [PSCustomObject]@{
                FolderName = $Folder.FullyQualifiedName
                TriggerName  = $asset.Name
                Jobpriority = $asset.JobPriority
                Enabled = $asset.Enabled
               SpecificPriorityValue= $asset.SpecificPriorityValue

            }
        }
    } catch {
        Write-Warning "Failed to fetch assets for folder $($Folder.FullyQualifiedName): $_"
    }
}

# === Export to Excel ===
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$excelPath = "TriggersFolderwise_Suite_$tenantName.xlsx"
$AssetResults | Export-Excel -Path $excelPath -WorksheetName "Triggers_Suite" -AutoSize -BoldTopRow

Write-Output "Excel report exported successfully to $excelPath."
 

 #-----------------------------------------------------

 $token_headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$token_headers.Add("Content-Type", "application/x-www-form-urlencoded")

$body = "grant_type=client_credentials&client_id=9fe53278-7de5-4c89-a9fb-3a2ef18cda92&client_secret=zDoybr0S0N6~7AWp&scope=OR.Default"

$response = Invoke-RestMethod 'https://cloud.uipath.com/micronorg/identity_/connect/token' -Method 'POST' -Headers $token_headers -Body $body
#$response | ConvertTo-Json
#Write-Output "Access Token: $($response.access_token)"
$token = $response.access_token
$accountLogicalName="micronorg"
$tenantName = "Test_FINANCE"#Tenantname
$logicalTenant = "f9168dbc-d471-4c3f-960f-bfe13a61993e"
$headers = @{ Authorization = "Bearer $token" }
$baseUrl = "https://cloud.uipath.com/$accountLogicalName/$logicalTenant/odata"
$headers["X-UIPATH-TenantName"] = $tenantName
$Folderheaders = @{ Authorization = "Bearer $token" }

# === Get Folders ===
$FoldersUrl = "$baseUrl/Folders"
$FoldersResponse = Invoke-RestMethod -Uri $FoldersUrl -Headers $headers
$Folders = $FoldersResponse.value

# === Initialize Results Array ===
$AssetResults = @()

foreach ($Folder in $Folders) {
    $Folderheaders = @{ Authorization = "Bearer $token" }
    $Folderheaders["X-UIPATH-OrganizationUnitId"] = $Folder.Id

    $assetsUrl = "https://cloud.uipath.com/$accountLogicalName/$tenantName/orchestrator_/odata/ProcessSchedules"

    try {
        $resp = Invoke-RestMethod -Uri $assetsUrl -Headers $Folderheaders
        foreach ($asset in $resp.value) {
            $SpecificpriorityValue = $asset.SpecificPriorityValue
            


            $AssetResults += [PSCustomObject]@{
                FolderName = $Folder.FullyQualifiedName
                TriggerName  = $asset.Name
                Jobpriority = $asset.JobPriority
                Enabled = $asset.Enabled
               SpecificPriorityValue= $asset.SpecificPriorityValue

            }
        }
    } catch {
        Write-Warning "Failed to fetch assets for folder $($Folder.FullyQualifiedName): $_"
    }
}

# === Export to Excel ===
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$excelPath = "TriggersFolderwise_Suite_$tenantName.xlsx"
$AssetResults | Export-Excel -Path $excelPath -WorksheetName "Triggers_Cloud" -AutoSize -BoldTopRow

Write-Output "Excel report exported successfully to $excelPath."