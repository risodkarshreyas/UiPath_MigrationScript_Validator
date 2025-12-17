$var_SuiteTenantName="<YourSourceTenantName>"
$var_CloudTenantName="<YourDesitnationTenantName>"


$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Content-Type", "application/x-www-form-urlencoded")

$body = "grant_type=client_credentials&client_id=9fe53278-7de5-4c89-a9fb-3a2ef18cda92&client_secret=zDoybr0S0N6~7AWp&scope=OR.Default"

$response = Invoke-RestMethod 'https://cloud.uipath.com/<YourOrgName>/identity_/connect/token' -Method 'POST' -Headers $headers -Body $body
#$response | ConvertTo-Json
#Write-Output "Access Token: $($response.access_token)"
$token = $response.access_token
#Write-Output "Authentication successful."

# === Initialize Results Array ===
$results = @()
$Folderresults = @()

$accountLogicalName = "<YourOrgName>"#cloudorg
#$defaultTenant = "DefaultTenant"  # Any valid tenant for auth

$tenantName = $var_CloudTenantName#Tenantname
$logicalTenant = "f9168dbc-d471-4c3f-960f-bfe13a61993e"
$headers = @{ Authorization = "Bearer $token" }
$baseUrl = "https://cloud.uipath.com/$accountLogicalName/$logicalTenant/odata"
$headers["X-UIPATH-TenantName"] = $tenantName
$Folderheaders = @{ Authorization = "Bearer $token" }

function Get-Count($endpoint) 
{
        $url = "$baseUrl/$endpoint`?\$count=true&\$top=1"
        try {
            $resp = Invoke-RestMethod -Uri $url -Headers $headers
            #if($endpoint -eq "Folders")  {$Folders = $resp.value}
            return $resp.'@odata.count'
        } catch {
            Write-Warning "Failed to fetch $endpoint for $tenantName. $_"
            return -1
        }

}

    $result = [PSCustomObject]@{
        Tenant   = $tenantName
        Machines = Get-Count "Machines"
        Users    = Get-Count "Users"
        Folders  = Get-Count "Folders"
        Packages = Get-Count "Processes"
        Roles    = Get-Count "Roles"
        #Queues   = Get-Count "QueueDefinitions"
        #Assets   = Get-Count "Assets"
    }

    $results += $result
    #Write-Output "Added counts for tenant: $tenantName"
    #Write-Output "$result"
    $results | Format-Table -AutoSize
    
    $FoldersUrl = "$baseUrl/Folders"
    $FoldersResponse = Invoke-RestMethod -Uri $FoldersUrl -Headers $headers
    $Folders =  $FoldersResponse.value
    Write-Output ("************************************************")
    Write-Output ("Total Folder Count on this tenant is:" + $FoldersResponse.'@odata.count')
    Write-Output ("************************************************")

    function Get-FolderLevelCount($endpoint) 
    {
        $url = "https://cloud.uipath.com/$accountLogicalName/$tenantName/orchestrator_/odata/$endpoint"
        #Write-Output $url
        #Write-Output $Folderheaders
        try {

            $resp = Invoke-RestMethod -Uri $url -Headers $Folderheaders
            #if($endpoint -eq "Folders")  {$Folders = $resp.value}
            # Write-Output $resp.'@odata.count'
            
            $Count = $resp.'@odata.count'
            return $Count 
            
        } catch {           
            $Exception = $_.Exception
            $ErrorMessage = $Exception.Message
            #Write-Warning $ErrorMessage
            return -1
        }

    }

    $Folderresults = @()

    foreach ($Folder in $Folders) { 
        #$Folderheaders["X-UIPATH-TenantName"] = $tenantName
        #$Folderheaders["X-UIPATH-TenantName"] = $tenantName
        $Folderheaders["X-UIPATH-OrganizationUnitId"] = $Folder.Id
       
        #Write-Output ("Folder Name is " + $Folder.FullyQualifiedName)  
        #Write-Output ("Folder Id is" + $Folder.Id)      
        $endpoint = ""
        $url = "https://cloud.uipath.com/$accountLogicalName/$tenantName/orchestrator_/odata/$endpoint"
        #Write-Output $url
        #Write-Output $Folderheaders
        #Write-Output $Folder.FullyQualifiedName
        $ProcessesCount = Get-FolderLevelCount "Releases"
        #Write-Output ("Processes Count is " + $ProcessesCount)
        $AssetCount = Get-FolderLevelCount "Assets"
        #Write-Output ("Asset Count is " + $AssetCount)
         $QueueDefinitionsCount = Get-FolderLevelCount "QueueDefinitions"
        #Write-Output ("Queues Count is " + $QueueDefinitionsCount)
         $ProcessSchedules = Get-FolderLevelCount "ProcessSchedules"
        #Write-Output ("Triggers Count is " + $ProcessSchedules)
        $Machines = Get-FolderLevelCount ("Machines/UiPath.Server.Configuration.OData.GetAssignedMachines(folderId=" + $Folder.Id + ")")
        #Write-Output ("Machines Count under this folder is " + $Machines)
        $Users = Get-FolderLevelCount ("Folders/UiPath.Server.Configuration.OData.GetUsersForFolder(key=" + $Folder.Id + ",includeInherited=true)?includeAlertsEnabled=false")
        #Write-Output ("Users Count under this folder is " + $Users)   
        

        $Folderresult = [PSCustomObject]@{
        FolderName = $Folder.FullyQualifiedName
        Processes = $ProcessesCount
        Assets   = $AssetCount
        Queues = $QueueDefinitionsCount
        Triggers = $ProcessSchedules
        Machines = $Machines
        Users    = $Users
        #Queues   = Get-Count "QueueDefinitions"
        #Assets   = Get-Count "Assets"
        }
        
        $Folderresults += $Folderresult     
        Write-Output ("extracted details for: " + $Folder.FullyQualifiedName)   
        
    }
    $Folderresults | Format-Table -AutoSize
    # === Export Report ===
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
#$csvPath = "OrchestratorTenantReport_$timestamp.csv"
#Write-Output "Exporting report to $csvPath..."
#$results | Export-Csv -Path $csvPath -NoTypeInformation



# Optional: Excel export using ImportExcel module
Install-Module ImportExcel -Scope CurrentUser
# $results | Export-Excel -Path "OrchestratorTenantReport_$timestamp.xlsx" -AutoSize -BoldTopRow


$excelPath = "TenantValidationReport_$tenantName.xlsx"
Write-Output "Exporting Excel report to $excelPath..."

$results | Export-Excel -Path $excelPath -WorksheetName "Cloud Tenant Summary" -AutoSize -BoldTopRow
$Folderresults | Export-Excel -Path $excelPath -WorksheetName "Cloud Folder Details" -AutoSize -BoldTopRow

Write-Output "Excel report exported successfully to $excelPath."


Write-Output "Script completed."

#--------------------------------SUITE REPORT-----------------------------------------
 

#--------------------------------SUITE REPORT-----------------------------------------

# === Suite Authentication ===
$SuiteTokenHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$SuiteTokenHeaders.Add("Content-Type", "application/json")

$SuiteAuthBody = @"
{
  `"tenancyName`": `"<TenantName>`",
  `"usernameOrEmailAddress`": `"apiuser`",
  `"password`": `"<YourPWd>`"
}
"@

$SuiteAuthResponse = Invoke-RestMethod 'https://<AS_SourceFQDN>/default/$var_SuiteTenantName/orchestrator_/api/Account/Authenticate' -Method 'POST' -Headers $SuiteTokenHeaders -Body $SuiteAuthBody
$SuiteToken = $SuiteAuthResponse.result

# === Suite Headers and Base URL ===
$SuiteHeaders = @{ Authorization = "Bearer $SuiteToken" }
$SuiteHeaders["X-UIPATH-TenantName"] = $var_SuiteTenantName
$SuiteAccountLogicalName = "default"
$SuiteTenantName = $var_SuiteTenantName
$SuiteLogicalTenant = "17531af8-128b-415a-9bad-c56bd4b211bc"
$SuiteBaseUrl = "https://<AS_SourceFQDN>/$SuiteAccountLogicalName/$SuiteLogicalTenant/odata"
$SuiteFolderHeaders = @{ Authorization = "Bearer $SuiteToken" }

function Get-SuiteCount($endpoint) {
    $url = "$SuiteBaseUrl/$endpoint`?\$count=true&\$top=1"
    try {
        $resp = Invoke-RestMethod -Uri $url -Headers $SuiteHeaders
        return $resp.'@odata.count'
    } catch {
        Write-Warning "Failed to fetch $endpoint for $SuiteTenantName. $_"
        return -1
    }
}

$SuiteResults = @()
$SuiteFolderResults = @()

$SuiteSummary = [PSCustomObject]@{
    Tenant   = $SuiteTenantName
    Machines = Get-SuiteCount "Machines"
    Users    = Get-SuiteCount "Users"
    Folders  = Get-SuiteCount "Folders"
    Packages = Get-SuiteCount "Processes"
    Roles    = Get-SuiteCount "Roles"
}
$SuiteResults += $SuiteSummary
$SuiteResults | Format-Table -AutoSize

# === Get Suite Folders ===
$SuiteFoldersUrl = "$SuiteBaseUrl/Folders"
$SuiteFoldersResponse = Invoke-RestMethod -Uri $SuiteFoldersUrl -Headers $SuiteHeaders
$SuiteFolders = $SuiteFoldersResponse.value
Write-Output ("************************************************")
Write-Output ("Total Folder Count on this tenant is:" + $SuiteFoldersResponse.'@odata.count')
Write-Output ("************************************************")

function Get-SuiteFolderLevelCount($endpoint) {
    $url = "https://<AS_SourceFQDN>/$SuiteAccountLogicalName/$SuiteTenantName/orchestrator_/odata/$endpoint"
    try {
        $resp = Invoke-RestMethod -Uri $url -Headers $SuiteFolderHeaders
        return $resp.'@odata.count'
    } catch {
        Write-Warning $_.Exception.Message
        return -1
    }
}

foreach ($SuiteFolder in $SuiteFolders) {
    $SuiteFolderHeaders["X-UIPATH-OrganizationUnitId"] = $SuiteFolder.Id

    $SuiteProcesses = Get-SuiteFolderLevelCount "Releases"
    $SuiteAssets = Get-SuiteFolderLevelCount "Assets"
    $SuiteQueues = Get-SuiteFolderLevelCount "QueueDefinitions"
    $SuiteTriggers = Get-SuiteFolderLevelCount "ProcessSchedules"
    $SuiteMachines = Get-SuiteFolderLevelCount ("Machines/UiPath.Server.Configuration.OData.GetAssignedMachines(folderId=" + $SuiteFolder.Id + ")")
    $SuiteUsers = Get-SuiteFolderLevelCount ("Folders/UiPath.Server.Configuration.OData.GetUsersForFolder(key=" + $SuiteFolder.Id + ",includeInherited=true)?includeAlertsEnabled=false")

    $SuiteFolderResults += [PSCustomObject]@{
        FolderName = $SuiteFolder.FullyQualifiedName
        Processes  = $SuiteProcesses
        Assets     = $SuiteAssets
        Queues     = $SuiteQueues
        Triggers   = $SuiteTriggers
        Machines   = $SuiteMachines
        Users      = $SuiteUsers
    }

    Write-Output ("Extracted details for: " + $SuiteFolder.FullyQualifiedName)
}

$SuiteFolderResults | Format-Table -AutoSize

# === Export Suite Report ===
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$SuiteExcelPath = "$excelPath"
Write-Output "Exporting Excel report to $SuiteExcelPath..."

$SuiteResults | Export-Excel -Path $SuiteExcelPath -WorksheetName "Suite Tenant Summary" -AutoSize -BoldTopRow
$SuiteFolderResults | Export-Excel -Path $SuiteExcelPath -WorksheetName "Suite Folder Details" -AutoSize -BoldTopRow

Write-Output "Excel report exported successfully to $SuiteExcelPath."
