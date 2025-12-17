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
$response | ConvertTo-Json
$token = $response.result

$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Content-Type", "application/x-www-form-urlencoded")

#$body = "grant_type=client_credentials&client_id=7b4ab67b-dc11-418c-af0a-2b0953dd599f&client_secret=2!EOECLW)BHlXWa9&scope=OR.Default"

#$response = Invoke-RestMethod 'https://<yourASFQDN>/identity_/connect/token' -Method 'POST' -Headers $headers -Body $body
#$response | ConvertTo-Json
#Write-Output "Access Token: $($response.access_token)"
#$token = $response.access_token
#Write-Output "Authentication successful."

# === Initialize Results Array ===
$results = @()
$Folderresults = @()

$accountLogicalName = "default"#cloudorg
#$defaultTenant = "DefaultTenant"  # Any valid tenant for auth

$tenantName = "<yourTenant>"#Tenantname
$logicalTenant = "17531af8-128b-415a-9bad-c56bd4b211bc"#admin page navigation from url
$headers = @{ Authorization = "Bearer $token" }
$baseUrl = "https://<yourASFQDN>/$accountLogicalName/$logicalTenant/odata"
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
        $url = "https://<yourASFQDN>/$accountLogicalName/$tenantName/orchestrator_/odata/$endpoint"
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
            Write-Warning $ErrorMessage
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
        #$url = "https://cloud.uipath.com/$accountLogicalName/$tenantName/orchestrator_/odata/$endpoint"
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


$excelPath = "OrchestratorTenantReport_$timestamp.xlsx"
Write-Output "Exporting Excel report to $excelPath..."

$results | Export-Excel -Path $excelPath -WorksheetName "Tenant Summary" -AutoSize -BoldTopRow
$Folderresults | Export-Excel -Path $excelPath -WorksheetName "Folder Details" -AutoSize -BoldTopRow

Write-Output "Excel report exported successfully to $excelPath."


Write-Output "Script completed."
 


   
