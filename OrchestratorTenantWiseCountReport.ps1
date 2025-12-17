$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Content-Type", "application/x-www-form-urlencoded")

$body = "grant_type=client_credentials&client_id=6759bc60-b038-4ca1-8ab2-266906a74f1f&client_secret=%3FVdBAR%243a)0C%40M7~&scope=OR.Default"

$response = Invoke-RestMethod 'https://cloud.uipath.com/identity_/connect/token' -Method 'POST' -Headers $headers -Body $body
#$response | ConvertTo-Json
#Write-Output "Access Token: $($response.access_token)"
$token = $response.access_token
#Write-Output "Authentication successful."

# === Initialize Results Array ===
$results = @()
$Folderresults = @()

$accountLogicalName = "uipatledsyvu"
#$defaultTenant = "DefaultTenant"  # Any valid tenant for auth

$tenantName = "DefaultTenant"
$logicalTenant = "ec1ab793-146f-41dc-b28d-24c6493233b2"
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


$excelPath = "OrchestratorTenantReport_$timestamp.xlsx"
Write-Output "Exporting Excel report to $excelPath..."

$results | Export-Excel -Path $excelPath -WorksheetName "Tenant Summary" -AutoSize -BoldTopRow
$Folderresults | Export-Excel -Path $excelPath -WorksheetName "Folder Details" -AutoSize -BoldTopRow

Write-Output "Excel report exported successfully to $excelPath."


Write-Output "Script completed."
 


   
