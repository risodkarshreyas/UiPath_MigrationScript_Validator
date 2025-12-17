# Install ImportExcel if not already installed
#Install-Module ImportExcel -Scope CurrentUser -Force


#Suite Auth details
$SuiteOrgName = "default"
$SuiteTenantName = "<yourTenantName>"
$SuiteTenantID = "2fb2a7b3-848c-4433-ba45-13c5a96d5d04" #Tenant Logical ID
$SuiteAuthTokenURL = "https://<yourASFQDN>/$SuiteOrgName/$SuiteTenantName/orchestrator_/api/Account/Authenticate"
$SuiteBaseURL = "https://<yourASFQDN>/$SuiteOrgName/$SuiteTenantID/odata"
$SuiteAPIUserName = "apiuser"
$SuiteAPIUserPassword = "<yourPassword>"

#Cloud Auth Details
$CloudOrgName = "micronorg"
$CloudTenantName = "<yourTenantName>"
$CloudTenantID = "b0c4f1c2-59b4-41a8-988f-85477c621c5e"   #Tenant Logical ID
$CloudAuthTokenURL = "https://cloud.uipath.com/$CloudOrgName/identity_/connect/token"
$CloudBaseURL = "https://cloud.uipath.com/$CloudOrgName/$CloudTenantName/odata"
$CloudClientID = "9fe53278-7de5-4c89-a9fb-3a2ef18cda92"
$CloudClientSecret = "zDoybr0S0N6~7AWp"

#OutputFile Details
$OutputFolderPath = "C:\Users\<Username>\Desktop\Reports"

$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$outputFilePath = ($OutputFolderPath + "\MigrationValidation_" + $CloudTenantName + "_" + $timestamp +".xlsx")

$Headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"

$AllFolders_InTenant = @()
$AllUsers_InTenant = @()
$AllRoles_InTenant = @()
$AllMachines_InTenant = @()
$AllPackages_InTenant = @()

$AllProcesses_InFolder = @()
$AllTriggers_InFolder = @()
$AllAssets_InFolder = @()
$AllQueues_InFolder = @()
$AllUsers_InFolder = @()
$AllMachines_InFolder = @()

$TenantDetails = [PSCustomObject]@{
    Name   = ""
    FoldersCount = -1
    UsersCount = -1
    RolesCount = -1
    MachinesCount = -1
    PackagesCount = -1
}

$FolderwiseDetails = [PSCustomObject]@()


function GetAccessToken_Cloud() {
    #CloudAuthentication
    $CloudAuthHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $CloudAuthHeaders.Add("Content-Type", "application/x-www-form-urlencoded")
    $body = ("grant_type=client_credentials&client_id=" + $CloudClientID + "&client_secret=" + $CloudClientSecret + "&scope=OR.Default")
    try{
        $response = Invoke-RestMethod $CloudAuthTokenURL -Method 'POST' -Headers $CloudAuthHeaders -Body $body

        #Write-Host ("Cloud Auth Token : " + $response.access_token)
        Write-Host "Cloud Authentication successful."
        
        return $response.access_token

    } catch {
        Write-Warning "Failed to Authenticate to Cloud. Exception.Message: $($_.Exception.Message)"
        return -1
    }





}

function GetAccessToken_Suite() {
    ## These headers are specific to authentication for Automation Suite only
    $SuiteAuthHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $SuiteAuthHeaders.Add("Content-Type", "application/json")

    $body = 
@"
    {
        "tenancyName": "$SuiteTenantName",
        "usernameOrEmailAddress": "$SuiteAPIUserName",
        "password": "$SuiteAPIUserPassword"
    }
"@

    try{
        $response = Invoke-RestMethod $SuiteAuthTokenURL -Method 'POST' -Headers $SuiteAuthHeaders -Body $body

        #Write-Host ("Suite Auth Token : " + $response.result)
        Write-Host "Suite Authentication successful."
        
        return $response.result

    } catch {
        Write-Warning "Failed to Authenticate to Suite. Exception.Message: $($_.Exception.Message)"
        return -1
    }

    #Write-Host ("Suite Auth Token : " + $token)
    Write-Host "Suite Authentication successful."
}

function GetAllFolders($environment) {

    switch ($environment.ToUpper()) {
        CLOUD { $baseUrl = $CloudBaseURL }

        SUITE { $baseUrl = $SuiteBaseURL } 

        default { 
            Write-Error "Wrong Environment value. Should be either CLOUD or SUITE"
            return null
        }
    }

    try {
        $response = Invoke-RestMethod -Uri "$baseUrl/Folders" -Headers $Headers

        $Folders = @()
        Foreach ($Folder in $response.value){
            $Folders += [PSCustomObject]@{
                Name = $Folder.FullyQualifiedName
                Id = $Folder.id
            }
        }

        Write-Host "----- Folder details retrieved successfully from tenant level on $environment"

        return $Folders
    } catch {
        Write-Warning "Failed to fetch Folders for $environment. Exception.Message: $($_.Exception.Message)"
        return null
    }

}

function GetAllUsersInTenant($environment) {

    switch ($environment.ToUpper()) {
        CLOUD { $baseUrl = $CloudBaseURL }

        SUITE { $baseUrl = $SuiteBaseURL } 

        default { 
            Write-Error "Wrong Environment value. Should be either CLOUD or SUITE"
            return null
        }
    }

    try {
        $response = Invoke-RestMethod -Uri "$baseUrl/Users" -Headers $Headers

        $Users = @()
        Foreach ($User in $response.value){
            $Users += [PSCustomObject]@{
                Name = $User.FullName
                Type = $User.Type
                Roles = ($User.RolesList -join ', ')
                
            }
        }

        Write-Host "----- User details retrieved successfully from tenant level on $environment"

        return $Users
    } catch {
        Write-Warning "Failed to fetch Users from tenant level on $environment. Exception.Message: $($_.Exception.Message)"
        return null
    }

}

function GetAllRolesInTenant($environment) {

    switch ($environment.ToUpper()) {
        CLOUD { $baseUrl = $CloudBaseURL }

        SUITE { $baseUrl = $SuiteBaseURL } 

        default { 
            Write-Error "Wrong Environment value. Should be either CLOUD or SUITE"
            return null
        }
    }

    try {
        $response = Invoke-RestMethod -Uri "$baseUrl/Roles" -Headers $Headers

        $Roles = @()
        Foreach ($Role in $response.value){
            $Roles += [PSCustomObject]@{
                Name = $Role.DisplayName
                Type = $Role.Type
            }
        }


        Write-Host "----- Roles details retrieved successfully from tenant level on $environment"

        return $Roles
    } catch {
        Write-Warning "Failed to fetch Roles from tenant level on $environment. Exception.Message: $($_.Exception.Message)"
        return null
    }

}

function GetAllMachinesInTenant($environment) {

    switch ($environment.ToUpper()) {
        CLOUD { $baseUrl = $CloudBaseURL }

        SUITE { $baseUrl = $SuiteBaseURL } 

        default { 
            Write-Error "Wrong Environment value. Should be either CLOUD or SUITE"
            return null
        }
    }

    try {
        $response = Invoke-RestMethod -Uri "$baseUrl/Machines" -Headers $Headers

        $Machines = @()
        Foreach ($Machine in $response.value){
            $Machines += [PSCustomObject]@{
                Name = $Machine.Name
                Type = $Machine.Type
                Scope = $Machine.Scope
                NonProdRuntimes = $Machine.NonProductionSlots
                UnattendedRuntimes = $Machine.UnattendedSlots
                TestingRuntimes = $Machine.TestAutomationSlots
            }
        }

        Write-Host "----- Machines details retrieved successfully from tenant level on $environment"

        return $Machines
    } catch {
        Write-Warning "Failed to fetch Machines from tenant level on $environment. Exception.Message: $($_.Exception.Message)"
        return null
    }

}

function GetAllPackagesInTenant($environment) {

    switch ($environment.ToUpper()) {
        CLOUD { $baseUrl = $CloudBaseURL }

        SUITE { $baseUrl = $SuiteBaseURL } 

        default { 
            Write-Error "Wrong Environment value. Should be either CLOUD or SUITE"
            return null
        }
    }

    try {
        $response = Invoke-RestMethod -Uri "$baseUrl/Processes" -Headers $Headers

        $Packages = @()
        Foreach ($Package in $response.value){
            $Packages += [PSCustomObject]@{
                Name = $Package.Title
                Version = $Package.Version
            }
        }

        Write-Host "----- Packages details retrieved successfully from tenant level on $environment"

        return $Packages
    } catch {
        Write-Warning "Failed to fetch Packages from tenant level on $environment. Exception.Message: $($_.Exception.Message)"
        return null
    }

}

function GetAllProcessesInFolder($environment, $folderName) {

    switch ($environment.ToUpper()) {
        CLOUD { $baseUrl = $CloudBaseURL }

        SUITE { $baseUrl = $SuiteBaseURL } 

        default { 
            Write-Error "Wrong Environment value. Should be either CLOUD or SUITE"
            return null
        }
    }

    try {
        $response = Invoke-RestMethod -Uri "$baseUrl/Releases" -Headers $Headers

        $Processes = @()
        Foreach ($Process in $response.value){
            $Processes += [PSCustomObject]@{
                FolderName = $folderName
                Name = $Process.Name
                InputArguments = $Process.InputArguments
                OuputArguments = $Process.Arguments.Output
                PackageVersion = $Process.ProcessVersion
                EntryPointId = $Process.EntryPointId
            }
        }

        Write-Host "----- Processes details retrieved successfully from $folderName Folder on $environment"

        return $Processes
    } catch {
        Write-Warning "Failed to fetch Processes from $folderName folder on $environment. Exception.Message: $($_.Exception.Message)"
        return null
    }

}

function GetAllTriggersInFolder($environment, $folderName) {

    switch ($environment.ToUpper()) {
        CLOUD { $baseUrl = $CloudBaseURL }

        SUITE { $baseUrl = $SuiteBaseURL } 

        default { 
            Write-Error "Wrong Environment value. Should be either CLOUD or SUITE"
            return null
        }
    }

    try {
        $response = Invoke-RestMethod -Uri "$baseUrl/ProcessSchedules" -Headers $Headers

        $Triggers = @()
        Foreach ($Trigger in $response.value){
            $Triggers += [PSCustomObject]@{
                FolderName = $folderName
                Name = $Trigger.Name
                Enabled = $Trigger.Enabled
                JobPriority = $Trigger.JobPriority
                ProcessName = $Trigger.PackageName
                InputArguments = $Trigger.InputArguments
                QueueName = $Trigger.QueueDefinitionName
            }
        }

        Write-Host "----- Triggers details retrieved successfully from $folderName Folder on $environment"

        return $Triggers
    } catch {
        Write-Warning "Failed to fetch Triggers from $folderName folder on $environment. Exception.Message: $($_.Exception.Message)"
        return null
    }

}

function GetAllAssetsInFolder($environment, $folderName) {

    switch ($environment.ToUpper()) {
        CLOUD { $baseUrl = $CloudBaseURL }

        SUITE { $baseUrl = $SuiteBaseURL } 

        default { 
            Write-Error "Wrong Environment value. Should be either CLOUD or SUITE"
            return null
        }
    }

    try {
        $response = Invoke-RestMethod -Uri "$baseUrl/Assets" -Headers $Headers

        $Assets = @()
        Foreach ($Asset in $response.value){

            $valueCode = switch ($Asset.ValueType) {
                "Credential" { "username: $($Asset.CredentialUsername)" }
                default { $Asset.Value }
            }

            $Assets += [PSCustomObject]@{
                FolderName = $folderName
                Name = $Asset.Name
                Description = $Asset.Description
                Scope      = $Asset.ValueScope
                Type       = $Asset.ValueType
                ValueCode  = $valueCode
                LinkedFolders = $Asset.FoldersCount
            }
        }

        Write-Host "----- Assets details retrieved successfully from $folderName Folder on $environment"

        return $Assets
    } catch {
        Write-Warning "Failed to fetch Assets from $folderName folder on $environment. Exception.Message: $($_.Exception.Message)"
        return null
    }

}

function GetAllQueuesInFolder($environment, $folderName) {

    switch ($environment.ToUpper()) {
        CLOUD { $baseUrl = $CloudBaseURL }

        SUITE { $baseUrl = $SuiteBaseURL } 

        default { 
            Write-Error "Wrong Environment value. Should be either CLOUD or SUITE"
            return null
        }
    }

    try {
        $response = Invoke-RestMethod -Uri "$baseUrl/QueueDefinitions" -Headers $Headers

        $Queues = @()
        Foreach ($Queue in $response.value){

            $Queues += [PSCustomObject]@{
                FolderName = $folderName
                Name = $Queue.Name
                MaxRetries = $Queue.MaxNumberOfRetries
                RetryFailedItems = $Queue.AcceptAutomaticallyRetry
                RetryAbandonedItems = $Queue.RetryAbandonedItems
                UniqueReference = $Queue.EnforceUniqueReference
                LinkedFolders = $Queue.FoldersCount
            }
        }

        Write-Host "----- Queues details retrieved successfully from $folderName Folder on $environment"

        return $Queues
    } catch {
        Write-Warning "Failed to fetch Queues from $folderName folder on $environment. Exception.Message: $($_.Exception.Message)"
        return null
    }

}

function GetAllUsersInFolder($environment, $folderName, $folderID) {

    switch ($environment.ToUpper()) {
        CLOUD { $baseUrl = $CloudBaseURL }

        SUITE { $baseUrl = $SuiteBaseURL } 

        default { 
            Write-Error "Wrong Environment value. Should be either CLOUD or SUITE"
            return null
        }
    }

    try {
    $uri = ("$baseUrl/Folders/UiPath.Server.Configuration.OData.GetUsersForFolder(key=" + $folderID + ",includeInherited=true)?includeAlertsEnabled=false")
        $response = Invoke-RestMethod -Uri $uri -Headers $Headers

        $Users = @()
        Foreach ($User in $response.value){

            $Users += [PSCustomObject]@{
                FolderName = $folderName
                Name = $User.UserEntity.UserName
                Type = $User.UserEntity.Type
                Inherited = $User.UserEntity.IsInherited
                Attended = $User.UserEntity.MayHaveAttended
                Unattended = $User.UserEntity.MayHaveUnattended
                Roles = (($User.Roles | ForEach-Object { $_.Name }) -join ', ')
                Alerts = $User.HasAlertsEnabled


            }
        }
        
        Write-Host "----- Users details retrieved successfully from $folderName Folder on $environment"

        return $Users
    } catch {
        Write-Warning "Failed to fetch Users from $folderName folder on $environment. Exception.Message: $($_.Exception.Message)"
        return null
    }

}

function GetAllMachinesInFolder($environment, $folderName, $folderID) {

    switch ($environment.ToUpper()) {
        CLOUD { $baseUrl = $CloudBaseURL }

        SUITE { $baseUrl = $SuiteBaseURL } 

        default { 
            Write-Error "Wrong Environment value. Should be either CLOUD or SUITE"
            return null
        }
    }

    try {
    $uri = ("$baseUrl/Machines/UiPath.Server.Configuration.OData.GetAssignedMachines(folderId=" + $folderID + ")")
        $response = Invoke-RestMethod -Uri $uri -Headers $Headers

        $Machines = @()
        Foreach ($Machine in $response.value){

            $Machines += [PSCustomObject]@{
                FolderName = $folderName
                Name = $Machine.Name
                Type = $Machine.Type
                Scope = $Machine.Scope
                NonProdRuntimes = $Machine.NonProductionSlots
                UnattendedRuntimes = $Machine.UnattendedSlots
                TestingRuntimes = $Machine.TestAutomationSlots
                AutomationType = $Machine.AutomationType

            }
        }
        
        Write-Host "----- Machines details retrieved successfully from $folderName Folder on $environment"

        return $Machines
    } catch {
        Write-Warning "Failed to fetch Machines from $folderName folder on $environment. Exception.Message: $($_.Exception.Message)"
        return null
    }

}


#================== EXECUTION FLOW  =====================
Write-Host "############ Starting execution flow for CLOUD ############"

#Authenticate
$token = GetAccessToken_Cloud
$Headers = @{ Authorization = ("Bearer " + $token) }
$Headers.Add("Content-Type", "application/x-www-form-urlencoded")
$Headers["X-UIPATH-TenantName"] = $CloudTenantName

#Get Tenant Level Data
Write-Host "### Fetching Tenant Level Details...."
$TenantDetails.Name = $CloudTenantName

$AllFolders_InTenant = GetAllFolders("Cloud")
$AllUsers_InTenant = GetAllUsersInTenant("Cloud")
$AllRoles_InTenant = GetAllRolesInTenant("Cloud")
$AllMachines_InTenant = GetAllMachinesInTenant("Cloud")
$AllPackages_InTenant = GetAllPackagesInTenant("Cloud")

$TenantDetails.FoldersCount = @($AllFolders_InTenant).Count
$TenantDetails.UsersCount = @($AllUsers_InTenant).Count
$TenantDetails.RolesCount = @($AllRoles_InTenant).Count
$TenantDetails.MachinesCount = @($AllMachines_InTenant).Count
$TenantDetails.PackagesCount = @($AllPackages_InTenant).Count

#Get Folder-wise Data
foreach($Folder in $AllFolders_InTenant){
    Write-Host ("`n### Fetching Folder level details for " + $Folder.Name + " folder....")
    $Headers["X-UIPATH-OrganizationUnitId"] = $Folder.Id

    $currentFolder_Processes = GetAllProcessesInFolder "Cloud" $Folder.Name
    $currentFolder_Triggers = GetAllTriggersInFolder "Cloud" $Folder.Name
    $currentFolder_Assets = GetAllAssetsInFolder "Cloud" $Folder.Name
    $currentFolder_Queues = GetAllQueuesInFolder "Cloud" $Folder.Name
    $currentFolder_Users = GetAllUsersInFolder "Cloud" $Folder.Name $Folder.Id
    $currentFolder_Machines = GetAllMachinesInFolder "Cloud" $Folder.Name $Folder.Id

    $AllProcesses_InFolder += $currentFolder_Processes
    $AllTriggers_InFolder += $currentFolder_Triggers
    $AllAssets_InFolder += $currentFolder_Assets
    $AllQueues_InFolder += $currentFolder_Queues
    $AllUsers_InFolder += $currentFolder_Users
    $AllMachines_InFolder += $currentFolder_Machines

    $FolderwiseDetails += [PSCustomObject]@{
        Name = $Folder.Name
        ProcessesCount = @($currentFolder_Processes).Count
        TriggersCount = @($currentFolder_Triggers).Count
        AssetsCount = @($currentFolder_Assets).Count
        QueuesCount = @($currentFolder_Queues).Count
        UsersCount = @($currentFolder_Users).Count
        MachinesCount = @($currentFolder_Machines).Count
    }
}

# === Export Cloud data to Excel ===

$TenantDetails | Export-Excel -Path $outputFilePath -WorksheetName "CloudTenantLevelCounts" -AutoSize -BoldTopRow
$FolderwiseDetails | Export-Excel -Path $outputFilePath -WorksheetName "CloudFolderwiseCounts" -AutoSize -BoldTopRow

$AllFolders_InTenant | Export-Excel -Path $outputFilePath -WorksheetName "Cloud_Folders" -AutoSize -BoldTopRow
$AllUsers_InTenant | Export-Excel -Path $outputFilePath -WorksheetName "Cloud_AllUsers" -AutoSize -BoldTopRow
$AllRoles_InTenant | Export-Excel -Path $outputFilePath -WorksheetName "Cloud_Roles" -AutoSize -BoldTopRow
$AllMachines_InTenant | Export-Excel -Path $outputFilePath -WorksheetName "Cloud_AllMachines" -AutoSize -BoldTopRow
$AllPackages_InTenant | Export-Excel -Path $outputFilePath -WorksheetName "Cloud_Packages" -AutoSize -BoldTopRow

$AllProcesses_InFolder | Export-Excel -Path $outputFilePath -WorksheetName "Cloud_Processes" -AutoSize -BoldTopRow
$AllTriggers_InFolder | Export-Excel -Path $outputFilePath -WorksheetName "Cloud_Triggers" -AutoSize -BoldTopRow
$AllAssets_InFolder | Export-Excel -Path $outputFilePath -WorksheetName "Cloud_Assets" -AutoSize -BoldTopRow
$AllQueues_InFolder | Export-Excel -Path $outputFilePath -WorksheetName "Cloud_Queues" -AutoSize -BoldTopRow
$AllUsers_InFolder | Export-Excel -Path $outputFilePath -WorksheetName "Cloud_UsersInFolder" -AutoSize -BoldTopRow
$AllMachines_InFolder | Export-Excel -Path $outputFilePath -WorksheetName "Cloud_MachinesInFolder" -AutoSize -BoldTopRow

Write-Host "############################## CLOUD TENANT DETAILS EXPORTED TO EXCEL"



Write-Host "############ Starting execution flow for SUITE ############"

#ResetVariables
$Headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"

$AllFolders_InTenant = @()
$AllUsers_InTenant = @()
$AllRoles_InTenant = @()
$AllMachines_InTenant = @()
$AllPackages_InTenant = @()

$AllProcesses_InFolder = @()
$AllTriggers_InFolder = @()
$AllAssets_InFolder = @()
$AllQueues_InFolder = @()
$AllUsers_InFolder = @()
$AllMachines_InFolder = @()

$TenantDetails = [PSCustomObject]@{
    Name   = ""
    FoldersCount = -1
    UsersCount = -1
    RolesCount = -1
    MachinesCount = -1
    PackagesCount = -1
}

$FolderwiseDetails = [PSCustomObject]@()

#Authenticate
$token = GetAccessToken_Suite #(Pending)TO BE TESTED
$Headers = @{ Authorization = ("Bearer " + $token) }
$Headers.Add("Content-Type", "application/x-www-form-urlencoded")
$Headers["X-UIPATH-TenantName"] = $SuiteTenantName

#Get Tenant Level Data
Write-Host "### Fetching Tenant Level Details...."
$TenantDetails.Name = $SuiteTenantName

$AllFolders_InTenant = GetAllFolders("Suite")
$AllUsers_InTenant = GetAllUsersInTenant("Suite")
$AllRoles_InTenant = GetAllRolesInTenant("Suite")
$AllMachines_InTenant = GetAllMachinesInTenant("Suite")
$AllPackages_InTenant = GetAllPackagesInTenant("Suite")

$TenantDetails.FoldersCount = @($AllFolders_InTenant).Count
$TenantDetails.UsersCount = @($AllUsers_InTenant).Count
$TenantDetails.RolesCount = @($AllRoles_InTenant).Count
$TenantDetails.MachinesCount = @($AllMachines_InTenant).Count
$TenantDetails.PackagesCount = @($AllPackages_InTenant).Count

#Get Folder-wise Data
foreach($Folder in $AllFolders_InTenant){
    Write-Host ("`n### Fetching Folder level details for " + $Folder.Name + " folder....")
    $Headers["X-UIPATH-OrganizationUnitId"] = $Folder.Id

    $currentFolder_Processes = GetAllProcessesInFolder "Suite" $Folder.Name
    $currentFolder_Triggers = GetAllTriggersInFolder "Suite" $Folder.Name
    $currentFolder_Assets = GetAllAssetsInFolder "Suite" $Folder.Name
    $currentFolder_Queues = GetAllQueuesInFolder "Suite" $Folder.Name
    $currentFolder_Users = GetAllUsersInFolder "Suite" $Folder.Name $Folder.Id
    $currentFolder_Machines = GetAllMachinesInFolder "Suite" $Folder.Name $Folder.Id

    $AllProcesses_InFolder += $currentFolder_Processes
    $AllTriggers_InFolder += $currentFolder_Triggers
    $AllAssets_InFolder += $currentFolder_Assets
    $AllQueues_InFolder += $currentFolder_Queues
    $AllUsers_InFolder += $currentFolder_Users
    $AllMachines_InFolder += $currentFolder_Machines

    $FolderwiseDetails += [PSCustomObject]@{
        Name = $Folder.Name
        ProcessesCount = @($currentFolder_Processes).Count
        TriggersCount = @($currentFolder_Triggers).Count
        AssetsCount = @($currentFolder_Assets).Count
        QueuesCount = @($currentFolder_Queues).Count
        UsersCount = @($currentFolder_Users).Count
        MachinesCount = @($currentFolder_Machines).Count
    }
}

# === Export Suite data to Excel ===

$TenantDetails | Export-Excel -Path $outputFilePath -WorksheetName "SuiteTenantLevelCounts" -AutoSize -BoldTopRow
$FolderwiseDetails | Export-Excel -Path $outputFilePath -WorksheetName "SuiteFolderwiseCounts" -AutoSize -BoldTopRow

$AllFolders_InTenant | Export-Excel -Path $outputFilePath -WorksheetName "Suite_Folders" -AutoSize -BoldTopRow
$AllUsers_InTenant | Export-Excel -Path $outputFilePath -WorksheetName "Suite_AllUsers" -AutoSize -BoldTopRow
$AllRoles_InTenant | Export-Excel -Path $outputFilePath -WorksheetName "Suite_Roles" -AutoSize -BoldTopRow
$AllMachines_InTenant | Export-Excel -Path $outputFilePath -WorksheetName "Suite_AllMachines" -AutoSize -BoldTopRow
$AllPackages_InTenant | Export-Excel -Path $outputFilePath -WorksheetName "Suite_Packages" -AutoSize -BoldTopRow

$AllProcesses_InFolder | Export-Excel -Path $outputFilePath -WorksheetName "Suite_Processes" -AutoSize -BoldTopRow
$AllTriggers_InFolder | Export-Excel -Path $outputFilePath -WorksheetName "Suite_Triggers" -AutoSize -BoldTopRow
$AllAssets_InFolder | Export-Excel -Path $outputFilePath -WorksheetName "Suite_Assets" -AutoSize -BoldTopRow
$AllQueues_InFolder | Export-Excel -Path $outputFilePath -WorksheetName "Suite_Queues" -AutoSize -BoldTopRow
$AllUsers_InFolder | Export-Excel -Path $outputFilePath -WorksheetName "Suite_UsersInFolder" -AutoSize -BoldTopRow
$AllMachines_InFolder | Export-Excel -Path $outputFilePath -WorksheetName "Suite_MachinesInFolder" -AutoSize -BoldTopRow

Write-Host "############################## SUITE TENANT DETAILS EXPORTED TO EXCEL"

