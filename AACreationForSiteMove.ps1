# 1. Obtain a service number.
# 2. Obtain a free Phone System - Virtual User license or a paid Phone System license to use with the resource account or a Phone System license.
# 3. Create the resource account. An auto attendant or call queue is required to have an associated resource account.
# 4. Assign the Phone System or a Phone System - Virtual user license to the resource account.
# 5. Assign a service phone number to the resource account you just assigned licenses to.
# 6. Create a Phone System call queue or auto attendant
# 7. Link the resource account with a call queue or auto attendant.

#------------------------------------------------------------------------------------------
# Connections Functions
#------------------------------------------------------------------------------------------
function global:Connect-AzureActiveDirectory {
    If ( !(Get-Module -Name AzureAD)) {Import-Module -Name AzureAD -ErrorAction SilentlyContinue}
    If ( !(Get-Module -Name AzureADPreview)) {Import-Module -Name AzureADPreview -ErrorAction SilentlyContinue}
    If ( (Get-Module -Name AzureAD) -or (Get-Module -Name AzureADPreview)) {
        If ( !($global:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
        If ( $global:myOffice365Services['Office365CredentialsMFA']) {
            Write-Host 'Connecting to Azure Active Directory with Modern Authentication ..'
            $Parms = @{'AzureEnvironment' = $global:myOffice365Services['AzureEnvironment']}
        }
        Else {
            Write-Host "Connecting to Azure Active Directory using $($global:myOffice365Services['Office365Credentials'].username) .."
            $Parms = @{'Credential' = $global:myOffice365Services['Office365Credentials']; 'AzureEnvironment' = $global:myOffice365Services['AzureEnvironment']}
        }
        Connect-AzureAD @Parms
    }
    Else {
        If ( !(Get-Module -Name MSOnline)) {Import-Module -Name MSOnline -ErrorAction SilentlyContinue}
        If ( Get-Module -Name MSOnline) {
            If ( !($global:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
            Write-Host "Connecting to Azure Active Directory using $($global:myOffice365Services['Office365Credentials'].username) .."
            Connect-MsolService -Credential $global:myOffice365Services['Office365Credentials'] -AzureEnvironment $global:myOffice365Services['AzureEnvironment']
        }
        Else {Write-Error -Message 'Cannot connect to Azure Active Directory - problem loading module.'}
    }
}
function global:Connect-SkypeOnline {
    If ( !(Get-Module -Name SkypeOnlineConnector)) {Import-Module -Name SkypeOnlineConnector -ErrorAction SilentlyContinue}
    If ( Get-Module -Name SkypeOnlineConnector) {
        If ( !($global:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
        If ( $global:myOffice365Services['Office365CredentialsMFA']) {
            Write-Host "Connecting to Skype for Business Online using $($global:myOffice365Services['Office365Credentials'].username) with Modern Authentication .."
            $Parms = @{'Username' = ($global:myOffice365Services['Office365Credentials']).username}
        }
        Else {
            Write-Host "Connecting to Skype for Business Online using $($global:myOffice365Services['Office365Credentials'].username) .."
            $Parms = @{'Credential' = $global:myOffice365Services['Office365Credentials']}
        }
        $global:myOffice365Services['SessionSFB'] = New-CsOnlineSession @Parms
        If ( $global:myOffice365Services['SessionSFB'] ) {
            Import-PSSession -Session $global:myOffice365Services['SessionSFB'] -AllowClobber
        }
    }
    Else {
        Write-Error -Message 'Cannot connect to Skype for Business Online - problem loading module.'
    }
}
Function global:Get-Office365Credentials {
    $global:myOffice365Services['Office365Credentials'] = $host.ui.PromptForCredential('Office 365 Credentials', 'Please enter your Office 365 credentials', '', '')
    $local:MFAenabledModulePresence= $false
    # Check for MFA-enabled modules 
    If ( (Get-Module -Name 'Microsoft.Exchange.Management.ExoPowershellModule') -or (Get-Module -Name 'MicrosoftTeams')) {
        $local:MFAenabledModulePresence= $true
    }
    Else {
        # Check for MFA-enabled modules with version dependency
        $MFAMods= @('SkypeOnlineConnector|7.0', 'Microsoft.Online.Sharepoint.PowerShell|16.0')
	ForEach( $MFAMod in $MFAMods) {
            $local:Item = ($local:MFAMod).split('|')
            If( (Get-Module -Name $local:Item[0] -ListAvailable)) {
                $local:MFAenabledModulePresence= $local:MFAenabledModulePresence -or ((Get-Module -Name $local:Item[0] -ListAvailable).Version -ge [System.Version]$local:Item[1] )
            }
        }
    }
    If( $local:MFAEnabledModulePresence) {
        $global:myOffice365Services['Office365CredentialsMFA'] = Get-MultiFactorAuthenticationUsage
    }
    Else {
        $global:myOffice365Services['Office365CredentialsMFA'] = $false
    }
    Get-TenantID
}
Function global:Get-Office365Tenant {
    $global:myOffice365Services['Office365Tenant'] = Read-Host -Prompt 'Enter tenant ID, e.g. contoso for contoso.onmicrosoft.com'
}
Function Get-FileName($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

#------------------------------------------------------------------------------------------
# 0. Modules Connection Status
#------------------------------------------------------------------------------------------
# Enable PIM to Get Access to Tenant
Connect-AzureActiveDirectory
Connect-SkypeOnline

#------------------------------------------------------------------------------------------
# 1. Obtain a service number from Skype for Business Online
#------------------------------------------------------------------------------------------
$PhaseNumber = 1
$NeedNewNumbers = $false

# Finding How Many Available Numbers Exists for Service and not Assigned
Write-Host "[Phase $PhaseNumber] - Searching for Available service numbers in the environment which are not assigned: "  -NoNewline
$ServiceNumbersAvailableInTenant = Get-CsOnlineTelephoneNumber -InventoryType Service -IsNotAssigned -WarningAction:SilentlyContinue
Write-Host "$(($ServiceNumbersAvailableInTenant | Measure-Object).Count) Available Service number in the Tenant are ready to use" -ForegroundColor Green

Write-Host "[Phase $PhaseNumber] - Searching for Available service numbers to purchase in the environment: " -NoNewline
$ServiceNumbersPurchasedInTenant = CsOnlineTelephoneNumberAvailableCount -InventoryType Service
Write-Host "$($ServiceNumbersPurchasedInTenant.Count) Available service numbers available to Purchase" -ForegroundColor Green

# Searching for Available Numbers Under UK
If ($ServiceNumbersPurchasedInTenant.Count -lt 1 -and $ServiceNumbersAvailableInTenant.Count -ge 1 -and $NeedNewNumbers -eq $true)
{
    $NumbersToPurchase = Read-Host "How many number would you like to purchase betweewn: [0-$($ServiceNumbersPurchasedInTenant.Count)]"
    #for ($i=0; $i -lt $NumbersToPurchase ;$i++)
    #{
        $MySearch = Search-CsOnlineTelephoneNumberInventory -InventoryType Service -Region EMEA -Country GB -City WLS_CA  -Area ALL -AreaCode '29' -Quantity $NumbersToPurchase
        Write-host $MySearch.Reservations[0].Numbers.Number "In GB, Cardiff was found, Purchasing and Adding to the tenant" -ForegroundColor Green
        Select-CsOnlineTelephoneNumberInventory -ReservationId $MySearch.ReservationId -TelephoneNumbers $MySearch.Reservations[0].Numbers.Number -Region EMEA -Country GB -City WLS_CA -Area All
    #}
}
elseif ($ServiceNumbersPurchasedInTenant.Count -gt 1)
{
    Write-Warning "There are available service numbers to use"
}
else
{
    Write-Warning "There are no Service Number Available to Purchase"
}


#------------------------------------------------------------------------------------------
# 2. Obtain a free Phone System - Virtual User license
#------------------------------------------------------------------------------------------
$PhaseNumber = 2
Write-Host "[Phase $PhaseNumber] - Searching for the AzureAD Virtual Subscriber SKU: " -NoNewline
$VirtualUserOnlineSKU = Get-AzureADSubscribedSku | Select SkuPartNumber,SkuID | Where {$_.SkuPartNumber -eq "PHONESYSTEM_VIRTUALUSER"}
Write-Host "$($VirtualUserOnlineSKU.SkuPartNumber)" -ForegroundColor Yellow

Write-Host "[Phase $PhaseNumber] - Searching for the AzureAD Tenant Name: " -NoNewline
$TenantDomainName = ((Get-AzureADTenantDetail).VerifiedDomains | Where {$_.Initial -eq $true} | Select-Object Name).Name
Write-Host "$TenantDomainName" -ForegroundColor Yellow

#------------------------------------------------------------------------------------------
# 3. Create the resource accounts
#------------------------------------------------------------------------------------------
$PhaseNumber = 3
$UsersCounter = 1
$CreatedUsers = @()
$NeededResourceAccounts = $true

    #Loading CSV File
    if (!$UsersCSV) 
        {
            $InputFile = Get-FileName -initialDirectory "$(Split-Path -Path $MyInvocation.MyCommand.Path)"
        }
    else 
        {
            if (Test-Path $CSVFile) {$InputFile = $CSVFile}
        }
    $UsersCSV = import-csv $InputFile -Delimiter ','


If ($UsersCSV.Count -gt 0 -and $NeededResourceAccounts -eq $true)
{
    foreach ($User in $UsersCSV)
    {
        try
        {
            Write-Host "[Phase $PhaseNumber] - [$UsersCounter out of $($UsersCSV.Count)] Create Resouce Account for Users - Creating: 'zRA-$(($User.UserName -split "@")[0])@$TenantDomainName' , 'zRA - $(($User.UserName -split "@")[0]) - $(($user.ServiceNumber).Replace('-',''))' > $($User.DestinationNumber)"
            $CreatedUsers += New-CsOnlineApplicationInstance -UserPrincipalName "zRA-$(($User.UserName -split "@")[0])@$TenantDomainName" -ApplicationId “ce933385-9390-45d1-9512-c8d228074e07” -DisplayName "zRA - $(($User.UserName -split "@")[0]) - $(($user.ServiceNumber).Replace('-','')) > $($User.DestinationNumber)"
        }
        catch
        {
            Write-Error "Error Creating New Resource Account" -Exception $error
        }
        
        $UsersCounter += 1
    }

}
elseif ($NeededResourceAccounts -eq $false)
{
    Write-Warning "No need in creating Resource Accounts at the moment"
}
else
{
    Write-Error "Can't Load CSV File with Users Information" ; break
}


#------------------------------------------------------------------------------------------
# 4. Assign the Phone System - Virtual user license to the resource account
#------------------------------------------------------------------------------------------
$PhaseNumber = 4
$AutoAttendantApplicationInstances = Get-CsOnlineApplicationInstance | Where {$_.ApplicationId -eq "ce933385-9390-45d1-9512-c8d228074e07" -and $_.PhoneNumber -eq $null}
$AssignmentCounter = 1

if ($CreatedUsers.Count -le $AutoAttendantApplicationInstances.Count)
{
    Write-Host "[Phase $PhaseNumber] - Assign Phone System Virtual User License" -ForegroundColor Yellow
    
    #Define License Type
    $License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
    $License.SkuId = $VirtualUserOnlineSKU.SkuId
    $LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
    $LicensesToAssign.AddLicenses = $License
    
    #Assign License
    foreach ($CreatedUser in $CreatedUsers)
    {
        Write-Host "[Phase $PhaseNumber] - [$AssignmentCounter out of $($CreatedUsers.Count)] - Assigning $($CreatedUser.DisplayName) with GB UsageLocation ; " -NoNewline
        Set-AzureADUser -ObjectId $CreatedUser.ObjectId -UsageLocation GB
        Write-Host "Assigning $($CreatedUser.DisplayName) with Virtual User License"
        Set-AzureADUserLicense -ObjectId $($CreatedUser.ObjectID) -AssignedLicenses $LicensesToAssign
        $AssignmentCounter += 1
    }
}
else {Write-Error "The Amount of Created Users do not match the amount available"}

#------------------------------------------------------------------------------------------
# 5. Assign a service phone number to the resource account you just assigned licenses to
#------------------------------------------------------------------------------------------
$PhaseNumber = 5
$ServicePhoneNumberAssignmentCounter = 1
$ServicePhoneNumberAssignmentDetails = @()
$ServicePhoneNumberAssignments = @()

foreach ($CreatedUser in $CreatedUsers)
{
    #$ServiceNumbersPurchasedInTenant = Get-CsOnlineTelephoneNumber -InventoryType Service -IsNotAssigned
    $tempServiceNumber = $false
    $tempServiceNumber = Get-CsOnlineTelephoneNumber -InventoryType Service -IsNotAssigned -TelephoneNumber $($($CreatedUser.DisplayName.Split("+")[1]) -replace " >","")
    if ($tempServiceNumber -ne $null)
    {
        Write-Host "[Phase $PhaseNumber] - [$ServicePhoneNumberAssignmentCounter out of $($CreatedUsers.Count)] - Assigning " -NoNewline ; Write-Host -ForegroundColor Yellow "'$($CreatedUser.DisplayName)'" -NoNewline ; Write-Host " with Service number:" -NoNewline ; Write-Host " +$($tempServiceNumber.Id)" -ForegroundColor Yellow
        $ServicePhoneNumberAssignments += Set-CsOnlineVoiceApplicationInstance -Identity $CreatedUser.UserPrincipalName -TelephoneNumber $tempServiceNumber.Id -Verbose
        $ServicePhoneNumberAssignmentCounter += 1
        $CSVSourceNumber =  ($CreatedUser.DisplayName.Split("+") -replace ">" -replace " " -replace "-")[1]
        $CSVDestinationNumber = ($CreatedUser.DisplayName.Split("+") -replace ">" -replace " " -replace "-")[2]
        $ServicePhoneNumberAssignmentDetails += New-Object PSObject -property @{
                        DisplayName          = $CreatedUser.DisplayName
                        UserPrincipalName    = $CreatedUser.UserPrincipalName
                        ObjectID             = $CreatedUser.ObjectId
                        CSVUSerName          = $null
                        CSVOldPhoneNumber    = $CSVSourceNumber
                        CSVNewPhoneNumber    = $CSVDestinationNumber
                        DestinationNumber    = $($CreatedUser.DisplayName.Split("+")[1])
                        ServiceNumber        = $tempServiceNumber.Id
                        CityCode             = $tempServiceNumber.CityCode
                    }
    }             
}


#------------------------------------------------------------------------------------------
# 6. Create a Phone System auto attendant
#------------------------------------------------------------------------------------------
$PhaseNumber = 6
$AutoAttendantCreationCounter = 1
$AutoAttendantCreationDetails = @()

foreach ($CreatedUser in $CreatedUsers)
{
    Write-Host "[Phase $PhaseNumber] - [$AutoAttendantCreationCounter out of $($CreatedUsers.Count)] - Creating '$($CreatedUser.DisplayName)' with AutoAttendant"
    
    $operatorObjectId = (Get-CsOnlineUser $CreatedUser.UserPrincipalName).ObjectId
    $operatorTelId = (Get-CsOnlineUser $CreatedUser.UserPrincipalName).lineURI
    #$operatorObjectId = $ServicePhoneNumberAssignment.ObjectId
    $operatorEntity = New-CsAutoAttendantCallableEntity -Identity $operatorObjectId -Type ApplicationEndpoint
    #$operatorEntity = New-CsAutoAttendantCallableEntity -Identity "tel:+19178099928" -Type User
    #$greetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt "Welcome to Contoso!"
    #$menuOptionZero = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone0
    $menuOptionDisconnect = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic

    $SourceNumber = $null
    $DestinationNumber  = $null
    $SourceNumber =  (($CreatedUser.DisplayName.Split("+") -replace ">" -replace " " -replace "-")[1] -replace (‘(.)’,’$1;’)).trim(';')
    $DestinationNumber = (($CreatedUser.DisplayName.Split("+") -replace ">" -replace " " -replace "-")[2] -replace (‘(.)’,’$1;’)).trim(';')

    #$menuPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt "Hello, Number $($ServicePhoneNumberAssignment.OldNumber) is no longer available, Please Dial $($ServicePhoneNumberAssignment.CSVNewPhoneNumber) to reach your extension."
    # Hello. Telephone number; 1;5;7;0;8;0;0;2;8;7;8; has now changed. Please redial using the following number; 1;6;1;0;8;8;2;6;7;3;5. That number again is; 1;6;1;0;8;8;2;6;7;3;5. Thank you and goodbye.
    #Hello. Telephone number; 1;5;7;0;8;0;0;2;8;7;8; has now changed. Please redial using the following number; 1;6;1;0;8;8;2;6;7;3;5. That number again is; 1;6;1;0;8;8;2;6;7;3;5. Thank you and goodbye.
    $menuPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt "Hello. Telephone number; $SourceNumber; has now changed. Please redial using the following number; $DestinationNumber;. That number again is; $DestinationNumber;. Thank you and goodbye."
    
    $defaultMenu = New-CsAutoAttendantMenu -Name "Default menu" -MenuOptions @($menuOptionDisconnect)
    #$defaultCallFlow = New-CsAutoAttendantCallFlow -Name "Default call flow" -Greetings @($greetingPrompt) -Menu $defaultMenu
    $defaultCallFlow = New-CsAutoAttendantCallFlow -Name "Default call flow" -Greetings @($menuPrompt) -Menu $defaultMenu

    $AutoAttendantCreation = $null
    $AutoAttendantCreation = New-CsAutoAttendant -Name "zAA - $($CreatedUser.DisplayName)" -DefaultCallFlow $defaultCallFlow -Language "en-GB" -TimeZoneId "UTC"
    #$aa = New-CsAutoAttendant -Name "$($ServicePhoneNumberAssignment.DisplayName) - AutoAttendant" -DefaultCallFlow $defaultCallFlow -EnableVoiceResponse -CallFlows @($afterHoursCallFlow) -CallHandlingAssociations @($afterHoursCallHandlingAssociation) -Language "en-US" -TimeZoneId "UTC" -Operator $operatorEntity -InclusionScope $inclusionScope

    $ApplicationInstanceAssociation = $null
    $ApplicationInstanceAssociation = New-CsOnlineApplicationInstanceAssociation -Identities $operatorObjectId -ConfigurationId $AutoAttendantCreation.Id -ConfigurationType AutoAttendant

    $AutoAttendantCreationCounter += 1
    $AutoAttendantCreationDetails += New-Object PSObject -property @{
                    DisplayName          = $CreatedUser.DisplayName
                    UserPrincipalName    = $CreatedUser.UserPrincipalName
                    ObjectID             = $CreatedUser.ObjectId
                    AAName               = $AutoAttendantCreation.Name
                    AAStatus             = $AutoAttendantCreation.Status
                    ApplicationInstanceAssociation = $ApplicationInstanceAssociation.Results
                    CSVUSerName          = $null
                    CSVOldPhoneNumber    = $SourceNumber
                    CSVNewPhoneNumber    = $DestinationNumber
                    OldNumber            = (($CreatedUser.DisplayName.Split("-")).Trim())[1]
                    ServiceNumber        = $ServiceNumbersPurchasedInTenant[0].Id
                }
}