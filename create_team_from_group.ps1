#Requires -Version 5.1

<#
	.SYNOPSIS
		Teams Synchronisation Education

	.DESCRIPTION
        Long term supported path is to use School Data Sync (SDS) - this Teams synchronisation script should only be used in the interim where SDS is not available.  

        The script keeps groups created by the existing IDM solution in sync. with a team, without direct integration between SDS and the class management system. (This script becomes the integration).

        The script is designed to run from Azure Automation under a schedule.

        High level overview:
            1. Checks master group $MasterGroup for groups that should be enabled for Teams
            2. Determines if Office 365 Group has been created
            3. If Office 365 Group is not created, creates group, then creates Team using Graph endpoint
            4. Keeps group in sync

    .PARAMETER ErrorHook
        A webhook to a teams channel. This is used to post any error information to a Microsoft Teams channel.

    .PARAMETER StatusHook
        A webhook to a teams channel. This is used to post status information, for instance, how many teams were created
    
    .PARAMETER MasterGroup
        The MasterGroup is the parent group, the members of this group are the groups that you would like teams created for.
    
    .PARAMETER regexGroupMatch
        Use this pattern to match the names of your AAD groups containing class membership, and use groupings () for use in $TeamName inside the loop below.

    .PARAMETER Tenancy
        Friendly name for your tenancy that's used in the WebHooks - useful if you have a PROD and a TEST tenancy
    
    .PARAMETER TenantName
        The .onmicrosoft.com name of your tenancy

    .PARAMETER ApplicationID
        The ID of the Azure AD Application that is used to connect to the tenancy. This will need to be provisioned.

    .PARAMETER Username
        The username of the account used to manage the team - there must be a matching azure automation credential under the same Username.

    .NOTES
    
        No warranty or support is provided with this script.

		Cam Murray
		Field Engineer - Microsoft
		cam.murray@microsoft.com
		
		Last update: 11/2/2020

	.LINK
		about_functions_advanced

#>
Param
(
	### Customise these values
	# Webhook urls will have format like https://outlook.office.com/webhook/<uuid>/IncomingWebhook/<uuid>
    [String]$ErrorHook="<TeamsChannelWebhookURL>",
    [String]$StatusHook="<TeamsChannelWebhookURL>",
    [String]$MasterGroup="MyO365GroupContaining_TeamsEnabledGroups",
    [String]$regexGroupMatch = "STUDENTS_(CLS|CRSE|PROG)_.*_.*_(UGRD|PGRD|PGRO)_(.*)",
    [String]$Tenancy="<StringIdentifyingTestorProdinMessages>",
    [String]$TenantName="<AAD Tenant Name (xxx.onmicrosoft.com)",
    [String]$ApplicationID="<ID of application registration>",
    [String]$Username="<AAD User Name of account to manage teams, AND name of Automation credential containing matching password>"
	###
)

Function PostChannel 
{
    <#
    
        Post status to Channel
        $Hook = Web API hook for the channel
        $Text = Text to post
        $RetryMax = Amount of attempts to post to the channel
        $Tenancy = Name of the tenancy
    
    #>
    Param
    (
        $Hook,
        $Text,
        $RetryMax = 3,
        $Tenancy
    )
    
    $Complete = $False

    $channelPostBody = ConvertTo-JSON @{
        text = "<p><b>$($Tenancy) Tenancy</b></p>$Text"
         }

    $RetryCounter=0 # for retry counter
    While($Complete -eq $False) 
    {
        $RetryCounter++
        $Return = Invoke-RestMethod -uri $Hook -Method Post -body $channelPostBody -ContentType 'application/json'
        If($Return -eq 1 -or $RetryCounter -ge $RetryMax ) { $Complete=$True } else { Write-Output -InputObject "$(Get-Date) Post to teams failed return $Return, retry count $RetryCounter"; Start-Sleep 10}
    }

    Return $Return
}

Function ExpandMembers 
{
    Param (
        [Parameter()][string]$groupId,
        $AppAuthParams
    )

    $members = Get-GraphOData -Endpoint "/beta/groups/$($groupId)/members" -AppAuthParams $AppAuthParams

    $allMembers = @()
    $allMembers += $members | where-object { $_.'@odata.type' -eq '#microsoft.graph.user' } 


    Write-Verbose "$($allMembers.length) members discovered"

    $toExpand = @()
    $toExpand = $members | where-object { $_.'@odata.type' -eq '#microsoft.graph.group' }

    $nextExpand = @()
    while ($toExpand) {
        Write-Verbose "Expanding $($toExpand.length) groups"

        $nextExpand = @() 

        foreach ($group in $toExpand) {

            $members = Get-GraphOData -Endpoint "/beta/groups/$($group.id)/members" -AppAuthParams $AppAuthParams

            $nextExpand += $members | where-object { $_.'@odata.type' -eq '#microsoft.graph.group' }
            $allMembers += $members | where-object { $_.'@odata.type' -eq '#microsoft.graph.user' }     

            Write-Verbose "$($group.DisplayName): $($allMembers.length) members discovered"
        }
        $toExpand = $nextExpand
    }
    return $allMembers

}  

Function Invoke-LoadADAL
{
    <#
    
        Finds a suitable ADAL library from AzureAD Preview and uses that
        This prevents us having to ship the .dll's ourself.

    #>
    $AadModule = Get-Module -Name "AzureADPreview" -ListAvailable

    $Latest_Version = ($AadModule | Select-Object version | Sort-Object)[-1]
    $aadModule      = $AadModule | ? { $_.version -eq $Latest_Version.version }
    $adal           = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
    $adalforms      = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"

    [System.Reflection.Assembly]::LoadFrom($adal) | Out-Null
    [System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null
}

Function Get-AccessToken 
{
    <#
    
        Fetch the Access Token using ADAL libraries
    $AppAuthParams = New-Object -TypeName PSObject -Property @{
    TenantName=$TenantName
    ClientID=$ApplicationID
    Secret=$ApplicationSecret
    Resource=$Resource
    Username=$Username
    Password=$Password
    }
    #>
    Param
    (
        $AppAuthParams
    )

    $authority          = "https://login.microsoftonline.com/$($AppAuthParams.TenantName)"
    $authContext        = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority

    $Token = $null

    # First attempt to silently (from cache or RT) get a token
    try {
        $Token = $authContext.AcquireTokenSilentAsync($($AppAuthParams.Resource),$($AppAuthParams.ClientID),[Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier]::AnyUser).GetAwaiter().GetResult();
        Write-Verbose "$(Get-Date) Got token from cache or RT"
        $Token = $Token.AccessToken
    }
    catch {
        Write-Verbose "$(Get-Date) Getting token from AAD"
        if($($AppAuthParams.secret))
        {
            $ClientCredential   = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.ClientCredential" -ArgumentList @($($AppAuthParams.ClientID),$($AppAuthParams.Secret))
            $authResult         = $authContext.AcquireTokenAsync($($AppAuthParams.Resource),$ClientCredential)
        }
        
        if($($AppAuthParams.Username))
        {
            $UserPasswordCredential     = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserPasswordCredential" -ArgumentList @($($AppAuthParams.Username),$($AppAuthParams.Password))
    
            $authResult = [Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContextIntegratedAuthExtensions]::AcquireTokenAsync($authContext, $($AppAuthParams.Resource), $($AppAuthParams.ClientID), $UserPasswordCredential)     
        }
        $Token = $authResult.Result.AccessToken   
    }

    if($Token -eq $null)
    {
        Throw "Failed to get Access Token either from cache, RT, or by access to AAD"
    }
    else 
    {
        Return $Token
    }
    
}

Function Get-GraphData
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $Endpoint,
        $AppAuthParams
    )

    $Headers = Get-AuthorizationHeaders -AppAuthParams $AppAuthParams

    Return Invoke-RestMethod -Method GET -Uri $Endpoint -UseBasicParsing -Headers $Headers


}

Function Get-AuthorizationHeaders
{
    [CmdletBinding()]
    param (
        $AppAuthParams
    )

    Try
    {
        $Token = Get-AccessToken -AppAuthParams $AppAuthParams
    } catch {
        Throw "Failed to get Access Token"
    }
    

    Return @{"Authorization"="Bearer $Token"}
}

Function Get-GraphRestMethod
{
    <#
    
        Get information back from Graph using a REST call

        Handles the backing off of REST in the event of throttling

    #>
    Param(
        $Uri,
        $Application,
        $Resource,
        $Token
    )

    # Used to back off the API. In seconds
    $BackOff = 30 # First back off is 30 seconds
    $BackOffInc = 60 # Next back off is backoff + 60 seconds
    $BackOffTry = 5 # Try this amount of times
    $CurrentTry = 0

    # Obtain the authorization header
    $Headers = Get-AuthorizationHeaders -AppAuthParams $AppAuthParams

    # Start of query
    While($True)
    {
        $CurrentTry++

        Try 
        {
            Return Invoke-RestMethod -Headers $Headers -Uri $Uri
        }
        Catch 
        {
            Write-Verbose "$(Get-Date) Get-GraphRestMethod $($uri) $($_.Exception.Response.StatusCode.value__ ) Description $($_.Exception.Response.StatusDescription) backing off for $BackOff seconds"

            If($CurrentTry -eq $BackOffTry)
            {
                # Failed
                Break
            }
            else 
            {
                Start-Sleep $BackOff
                # Increase next back off
                $BackOff += $BackOffInc
            }
        }
    }
    
}

Function Get-GraphOData
{
    Param
    (
        $Token,
        $Resource="https://graph.microsoft.com",
        $Endpoint,
        $ODataParams,
        [Switch] $NoNext
    )

    # Structure the URI
    $Uri = "$($Resource)/$Endpoint"
    If($ODataParams)
    {
        $Uri += "?$ODataParams"
    }

    Write-Verbose "$(Get-Date) Getting Graph OData $($Application.TenantName) URI $Uri"

    # Count for verbosity
    $Count = 0

    $Return = Get-GraphRestMethod -Uri $Uri -AppAuthParams $AppAuthParams

    # Output first chunk
    $Return.Value

    # Count for verbosity
    $Count += $($Return.Value).Count
    Write-Verbose "$(Get-Date) Request initial count $($Count)"

    # Continue NextLinks
    If(!$NoNext)
    {
        $NextLink = $Return."@odata.nextLink"
        While($NextLink)
        {
    
            $Return = Get-GraphRestMethod -Uri $NextLink -Application $Application -Resource $Resource

            # Output data
            $Return.Value
    
            # Count for verbosity
            $PageCount = $($Return.Value).Count
            $Count += $PageCount
            Write-Verbose "$(Get-Date) Page count $($PageCount) total count $($Count)"
    
            # Determine next link if applicable
            $NextLink = $Return."@odata.nextLink"
        }
    }

}


Function Get-TeamInfo
{
    [CmdletBinding()]
    param (
        $GroupID,
        $AppAuthParams
    )

    $Headers = Get-AuthorizationHeaders -AppAuthParams $AppAuthParams

    $Uri = "https://graph.microsoft.com/beta/teams/$($GroupID)"

    $Return = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers -UseBasicParsing

    Return $Return
}

# Load the AAD Libraries
Invoke-LoadADAL

# Variables used, stats etc
$StartTime = (Get-Date)
$TeamsCreated = 0
$UsersAdded = 0
$UsersRemoved = 0
$Errors = 0

# Create credential object
$securePwd = (Get-AutomationPSCredential -Name $Username).Password

# Connect

$AppAuthParams = New-Object -TypeName PSObject -Property @{
    TenantName=$TenantName
    ClientID=$ApplicationID
    Secret=$ApplicationSecret
    Resource="https://graph.microsoft.com"
    Username=$Username
    Password=$securePwd
}

# Variables required
$syncedGroups = @{}

###

# Try find the MasterGroup
$objMasterGroup = Get-GraphOData -Endpoint "/beta/groups" -ODataParams "`$filter=displayName eq '$($MasterGroup)'" -AppAuthParams $AppAuthParams

If($objMasterGroup.displayName -ne $MasterGroup) {
    PostChannel -Hook $ErrorHook -Text "Cannot find the master group, or there are multiple groups that match $MasterGroup"
    Write-Error "Cannot find the master group, or there are multiple groups that match $MasterGroup"
    $Errors++
    exit
}

Write-Output -InputObject "$(Get-Date) Getting groups to sync to teams..."

# Get group members of master groups
$syncGroups = Get-GraphOData -Endpoint "/beta/groups/$($objMasterGroup.id)/members" -AppAuthParams $AppAuthParams

# Determine members and also what the Team should be called
ForEach($grp in $syncGroups) {
    If($grp.DisplayName -match $regexGroupMatch) {
		### CUSTOMISE THIS to set the team name and type
		# If $TeamType -eq "PROG" then team will a standard team with no extensions, otherwise it will be created with EDU_Class template.
        $TeamName = $matches[1] + "-" + $matches[3]
        $TeamType = $matches[1]
		###

        # Try find Office 365 Group under same name
        $OGroup = Get-GraphOData -Endpoint "/beta/groups" -ODataParams "`$filter=MailNickname eq '$TeamName'" -AppAuthParams $AppAuthParams

        # Get Team information
        if($OGroup.id)
        {
            $TeamInfo = Get-TeamInfo -GroupID $OGroup.id -AppAuthParams $AppAuthParams
        }

        $Errored = $false

        Try {
            # $Members = $(Get-AzureADGroupMember -ObjectId $grp.ObjectId -all:$True| Select-Object ObjectId)
            $Members = ExpandMembers -groupId $grp.id -AppAuthParams $AppAuthParams | Select-Object id,'@odata.type'
        } Catch {
            $Errored = $true
            Write-Error "Failed to get group members for group object id $($grp.ObjectId)"
            PostChannel -Hook $ErrorHook -Text "Failed to get group members for group object id $($grp.ObjectId)"
        }

        If($Errored -eq $false) {
            # Add this group to the mapping list
            $grpInfo = New-Object -TypeName PSObject -Property @{
                OGroup = $($OGroup.id)
                TeamInfo = $TeamInfo
                TeamName = $TeamName
                Members = $Members
                TeamType = $TeamType
           }

            Write-Output -InputObject "$($grp.DisplayName) to $($TeamName)"
            
            $syncedGroups.Add($TeamName,$grpInfo)
        } Else {
            # Do not add this group to the mapping list - doing so would wipe the existing group
            # There may need to be clauses here to remove the group and Team?
        }
    } else {
        PostChannel -Hook $ErrorHook -Text "$($grp.DisplayName) doesn't match expected regex match of $regexGroupMatch"
        Write-Error "$($grp.DisplayName) doesn't match expected regex match of $regexGroupMatch"
        $Errors++
    }
}

Function Set-GroupForTeams 
{
    Param
    (
        $AppAuthParams,
        $GroupID,
        $Type
    )

    $Headers = Get-AuthorizationHeaders -AppAuthParams $AppAuthParams

    # Used to back off the API. In seconds
    $BackOff = 10 # First back off is 30 seconds
    $BackOffInc = 10 # Next back off is backoff + 30 seconds
    $BackOffTry = 20 # Try this amount of times
    $CurrentTry = 0

    if($type -eq "PROG")
    {
        $Template = "https://graph.microsoft.com/beta/teamsTemplates('standard')"
    }
    else
    {
        $Template = "https://graph.microsoft.com/beta/teamsTemplates`('educationClass'`)"
    }
    

    # Object for enablement
    $EnablementObject = New-Object -TypeName PSOBject -Property @{
        "template@odata.bind"= $Template
        "group@odata.bind"= "https://graph.microsoft.com/beta/groups/$GroupID"
    }

    $json = (ConvertTo-Json $EnablementObject)

    while($true)
    {
        $CurrentTry++

        Write-Output -InputObject "$(Get-Date) Attempting to teams enable $($GroupID) try $($CurrentTry)"

        try {
            $Result = Invoke-RestMethod -Method POST -ContentType 'application/json' -Headers $Headers -Uri "https://graph.microsoft.com/beta/teams" -UseBasicParsing -Body $json
            Write-Output -InputObject "$(Get-Date) Teams enabled $($GroupID)"
            Return $Result
        }
        catch {
            If($CurrentTry -eq $BackOffTry)
            {
                # Failed
                Break
            }
            else 
            {
                Start-Sleep $BackOff
                # Increase next back off
                $BackOff += $BackOffInc
            } 
        }

    }

}

Function New-AADGroup
{
    Param
    (
        $AppAuthParams,
        $displayName,
        $mailNickname,
        $Type
    )

    Write-Output "$(Get-Date) AAD Creating Group $($displayName) type $($Type)"

    if($type -eq "PROG")
    {
        # These params are taken from the Teams PowerShell module as of 0.9.8
        $object = New-Object -TypeName PSObject -Property @{
            creationOptions=@(
                "ExchangeProvisioningFlags:481"
            )
            displayName=$displayName
            groupTypes=@(
                "Unified"
            )
            mailEnabled=$true
            mailNickname=$mailNickname
            securityEnabled=$False
            visibility="Private"
        }

        $uri = "https://graph.microsoft.com/edu/groups"
    } else {
        # These params are taken from the Teams PowerShell module as of 0.9.8
        $object = New-Object -TypeName PSObject -Property @{
            creationOptions=@(
                "ExchangeProvisioningFlags:461",
                "classAssignments"
            )
            displayName=$displayName
            extension_fe2174665583431c953114ff7268b7b3_Education_ObjectType="Section"
            groupTypes=@(
                "Unified"
            )
            mailEnabled=$true
            mailNickname=$mailNickname
            securityEnabled=$False
            visibility="HiddenMembership"
        }

        $uri = "https://graph.microsoft.com/edu/groups"
    }

    if($uri)
    {
        $Headers = Get-AuthorizationHeaders -AppAuthParams $AppAuthParams

        $json = ConvertTo-Json $object
    
        $Result = Invoke-RestMethod -Method POST -Uri $uri -Headers $Headers -UseBasicParsing -Body $json -ContentType 'application/json'
    }

    Return $Result
}

Function Add-GroupMember
{
    Param
    (
        $UserID,
        $GroupID,
        $AppAuthParams
    )

    Write-Output -InputObject "$(Get-Date) Adding $($UserID) to group $($GroupID)"

    $Headers = Get-AuthorizationHeaders -AppAuthParams $AppAuthParams

    $obj = New-Object -TypeName PSObject -Property @{
        "@odata.id"="https://graph.microsoft.com/v1.0/directoryObjects/$($UserID)"
    }

    $Uri = "https://graph.microsoft.com/beta/groups/$($GroupID)/members/`$ref"

    Invoke-RestMethod -Method POST -Uri $Uri -ContentType 'application/json' -Headers $Headers -Body (ConvertTo-Json $obj) -UseBasicParsing
}

Function Remove-GroupMember
{
    Param
    (
        $UserID,
        $GroupID,
        $AppAuthParams
    )

    Write-Output -InputObject "$(Get-Date) Removing $($UserID) from group $($GroupID)"

    $Headers = Get-AuthorizationHeaders -AppAuthParams $AppAuthParams

    $Uri = "https://graph.microsoft.com/beta/groups/$($GroupID)/members/$($UserID)/`$ref"

    Invoke-RestMethod -Method DELETE -Uri $Uri -Headers $Headers -UseBasicParsing
}


# Create Office 365 Groups where required
$NewTeams = @()
ForEach($grpName in $syncedGroups.Keys)
{
    $thisGroup = $syncedGroups[$grpName]

    # Determine if group needs creation
    If(!$thisGroup.OGroup)
    {
        $Result = New-AADGroup -displayName $grpName -mailNickname $grpName -Type $($thisGroup.TeamType) -AppAuthParams $AppAuthParams

        if($Result.objectId)
        {
            Write-Output "$(Get-Date) Successfully created AAD Group for $grpName as $($Result.objectId)"

            $thisGroup | Add-Member -Name OGroup -MemberType NoteProperty -Value $Result.objectId -Force

            $NewTeams += $thisGroup
        }
        else 
        {
            PostChannel -Hook $ErrorHook -Text "Error in AAD stage creating group $($grpName)"
        }
    }
}

Write-Output -InputObject "$(Get-Date) Creating any required teams..."
# Create any teams required
if($NewTeams.Count -gt 0)
{
    ForEach($t in $NewTeams)
    {
        $Result = Set-GroupForTeams -GroupID $($t.OGroup) -AppAuthParams $AppAuthParams -Type $($t.TeamType)
        if($Result.id)
        {
            Write-Output "$(Get-Date) Successfully teams enabled AAD Group $($t.TeamName) as $($Result.id)"
            $TeamsCreated++
        }
        else 
        {
            PostChannel -Hook $ErrorHook -Text "Failed to teams enable $($t.TeamName)"   
        }
    }
}

Write-Output -InputObject "$(Get-Date) Syncing membership..."

# Loop through group objects
ForEach($grpName in $syncedGroups.Keys) {
    
    $thisGroup = $syncedGroups[$grpName]

    # Sync membership
    If($thisGroup.OGroup) {

        Write-Output -InputObject "$(Get-Date) Syncing $grpName membership"

        # Get Team Members/Owners
        Try {
            $TeamMembers = Get-GraphOData -Endpoint "/beta/groups/$($thisGroup.OGroup)/members" -AppAuthParams $AppAuthParams
            $TeamOwners = Get-GraphOData -Endpoint "/beta/groups/$($thisGroup.OGroup)/owners" -AppAuthParams $AppAuthParams
        } Catch {
            PostChannel -Hook $ErrorHook -Text "Failed to get Team membership and or owners for $grpName"
            Write-Error "Failed to get Team membership for $grpName exiting"
            $Errors++
            Continue
        }

        if($TeamOwners.userPrincipalName -notcontains $Username)
        {
            PostChannel -Hook $ErrorHook -Text  "$grpName $($thisGroup.OGroup) doesn't contain an owner called $username"
            Write-Error "$grpName $($thisGroup.OGroup) doesn't contain an owner called $username"
            $Errors++
            Continue
        }

        $TeamMembers = $TeamMembers.id
        $TeamOwners = $TeamOwners.id

        # Determine group members to add
        ForEach($oid in $thisGroup.Members.id) {
            If($($TeamMembers) -notcontains $oid -and $($TeamOwners) -notcontains $oid) {
                Add-GroupMember -GroupID $($thisGroup.OGroup) -UserID $oid -AppAuthParams $AppAuthParams
                $UsersAdded++
            }
        }

        # Determine team members to remove
        ForEach($oid in $($TeamMembers)) {
            # If the Group Object doesn't contain this user, and this user is not an owner of the team - then remove
            If($thisGroup.Members.id -notcontains $oid) {
                Remove-GroupMember -GroupId $($thisGroup.OGroup) -UserID $oid -AppAuthParams $AppAuthParams
                $UsersRemoved++
            }    
        }

    } Else {
        Write-Output -InputObject "$(Get-Date) Skipping $grpName as it doesn't exist"
    }

}

$TimeTaken = $((Get-Date)-$StartTime).TotalMinutes
PostChannel -Hook $StatusHook -Text "Completed running in $([math]::Round($TimeTaken,2)) minutes. Teams created $TeamsCreated. Users Added $UsersAdded. Users Removed $UsersRemoved. Errors $Errors"
