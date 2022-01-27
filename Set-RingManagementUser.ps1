<#
.SYNOPSIS

.DESCRIPTION

.NOTES
  Version:        2.0
  Author:         Morten Rønborg (mr@endpointadmin.com)
#>

#Variables
#Early Adopters
[string]$Ring1UserGroupName = "Sec-MEM-Ring1-Users"  # Supports nested groups (excluding device objects)
[string]$PrefixGroupRing1 = "Sec-MEM-Rollout-Ring1-" # Prefix of ring groups
[int]$NumberOfGroupsRing1 = 2 # Number of groups in this ring which users will be spread equally on
[string]$SuffixGroupRing1 = "-Users"

#Early verification
[string]$Ring2UserGroupName = "Sec-MEM-Ring2-Users" # Supports nested groups (excluding device objects)
[string]$PrefixGroupRing2 = "Sec-MEM-Rollout-Ring2-"
[int]$NumberOfGroupsRing2 = 4 # Number of groups in this ring which users will be spread equally on
[string]$SuffixGroupRing2 = "-Users"

#Early production
[string]$Ring3UserGroupName = "Sec-MEM-Ring3-Users" # Supports nested groups (excluding device objects)
[string]$PrefixGroupRing3 = "Sec-MEM-Rollout-Ring3-"
[int]$NumberOfGroupsRing3 = 4 # Number of groups in this ring which users will be spread equally on
[string]$SuffixGroupRing3 = "-Users"

#Global production
[string]$PrefixGroupRing4 = "Sec-MEM-Rollout-Ring4-"
[int]$NumberOfGroupsRing4 = 8 # Number of groups in this ring which users will be spread equally on
[string]$SuffixGroupRing4 = "-Users"

#Global excluded user
[string]$GroupExcludedUsersName = "Sec-AutoRunbook-MEMUserRingsExcluded" # Supports nested groups (excluding device objects)

#Webhook notifictation
$EnableTeamsNotification = $false
$JobName = "User Ring Management"
$PictureBase64 = "data:image/jpeg;base64,<STRINGVALUE>"
$WebHookUrl = ""

#Start
$StartTime = Get-Date

#region begin functions
############### Functions - Start ###############
function Send-Teams {
    param (
        $Title = "Runbook status message",
        $StatusText = "Ring management",
        $JobName = "JobName",
        $JobTitle = "JobTitle",
        $Ring1Text = "N/A",
        $Ring2Text = "N/A",
        $Ring3Text = "N/A",
        $Ring4Text = "N/A",
        $ExcludedText = "N/A",
        $IncludedText = "N/A",
        $URL,
        $Image = "None"
    )

    $body = ConvertTo-Json -Depth 4 @{
        title = $Title
        text = $StatusText
        sections = @(
            @{
                activityTitle =  $JobName
                activitySubtitle = $JobTitle
                activityText = $StatusText
                activityImage = $Image 
                
            },
            @{
                title = 'Details'
                facts = @(
                    @{
                    name = 'Ring 1 :'
                    value = $Ring1Text
                    },
                    @{
                    name = 'Ring 2 :'
                    value = $Ring2Text
                    },
                    @{
                    name = 'Ring 3 :'
                    value = $Ring3Text
                    },
                    @{
                    name = 'Ring 4 :'
                    value = $Ring4Text
                    },
                    @{
                    name = 'Total excluded objects :'
                    value = $ExcludedText
                    },
                    @{
                    name = 'Total included objects:'
                    value = $IncludedText
                    }
                )
            }
        )
        potentialAction = @(@{
            '@context' = 'http://schema.org'
            '@type' = 'ViewAction'
            name = 'Click here to go to the Azure portal'
            target = @("https://portal.azure.com")
        })
    }

    #Invoke rest method
    $Response = Invoke-RestMethod -uri $URL -Method Post -body $body -ContentType 'application/json'
}
function Split-Array {
    param (
        [array]$InArray,
        [int]$Parts,
        [int]$Size
    )
  
    #In case the objects are less than the parts
    If($InArray.Count -le $Parts){

        $Parts = $InArray.Count
    }

    #Define parts or size
    if($Parts){
        $PartSize = [Math]::Ceiling($InArray.Count / $Parts)
    }
    if($Size){
        $PartSize = $Size
        $Parts = [Math]::Ceiling($InArray.Count / $Size)
    }

    #Define list object array
    $OutArray = New-Object 'System.Collections.Generic.List[psobject]'

    #Run through all parts
    for ($i=1; $i -le $Parts; $i++) {

        #Define start and end index
        $Start = (($i-1)*$PartSize)
        $End = (($i)*$PartSize) - 1
        if($End -ge $InArray.Count){
            $End = $InArray.Count -1
        }

        #Add to list object array
        $OutArray.Add(@($InArray[$Start..$End]))
    }

    #Return output
    Return ,$OutArray
}
#get azure ad group from graph
function Get-AzureADGroup {
    param (
        [Parameter(Mandatory=$true)]$AuthHeader,
        [Parameter(Mandatory=$true)]$Search,
        [Parameter(Mandatory=$false)]$APIVersion = "v1.0" # v1.0 or beta   
    )

    #Create request headers.
    $Headers = $AuthHeader
    $Headers["ConsistencyLevel"] = "eventual"
    $Headers["content-type"] = "application/json"

    #Do the call
    $Group = Invoke-RestMethod -Method Get -Headers $Headers -Uri "https://graph.microsoft.com/$APIVersion/groups?`$search=$Search" -ContentType "application/json"

    #Return reponse
    return [array]$Group.value
}
#get azure ad user owned device from graph
function Get-AzureADUserOwnedDevice {
    param (
        [Parameter(Mandatory=$true)]$AuthHeader,
        [Parameter(Mandatory=$true)]$Id,
        [Parameter(Mandatory=$false)]$APIVersion = "v1.0" # v1.0 or beta   
    )

    #Create request headers.
    $Headers = $AuthHeader
    $Headers["content-type"] = "application/json"

    #Do the call
    $Devices = Invoke-RestMethod -Method Get -Headers $Headers -Uri "https://graph.microsoft.com/$APIVersion/users/$Id/ownedDevices" -ContentType "application/json"

    #Return reponse
    return [array]$Devices.value
}
#get azure ad users from graph
function Get-AzureADUsers {
    param (
        [Parameter(Mandatory=$true)]$AuthHeader,
        [Parameter(Mandatory=$false)]$APIVersion = "v1.0" # v1.0 or beta
    )

    #Create request headers.
    $Headers = $AuthHeader
    $Headers["content-type"] = "application/json"

    #Create application in Intune.
    $Response = Invoke-RestMethod -Method Get -Headers $Headers -Uri "https://graph.microsoft.com/$APIVersion/users" -ContentType "application/json"

    #In case the list is longer than 100 items
    while ($Response."@odata.nextLink") {
        
        #Add members and do call
        $Members += $Response.value
        $Response = Invoke-RestMethod -Method Get -Headers $Headers -Uri $Response."@odata.nextLink" -ContentType "application/json"
    }

    #Members
    $Members += $Response.value

    #Return reponse
    return [array]$Members
}
#add azure ad group member
function Add-AzureADGroupMember {
    param (
        [Parameter(Mandatory=$true)]$AuthHeader,
        [Parameter(Mandatory=$true)]$GroupID,
        [Parameter(Mandatory=$false)]$MemberID,
        [Parameter(Mandatory=$false)]$APIVersion = "v1.0" # v1.0 or beta
    )

    #Create request headers.
    $Headers = $AuthHeader
    $Headers["content-type"] = "application/json"

    #Split into array of 20 (limit by API)
    $MemberGroups = Split-Array -InArray $MemberID -Size 20
    
    #Graph body
    foreach($MemberGroup in $MemberGroups){

        #Define constants
        $Body = @{}
        $Body['members@odata.bind'] = @()

        #Add each ID
        foreach($id in $MemberGroup){
            $Body['members@odata.bind'] += "https://graph.microsoft.com/$APIVersion/directoryObjects/$id"
        }

        #Convert body to JSON
        $Json = $Body | ConvertTo-Json

        #Do the call
        $Response = Invoke-RestMethod -Method Patch -Headers $Headers -Body $json -Uri "https://graph.microsoft.com/$APIVersion/groups/$GroupID" -ContentType "application/json"
    }
}
#remove azure ad group member
function Remove-AzureADGroupMember {
    param (
        [Parameter(Mandatory=$true)]$AuthHeader,
        [Parameter(Mandatory=$true)]$GroupID,
        [Parameter(Mandatory=$true)]$MemberID,
        [Parameter(Mandatory=$false)]$APIVersion = "v1.0" # v1.0 or beta
    )

    #Create request headers
    $Headers = $AuthHeader
    $Headers["content-type"] = "application/json"

    #Do the call
    $Response = Invoke-RestMethod -Method Delete -Headers $Headers -Uri "https://graph.microsoft.com/$APIVersion/groups/$GroupID/members/$MemberID/`$ref" -ContentType "application/json"
}
#get azure ad group members
function Get-AzureADGroupMembers {
    param (
        [Parameter(Mandatory=$true)]$AuthHeader,
        [Parameter(Mandatory=$true)]$GroupID,
        [Parameter(Mandatory=$false)]$APIVersion = "v1.0" # v1.0 or beta
    )

    #Create request headers.
    $Headers = $AuthHeader
    $Headers["content-type"] = "application/json"
    $Members = @()

    #Do the call
    $Response = Invoke-RestMethod -Method Get -Headers $Headers -Uri "https://graph.microsoft.com/$APIVersion/groups/$GroupID/members?`$select=id,displayName,description" -ContentType "application/json"

    #In case the list is longer than 100 items
    while ($Response."@odata.nextLink") {
        
        foreach ($ValueObject in $Response.value) {

            #Add members and do call
            $Obj = [PSCustomObject]@{
                '@odata.type' = $ValueObject."@odata.type"
                id = $ValueObject.Id
                displayName =  $ValueObject.displayName
                groupId =  $GroupID
            }
            $Members += $Obj
        }

        $Response = Invoke-RestMethod -Method Get -Headers $Headers -Uri $Response."@odata.nextLink" -ContentType "application/json"
    }

    #Add members
    foreach ($ValueObject in $Response.value) {

        #Add members and do call
        $Obj = [PSCustomObject]@{
            '@odata.type' = $ValueObject."@odata.type"
            id = $ValueObject.Id
            displayName =  $ValueObject.displayName
            groupId =  $GroupID
        }
        $Members += $Obj
    }

    #Return reponse
    return [array]$Members
}
#get azure ad group members in nested groups
function Get-AzureADNestedGroupObjects {
    Param
    (
        [Parameter(Mandatory=$true)]$AuthHeader,
        [Parameter(Mandatory=$true)]$GroupObj
    )

    #Get the AD object, and get group membership
    $Members = Get-AzureADGroupMembers -AuthHeader $AuthHeader -GroupID $GroupObj.id

    #Foreach member in the group.
    Foreach($Member in $Members)
    {

        #If the member is a group.
        If($Member."@odata.type" -eq "#microsoft.graph.group")
        {
            #Run the function again against the group
            $Objects += Get-AzureADNestedGroupObjects -AuthHeader $AuthHeader -GroupObj $Member
        }
        Else
        {

            #Add the object to the object array
            If(!($Member.id -in $Objects.id)){

                #Add to the array
                $Objects += @($Member)
            }
        }
    }

    #Return the users (in case object belongs to multiple nested groups, get unique)
    Return ($Objects | Sort-Object -Property id -Unique)
}
#create new azure ad groups
function New-AzureADGroup {
    param (
        [Parameter(Mandatory=$true)]$AuthHeader,
        [Parameter(Mandatory=$true)]$DisplayName,
        [Parameter(Mandatory=$true)]$Description,
        [Parameter(Mandatory=$false)]$APIVersion = "v1.0" # v1.0 or beta
    )
    
    #Graph connection strings.
    $Body = @{
        "displayName" = $DisplayName
        "mailEnabled" = $false
        "mailNickname" = $DisplayName
        "securityEnabled" = $true
        "description" = $Description
    } | ConvertTo-Json

    #Create request headers
    $Headers = $AuthHeader
    $Headers["content-type"] = "application/json"

    #Create group
    $Response = Invoke-RestMethod -Method Post -Headers $Headers -Body $Body -Uri "https://graph.microsoft.com/$APIVersion/groups" -ContentType "application/json"

    #Return object
    return $Response
}
#get or create azure ad group
function Get-CreateOrGetAzureADGroup {
    param (
        [Parameter(Mandatory=$true)]$AuthHeader,
        [Parameter(Mandatory=$true)]$DisplayName,
        [Parameter(Mandatory=$true)]$Description
    )

    #Get group
    $Group = Get-AzureADGroup -AuthHeader $AuthHeader -Search ("`"description:{0}`" AND `"displayName:{1}`"" -f $Description,$DisplayName)

    #Create if not there
    if([string]::IsNullOrEmpty($Group)){

        #Create the group
        Write-Output ("[Get-CreateOrGetAzureADGroup]::The group '{0}' does not exist. Creating it..." -f $DisplayName)
        $Group = New-AzureADGroup -AuthHeader $AuthHeader -DisplayName $DisplayName -Description $Description
        
        #Add a timeout for  the API to do changes in the Graph database, otherwise adding members will not work in some cases as the
        #group object is still not present in the backend for the add group members function
        Start-Sleep 30
    }
    
    #Return group
    return $Group
}
#invoke alligment
function Invoke-RingGroupsMembershipAlligment {
    param (
        $GroupPrefix,
        $GrpoupSuffix,
        $NumberOfGroups,
        $GroupMembers,
        $AuthHeader
    )

    #Write output
    Write-Output "[Invoke-RingGroupsMembershipAlligment]::Starting ring memebership alligment..."
    
    #Get all members of all ring subgroups
    for ($i = 0; $i -lt $NumberOfGroups; $i++) {

        #Define variables
        [string]$GroupName = ($GroupPrefix + ($i + 1).ToString().PadLeft(4,"0") + $GrpoupSuffix)
        $Group = Get-CreateOrGetAzureADGroup -AuthHeader $AuthHeader -DisplayName $GroupName -Description "Do not change the name or description of this group. This group is maintained by a runbook"
        [array]$AllGroups += $Group

        #Write output
        Write-Output "[Invoke-RingGroupsMembershipAlligment]::Fetching members of the group '$GroupName'"

        #Get all members of the groups
        [array]$CurrentRingMembers += (Get-AzureADGroupMembers -AuthHeader $AuthHeader -GroupID $Group.id)
    }

    #Remove the members that is not supposed to be there
    [array]$MembersToRemove = $CurrentRingMembers | Where-Object{$_.Id -notin $GroupMembers.Id}
    foreach ($Member in $MembersToRemove) {
        
        #Remove member
        Write-Output "[Invoke-RingGroupsMembershipAlligment]::Removing member '$($Member.id)' from the group '$($Member.groupId)'"
        Remove-AzureADGroupMember -AuthHeader $AuthHeader -GroupID $Member.groupId -MemberId $Member.id
        $CurrentRingMembers = $CurrentRingMembers | Where-Object{$_.id -notin $Member.id}
    }

    #Define members to add, how many in each group
    [array]$AllMembersToAdd = $GroupMembers | Where-Object{$_.Id -notin $CurrentRingMembers.Id}
    $MaxGroupMemberships = [Math]::Ceiling(($AllMembersToAdd.Count + $CurrentRingMembers.Count) / $NumberOfGroups)
    $CurrentRingMembersGrouped = $CurrentRingMembers | Group-Object -Property groupId | Sort-Object -Property Count
    
    #Write output
    Write-Output "[Invoke-RingGroupsMembershipAlligment]::Maximum group members in each sub group in this ring '$($MaxGroupMemberships)'"
    Write-Output "[Invoke-RingGroupsMembershipAlligment]::Total group members in this ring '$($CurrentRingMembers.count)'"
    Write-Output "[Invoke-RingGroupsMembershipAlligment]::Total group members to add in this ring '$($AllMembersToAdd.count)'"

    #First go through all empty groups
    foreach ($Group in ($AllGroups | Where-Object {$_.id -notin $CurrentRingMembers.groupId})) {
        
        if($AllMembersToAdd.Count -gt 0){
            #Define variables
            $MembersToAdd = $AllMembersToAdd[0..($MaxGroupMemberships - 1)]

            #Add objects (adding allows to add in bulks 2021-10-13)
            Write-Output "[Invoke-RingGroupsMembershipAlligment]::Adding '$($MembersToAdd.count)' objects to the group '$($Group.displayName)'"
            Add-AzureADGroupMember -AuthHeader $AuthHeader -GroupID $Group.id -MemberId $MembersToAdd.Id

            #Remove from objects to add
            $AllMembersToAdd = $AllMembersToAdd | Where-Object{$_.id -notin $MembersToAdd.id}
        }
    }

    #Go through existing groups
    foreach ($Group in $CurrentRingMembersGrouped) {
        
        if($AllMembersToAdd.Count -gt 0){

            #Define variables
            $NeededMembersInGroup = ($MaxGroupMemberships - $Group.Count)
            $MembersToAdd = $AllMembersToAdd[0..($NeededMembersInGroup - 1)]
            Write-Output "[Invoke-RingGroupsMembershipAlligment]::Adding '$($MembersToAdd.count)' objects to the group '$($Group.Name)'"

            #Add objects (adding allows to add in bulks 2021-10-13)
            Add-AzureADGroupMember -AuthHeader $AuthHeader -GroupID $Group.Name -MemberId $MembersToAdd.Id

            #Remove from objects to add
            $AllMembersToAdd = $AllMembersToAdd | Where-Object{$_.id -notin $MembersToAdd.id}
        }
    }
}
############### Functions - End ###############
#endregion
#region begin main
############### Main - Start ###############
try {

    #Obtain AccessToken for Microsoft Graph via the managed identity
    $ResourceURL = "https://graph.microsoft.com/" 
    $Response = [System.Text.Encoding]::Default.GetString((Invoke-WebRequest -UseBasicParsing -Uri "$($env:IDENTITY_ENDPOINT)?resource=$resourceURL" -Method 'GET' -Headers @{'X-IDENTITY-HEADER' = "$env:IDENTITY_HEADER"; 'Metadata' = 'True'}).RawContentStream.ToArray()) | ConvertFrom-Json 

    #Construct AuthHeader
    $AuthHeader = @{
        'Content-Type' = 'application/json'
        'Authorization' = "Bearer " + $Response.access_token
    }

}
catch {

    #Exit
    if($EnableTeamsNotification){
        Send-Teams -JobName $JobName -JobTitle "Failed" -StatusText ("Execution failed with: {0}" -f $_) -URL $WebHookUrl -Image $PictureBase64
    }
    Throw $_
}
try {

    #Get/Create user groups
    $Ring1UserGroup = Get-CreateOrGetAzureADGroup -AuthHeader $AuthHeader -DisplayName $Ring1UserGroupName -Description "Do not change the name or description of this group. This group contains users for Ring 1"
    $Ring2UserGroup = Get-CreateOrGetAzureADGroup -AuthHeader $AuthHeader -DisplayName $Ring2UserGroupName -Description "Do not change the name or description of this group. This group contains users for Ring 2"
    $Ring3UserGroup = Get-CreateOrGetAzureADGroup -AuthHeader $AuthHeader -DisplayName $Ring3UserGroupName -Description "Do not change the name or description of this group. This group contains users for Ring 3"
    $GroupExcludedUsers = Get-CreateOrGetAzureADGroup -AuthHeader $AuthHeader -DisplayName $GroupExcludedUsersName -Description "Do not change the name or description of this group. This group contains users excluded from the rings. Nested groups are supported. User objects will be ignored."

    #Get all users
    [array]$AllExcludedUsers = Get-AzureADNestedGroupObjects -AuthHeader $AuthHeader -GroupObj $GroupExcludedUsers | Where-Object {$_."@odata.type" -eq "#microsoft.graph.user"}
    [array]$AllSupportedUsers = Get-AzureADUsers -AuthHeader $AuthHeader | Where-Object {($_.Id -notin $AllExcludedUsers.Id)}

    #Get group users
    [array]$Ring1GroupUsers = Get-AzureADNestedGroupObjects -AuthHeader $AuthHeader -GroupObj $Ring1UserGroup | Where-Object {$_."@odata.type" -eq "#microsoft.graph.user"}
    [array]$Ring2GroupUsers = Get-AzureADNestedGroupObjects -AuthHeader $AuthHeader -GroupObj $Ring2UserGroup | Where-Object {$_."@odata.type" -eq "#microsoft.graph.user"}
    [array]$Ring3GroupUsers = Get-AzureADNestedGroupObjects -AuthHeader $AuthHeader -GroupObj $Ring3UserGroup | Where-Object {$_."@odata.type" -eq "#microsoft.graph.user"}
}
catch {

    #Exit
    if($EnableTeamsNotification){
        Send-Teams -JobName $JobName -JobTitle "Failed" -StatusText ("Execution failed with: {0}" -f $_) -URL $WebHookUrl -Image $PictureBase64
    }
    Throw $_
}


################################
#Running define user groupings
################################

#Add the amout of users to each major ring
[array]$AllRing1Users = $Ring1GroupUsers | Where-Object {($_.Id -in $AllSupportedUsers.Id)} |Sort-Object -Property Id -Unique
[array]$AllRing2Users = $Ring2GroupUsers | Where-Object {($_.Id -notin $AllRing1Users.Id) -and ($_.Id -in $AllSupportedUsers.Id)} | Sort-Object -Property Id -Unique
[array]$AllRing3Users = $Ring3GroupUsers | Where-Object {($_.Id -notin $AllRing1Users.Id) -and ($_.Id -notin $AllRing2Users.Id) -and ($_.Id -in $AllSupportedUsers.Id)} | Sort-Object -Property Id -Unique
[array]$AllRing4Users= $AllSupportedUsers | Where-Object {($_.Id -notin $AllRing1Users.Id) -and ($_.Id -notin $AllRing2Users.Id) -and ($_.Id -notin $AllRing3Users.Id)} | Sort-Object -Property Id -Unique

#Write statistics to output
$Ring1Text = "$($AllRing1Users.Count) ($($Ring1GroupUsers.Count) from $Ring1UserGroupName)"
$Ring2Text = "$($AllRing2Users.Count) ($($Ring2GroupUsers.Count) from $Ring2UserGroupName)"
$Ring3Text = "$($AllRing3Users.Count) ($($Ring3GroupUsers.Count) from $Ring3UserGroupName)"
$Ring4Text = "$($AllRing4Users.Count)"
@"
[Rings]::User statistics:
Ring1: $Ring1Text
Ring2: $Ring2Text
Ring3: $Ring3Text
Ring4: $Ring4Text
Total excluded users: $($AllExcludedUsers.Count)
Total included users: $($AllSupportedUsers.Count)
"@ | Write-Output

try {

    #Invoke alligment
    Invoke-RingGroupsMembershipAlligment -GroupPrefix $PrefixGroupRing1 -GrpoupSuffix $SuffixGroupRing1 -NumberOfGroups $NumberOfGroupsRing1 -GroupMembers $AllRing1Users -AuthHeader $AuthHeader 
    Invoke-RingGroupsMembershipAlligment -GroupPrefix $PrefixGroupRing2 -GrpoupSuffix $SuffixGroupRing2 -NumberOfGroups $NumberOfGroupsRing2 -GroupMembers $AllRing2Users -AuthHeader $AuthHeader 
    Invoke-RingGroupsMembershipAlligment -GroupPrefix $PrefixGroupRing3 -GrpoupSuffix $SuffixGroupRing3 -NumberOfGroups $NumberOfGroupsRing3 -GroupMembers $AllRing3Users -AuthHeader $AuthHeader 
    Invoke-RingGroupsMembershipAlligment -GroupPrefix $PrefixGroupRing4 -GrpoupSuffix $SuffixGroupRing4 -NumberOfGroups $NumberOfGroupsRing4 -GroupMembers $AllRing4Users -AuthHeader $AuthHeader 
}
catch {

    #Exit
    if($EnableTeamsNotification){
        Send-Teams -JobName $JobName -JobTitle "Failed" -StatusText ("Execution failed with: {0}" -f $_) -URL $WebHookUrl -Image $PictureBase64
    }
    Throw $_
}

#Completion
$CompletionText = ("Script completed in {0}" -f (New-TimeSpan -Start $StartTime -End (Get-Date)).ToString("dd' days 'hh' hours 'mm' minutes 'ss' seconds'"))
Write-Output $CompletionText
if ($EnableTeamsNotification) {
    Send-Teams -JobName $JobName -JobTitle "Completed" -StatusText $CompletionText -Ring1Text $Ring1Text -Ring2Text $Ring2Text -Ring3Text $Ring3Text -Ring4Text $Ring4Text -ExcludedText $($AllExcludedUsers.Count) -IncludedText $($AllSupportedUsers.Count) -URL $WebHookUrl -Image $PictureBase64
}
############### Main - End ###############
#endregion