<#
.SYNOPSIS

.DESCRIPTION

.NOTES
  Version:        1.0
  Author:         Morten Rønborg (mr@zwable.com)
  Creation Date:  2021-10-13
  Purpose/Change: 2021-10-13
#>

#Variables
#Early Adopters
[string]$Ring1UserGroupName = "Sec-MEM-EarlyAdopters-Users"  # Supports nested groups (excluding device objects)
[string]$PrefixGroupRing1 = "Sec-AutoRunbook-MEMEarlyAdopters-" # Prefix of ring groups
[int]$NumberOfGroupsRing1 = 2 # Number of groups in this ring which devices will be spread equally on

#Early verification
[string]$Ring2UserGroupName = "Sec-MEM-EarlyVerification-Users" # Supports nested groups (excluding device objects)
[string]$PrefixGroupRing2 = "Sec-AutoRunbook-MEMEarlyVerification-"
[int]$NumberOfGroupsRing2 = 4 # Number of groups in this ring which devices will be spread equally on

#Early production
[string]$Ring3UserGroupName = "Sec-MEM-EarlyProduction-Users" # Supports nested groups (excluding device objects)
[string]$PrefixGroupRing3 = "Sec-AutoRunbook-MEMEarlyProduction-"
[int]$NumberOfGroupsRing3 = 8 # Number of groups in this ring which devices will be spread equally on

#Global production
[string]$PrefixGroupRing4 = "Sec-AutoRunbook-MEMGlobalProduction-"
[int]$NumberOfGroupsRing4 = 16 # Number of groups in this ring which devices will be spread equally on

#Global excluded user
[string]$GroupExcludedUsersName = "Sec-AutoRunbook-MEMDeviceRingsExcluded" # Supports nested groups (excluding device objects)

#Start
$StartTime = Get-Date

################################
#Variable end
################################

function Split-Array
{
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
function Get-AzureADGroup
{
    param (
        [Parameter(Mandatory=$true)]$AuthHeader,
        [Parameter(Mandatory=$true)]$Search
    )

    #Create request headers.
    $Headers = $AuthHeader
    $Headers["ConsistencyLevel"] = "eventual"
    $Headers["content-type"] = "application/json"

    #Do the call
    $Group = Invoke-RestMethod -Method Get -Headers $Headers -Uri "https://graph.microsoft.com/beta/groups?`$search=$Search" -ContentType "application/json"

    #Return reponse
    return [array]$Group.value
}
function Get-AzureADUserOwnedDevice 
{
    param (
        [Parameter(Mandatory=$true)]$AuthHeader,
        [Parameter(Mandatory=$true)]$Id
    )

    #Create request headers.
    $Headers = $AuthHeader
    $Headers["content-type"] = "application/json"

    #Do the call
    $Devices = Invoke-RestMethod -Method Get -Headers $Headers -Uri "https://graph.microsoft.com/beta/users/$Id/ownedDevices" -ContentType "application/json"

    #Return reponse
    return [array]$Devices.value
}
function Get-AzureADDevice 
{
    param (
        [Parameter(Mandatory=$true)]$AuthHeader
    )


    #Create request headers.
    $Headers = $AuthHeader
    $Headers["content-type"] = "application/json"

    #Create application in Intune.
    $Response = Invoke-RestMethod -Method Get -Headers $Headers -Uri "https://graph.microsoft.com/beta/devices?`$filter=startswith(operatingSystem, 'Windows')" -ContentType "application/json"

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
function Add-AzureADGroupMember 
{
    param (
        [Parameter(Mandatory=$true)]$AuthHeader,
        [Parameter(Mandatory=$true)]$GroupID,
        [Parameter(Mandatory=$false)]$MemberID
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
            $Body['members@odata.bind'] += "https://graph.microsoft.com/beta/directoryObjects/$id"
        }

        #Convert body to JSON
        $Json = $Body | ConvertTo-Json

        #Do the call
        $Response = Invoke-RestMethod -Method Patch -Headers $Headers -Body $json -Uri "https://graph.microsoft.com/beta/groups/$GroupID" -ContentType "application/json"
    }
}
function Remove-AzureADGroupMember 
{
    param (
        [Parameter(Mandatory=$true)]$AuthHeader,
        [Parameter(Mandatory=$true)]$GroupID,
        [Parameter(Mandatory=$true)]$MemberID
    )

    #Create request headers
    $Headers = $AuthHeader
    $Headers["content-type"] = "application/json"

    #Do the call
    $Response = Invoke-RestMethod -Method Delete -Headers $Headers -Uri "https://graph.microsoft.com/beta/groups/$GroupID/members/$MemberID/`$ref" -ContentType "application/json"
}
function Get-AzureADGroupMembers 
{
    param (
        [Parameter(Mandatory=$true)]$AuthHeader,
        [Parameter(Mandatory=$true)]$GroupID
    )

    #Create request headers.
    $Headers = $AuthHeader
    $Headers["content-type"] = "application/json"

    #Do the call
    $Response = Invoke-RestMethod -Method Get -Headers $Headers -Uri "https://graph.microsoft.com/beta/groups/$GroupID/members?`$select=id,displayName,description" -ContentType "application/json"

    #In case the list is longer than 100 items
    while ($Response."@odata.nextLink") {
        
        #Add members and do call
        $Members += $Response.value
        $Response = Invoke-RestMethod -Method Get -Headers $Headers -Uri $Response."@odata.nextLink" -ContentType "application/json"
    }

    $Members += $Response.value

    #Return reponse
    return [array]$Members
}
function Get-AzureADNestedGroupObjects
{
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
function New-AzureADGroup 
{
    param (
        [Parameter(Mandatory=$true)]$AuthHeader,
        [Parameter(Mandatory=$true)]$DisplayName,
        [Parameter(Mandatory=$true)]$Description

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
    $Response = Invoke-RestMethod -Method Post -Headers $Headers -Body $Body -Uri "https://graph.microsoft.com/beta/groups" -ContentType "application/json"

    #Return object
    return $Response
}
function Get-CreateOrGetAzureADGroup
{
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
    }
    
    #Return group
    return $Group
}
function Set-AllignGroupMemberships 
{
    param (
        [Parameter(Mandatory=$true)]$AuthHeader,
        [Parameter(Mandatory=$true)]$GroupObj,
        [Parameter(Mandatory=$false)][array]$AADObjectIDs
    )

    #State amount to 
    Write-Output "[Set-AllignGroupMemberships:$($GroupObj.DisplayName)]::Expected objects in group: $($AADObjectIDs.Count)"

    #Get current members of the group
    [array]$CurrentMembersIDs = (Get-AzureADGroupMembers -AuthHeader $AuthHeader -GroupID $GroupObj.id).id
    Write-Output "[Set-AllignGroupMemberships:$($GroupObj.DisplayName)]::Current objects in group: $($CurrentMembersIDs.Count)"

    #Declare variables
    [array]$ObjectsIDsToRemove = $CurrentMembersIDs | Where-Object{$_ -notin $AADObjectIDs}
    [array]$ObjectsIDsToAdd =  $AADObjectIDs | Where-Object {$_ -notin $CurrentMembersIDs}

    #Write host
    Write-Output "[Set-AllignGroupMemberships:$($GroupObj.DisplayName)]::Objects to add: $($ObjectsIDsToAdd.Count)"
    Write-Output "[Set-AllignGroupMemberships:$($GroupObj.DisplayName)]::Objects to remove: $($ObjectsIDsToRemove.Count)"

    #Removing objects (removing supports only one object per API call 2021-10-13)
    Foreach($ObjectID in $ObjectsIDsToRemove){
        
        #Remove member
        Remove-AzureADGroupMember -AuthHeader $AuthHeader -GroupID $GroupObj.id -MemberId $ObjectID
    }

    #Add objects (adding allows to add in bulks 2021-10-13)
    Add-AzureADGroupMember -AuthHeader $AuthHeader -GroupID $GroupObj.id -MemberId $ObjectsIDsToAdd
}
################################
#Functions start
################################

################################
#Functions end
################################

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
    Throw $_
}

#region begin main
############### Main - Start ###############

#Get all current supported PC types
[array]$AllSupportedWinDevices = Get-AzureADDevice -AuthHeader $AuthHeader

#Get/Create user groups
$Ring1UserGroup = Get-CreateOrGetAzureADGroup -AuthHeader $AuthHeader -DisplayName $Ring1UserGroupName -Description "Do not change the name or description of this group. This group contains users for Ring 1"
$Ring2UserGroup = Get-CreateOrGetAzureADGroup -AuthHeader $AuthHeader -DisplayName $Ring2UserGroupName -Description "Do not change the name or description of this group. This group contains users for Ring 2"
$Ring3UserGroup = Get-CreateOrGetAzureADGroup -AuthHeader $AuthHeader -DisplayName $Ring3UserGroupName -Description "Do not change the name or description of this group. This group contains users for Ring 3"

#Get/Create device groups
$Ring1DeviceGroup = Get-CreateOrGetAzureADGroup -AuthHeader $AuthHeader -DisplayName $Ring1DeviceGroupName -Description "Do not change the name or description of this group. This group contains devices for Ring 1"
$Ring2DeviceGroup = Get-CreateOrGetAzureADGroup -AuthHeader $AuthHeader -DisplayName $Ring2DeviceGroupName -Description "Do not change the name or description of this group. This group contains devices for Ring 2" 
$Ring3DeviceGroup = Get-CreateOrGetAzureADGroup -AuthHeader $AuthHeader -DisplayName $Ring3DeviceGroupName -Description "Do not change the name or description of this group. This group contains devices for Ring 3"
$Ring4DeviceGroup = Get-CreateOrGetAzureADGroup -AuthHeader $AuthHeader -DisplayName $Ring4DeviceGroupName -Description "Do not change the name or description of this group. This group contains devices for Ring 4"
$GroupExcludedUsers= Get-CreateOrGetAzureADGroup -AuthHeader $AuthHeader -DisplayName $GroupExcludedUsersName -Description "Do not change the name or description of this group. This group contains devices excluded from the rings. Nested groups are supported. User objects will be ignored."

#Get devices for device groups, we need to exclude these from the global scope as they need to be enforced to each ring later
[array]$Ring1GroupDevices = Get-AzureADNestedGroupObjects -AuthHeader $AuthHeader -GroupObj $Ring1DeviceGroup | Where-Object {$_."@odata.type" -eq "#microsoft.graph.user"}
[array]$Ring2GroupDevices = Get-AzureADNestedGroupObjects -AuthHeader $AuthHeader -GroupObj $Ring2DeviceGroup | Where-Object {$_."@odata.type" -eq "#microsoft.graph.user"}
[array]$Ring3GroupDevices = Get-AzureADNestedGroupObjects -AuthHeader $AuthHeader -GroupObj $Ring3DeviceGroup | Where-Object {$_."@odata.type" -eq "#microsoft.graph.user"}
[array]$Ring4GroupDevices = Get-AzureADNestedGroupObjects -AuthHeader $AuthHeader -GroupObj $Ring4DeviceGroup | Where-Object {$_."@odata.type" -eq "#microsoft.graph.user"}
[array]$AllExcludedDevices = Get-AzureADNestedGroupObjects -AuthHeader $AuthHeader -GroupObj $GroupExcludedUsers | Where-Object {$_."@odata.type" -eq "#microsoft.graph.user"}

#Remove the excluded devices from the allsupported win devcices
[array]$AllSupportedWinDevices = $AllSupportedWinDevices | Where-Object {$_.Id -notin $AllExcludedDevices.Id}

#Define all global group
[array]$GlobalWinDevices = $AllSupportedWinDevices | Where-Object {($_.Id -notin $Ring1GroupDevices.Id) `
                                                    -and ($_.Id -notin $Ring2GroupDevices.Id) `
                                                    -and ($_.Id -notin $Ring3GroupDevices.Id) `
                                                    -and ($_.Id -notin $Ring4GroupDevices.Id)}

################################
#Running define device groupings
################################

#Allign Ring1 groups with users devices (Primary Devices in Intune)
[array]$Ring1UserGroupMembers = Get-AzureADNestedGroupObjects -AuthHeader $AuthHeader -GroupObj $Ring1UserGroup | Where-Object {$_."@odata.type" -eq "#microsoft.graph.user"}
Foreach($User in $Ring1UserGroupMembers){

    #Add all users Primary Devices to an array and ensure they are part of the supported device list
    [array]$Ring1Devices += (Get-AzureADUserOwnedDevice -AuthHeader $AuthHeader  -Id $User.Id | Where-Object {$_.Id -in $GlobalWinDevices.Id}) 
}

#Allign Ring2 groups with users devices (Primary Devices in Intune)
[array]$Ring2UserGroupMembers = Get-AzureADNestedGroupObjects -AuthHeader $AuthHeader -GroupObj $Ring2UserGroup | Where-Object {$_."@odata.type" -eq "#microsoft.graph.user"}
Foreach($User in $Ring2UserGroupMembers){

    #Add all users Primary Devices to an array and ensure they are part of the supported device list (sort out devices from Ring1 as users can be in more groups)
    [array]$Ring2Devices += (Get-AzureADUserOwnedDevice -AuthHeader $AuthHeader  -Id $User.Id | Where-Object {($_.Id -in $GlobalWinDevices.Id) -and ($_.Id -notin $Ring1Devices.Id)})
}

#Allign Ring2 groups with users devices (Primary Devices in Intune)
[array]$Ring3UserGroupMembers = Get-AzureADNestedGroupObjects -AuthHeader $AuthHeader -GroupObj $Ring3UserGroup | Where-Object {$_."@odata.type" -eq "#microsoft.graph.user"}
Foreach($User in $Ring3UserGroupMembers){

    #Add all users Primary Devices to an array and ensure they are part of the supported device list (sort out devices from Ring1 and Ring2 as users can be in more groups)
    [array]$Ring3Devices += (Get-AzureADUserOwnedDevice -AuthHeader $AuthHeader  -Id $User.Id | Where-Object {($_.Id -in $GlobalWinDevices.Id) -and ($_.Id -notin $Ring1Devices.Id) -and ($_.Id -notin $Ring2Devices.Id)})
}

#Remove all the Primary User based devices from the device pool before defining major groups
[array]$AllSupportedWinDevicesNoPrimaryDevices = $GlobalWinDevices | Where-Object{($_.Id -notin $Ring1Devices.Id) `
                                                                                        -and ($_.Id -notin $Ring2Devices.Id) `
                                                                                        -and ($_.Id -notin $Ring3Devices.Id)} | Sort-Object -Property Id

#Add the amout of devices to each major ring (use unique as one device can have multiple owners)
[array]$AllRing1Devices = ($Ring1Devices + $Ring1GroupDevices) | Sort-Object -Property Id -Unique
[array]$AllRing2Devices = ($Ring2Devices + $Ring2GroupDevices) | Where-Object {($_.Id -notin $AllRing1Devices.Id)} | Sort-Object -Property Id -Unique
[array]$AllRing3Devices = ($Ring3Devices + $Ring3GroupDevices) | Where-Object {($_.Id -notin $AllRing1Devices.Id) -and ($_.Id -notin $AllRing2Devices.Id)} | Sort-Object -Property Id -Unique
[array]$AllRing4Devices = ($AllSupportedWinDevicesNoPrimaryDevices[0..$AllSupportedWinDevicesNoPrimaryDevices.Count] + $Ring4GroupDevices) | Where-Object {($_.Id -notin $AllRing1Devices.Id) -and ($_.Id -notin $AllRing2Devices.Id) -and ($_.Id -notin $AllRing3Devices.Id)} | Sort-Object -Property Id -Unique

Write-Output "[Rings]::Device statistics:`nRing1: $($AllRing1Devices.Count) ($($Ring1GroupDevices.Count) from $Ring1DeviceGroupName)`nRing2: $($AllRing2Devices.Count) ($($Ring2GroupDevices.Count) from $Ring2DeviceGroupName)`nRing3: $($AllRing3Devices.Count) ($($Ring3GroupDevices.Count) from $Ring3DeviceGroupName)`nRing4: $($AllRing4Devices.Count) ($($Ring4GroupDevices.Count) from $Ring4DeviceGroupName)`nTotal excluded devices: $($AllExcludedDevices.Count)`nTotal included devices: $($AllSupportedWinDevices.Count)"

#Split the objects into the count of groups sorting them (sort on multiple properties in case they where created on the same date/time)
[array]$Ring1Groupings = Split-Array -InArray ($AllRing1Devices | Sort-Object -Property createdDateTime,displayName) -Parts $NumberOfGroupsRing1
[array]$Ring2Groupings = Split-Array -InArray ($AllRing2Devices | Sort-Object -Property createdDateTime,displayName) -Parts $NumberOfGroupsRing2
[array]$Ring3Groupings = Split-Array -InArray ($AllRing3Devices | Sort-Object -Property createdDateTime,displayName) -Parts $NumberOfGroupsRing3
[array]$Ring4Groupings = Split-Array -InArray ($AllRing4Devices | Sort-Object -Property createdDateTime,displayName) -Parts $NumberOfGroupsRing4

################################
#Running Ring1
################################

for ($i = 0; $i -lt $NumberOfGroupsRing1; $i++) {

    #Define variables
    [array]$GroupMemberIDs = ($Ring1Groupings[$i]).Id
    [string]$GroupName = ($PrefixGroupRing1 + ($i + 1).ToString().PadLeft(4,"0") + "-Users")
    $Group = Get-CreateOrGetAzureADGroup -AuthHeader $AuthHeader -DisplayName $GroupName -Description "Do not change the name or description of this group. This group is maintained by a runbook"

    #Allign memberships
    Write-Output "[Ring1]::Alligning members of the group '$($Group.DisplayName)' with the ID of '$($Group.Id)'"
    Set-AllignGroupMemberships -AuthHeader $AuthHeader -Group $Group -AADObjectIDs $GroupMemberIDs
}

################################
#Running Ring2
################################

for ($i = 0; $i -lt $NumberOfGroupsRing2; $i++) {
    
    #Define variables
    [array]$GroupMemberIDs = ($Ring2Groupings[$i]).Id
    [string]$GroupName = ($PrefixGroupRing2 + ($i + 1).ToString().PadLeft(4,"0") + "-Users")
    $Group = Get-CreateOrGetAzureADGroup -AuthHeader $AuthHeader -DisplayName $GroupName -Description "Do not change the name or description of this group. This group is maintained by a runbook"

    #Allign memberships
    Write-Output "[Ring2]::Alligning members of the group '$($Group.DisplayName)' with the ID of '$($Group.Id)'"
    Set-AllignGroupMemberships -AuthHeader $AuthHeader -Group $Group -AADObjectIDs $GroupMemberIDs
}

################################
#Running Ring3
################################

for ($i = 0; $i -lt $NumberOfGroupsRing3; $i++) {
    
    #Define variables
    [array]$GroupMemberIDs = ($Ring3Groupings[$i]).Id
    [string]$GroupName = ($PrefixGroupRing3 + ($i + 1).ToString().PadLeft(4,"0") + "-Users")
    $Group = Get-CreateOrGetAzureADGroup -AuthHeader $AuthHeader -DisplayName $GroupName -Description "Do not change the name or description of this group. This group is maintained by a runbook"

    #Allign memberships
    Write-Output "[Ring3]::Alligning members of the group '$($Group.DisplayName)' with the ID of '$($Group.Id)'"
    Set-AllignGroupMemberships -AuthHeader $AuthHeader -Group $Group -AADObjectIDs $GroupMemberIDs
}

################################
#Running Ring4
################################

for ($i = 0; $i -lt $NumberOfGroupsRing4; $i++) {
    
    #Define variables
    [array]$GroupMemberIDs = ($Ring4Groupings[$i]).Id
    [string]$GroupName = ($PrefixGroupRing4 + ($i + 1).ToString().PadLeft(4,"0") + "-Users")
    $Group = Get-CreateOrGetAzureADGroup -AuthHeader $AuthHeader -DisplayName $GroupName -Description "Do not change the name or description of this group. This group is maintained by a runbook"

    #Allign memberships
    Write-Output "[Ring4]::Alligning members of the group '$($Group.DisplayName)' with the ID of '$($Group.Id)'"
    Set-AllignGroupMemberships -AuthHeader $AuthHeader -Group $Group -AADObjectIDs $GroupMemberIDs
}

############### Main - End ###############
#endregion

#Completion
Write-Output ("Script completed in {0}" -f (New-TimeSpan -Start $StartTime -End (Get-Date)).ToString("dd' days 'hh' hours 'mm' minutes 'ss' seconds'"))