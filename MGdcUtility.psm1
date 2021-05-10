function Get-MGdcEstimatedNumberOfItems
{
    [OutputType([System.Collections.Hashtable])]
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true)]
        [System.String]
        $AppId,
        
        [parameter(Mandatory=$true)]
        [System.String]
        $TenantId,
        
        [parameter(Mandatory=$true)]
        [System.String]
        $Secret,
        
        [parameter(Mandatory = $true)]
        [System.String]
        [ValidateSet('Messages')]
        $Entity,

        [parameter()]
        [System.UInt32]
        $NumberOfDays = 7,

        [parameter()]
        [System.String[]]
        $GroupsID
    )
    $url = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{
        scope = "https://graph.microsoft.com/.default"
        grant_type = "client_credentials"
        client_secret = $Secret
        client_info = 1
        client_id = $AppId
    }
    Write-Verbose -Message "Requesting Access Token for Microsoft Graph"
    $OAuthReq = Invoke-RestMethod -Uri $url -Method Post -Body $body
    $AccessToken = $OAuthReq.access_token

    Write-Verbose -Message "Connecting to Microsoft Graph"
    Connect-MgGraph -AccessToken $AccessToken | Out-Null

    Write-Verbose "Obtaining report file for {$Entity}"
    $tempReportPath = Join-Path -Path $env:TEMP -ChildPath "graphreport.csv"
    Remove-Item $tempReportPath -Force -ErrorAction SilentlyContinue
    New-Item -Path $tempReportPath -Force | Out-Null

    #region Determine Period's Length
    $period = "D180"
    if ($NumberOfDays -le 7)
    {
        $period = "D7"
    }
    elseif ($NumberOfDays -gt 7 -and $NumberOfDays -le 30)
    {
        $period = "D30"
    }
    elseif ($NumberOfDays -gt 30 -and $NumberOfDays -le 90)
    {
        $period = "D90"
    }
    #endregion

    #region Determine users in specified groups
    [array]$users = @()
    foreach ($groupID in $GroupsID)
    {
        $groupMembers = Get-MgGroupMember -GroupId $groupId

        foreach ($member in $groupMembers)
        {
            $user = Get-MgUser -UserId $member.Id
            $users += $user.UserPrincipalName
        }
    }
    #endregion

    switch($entity)
    {
        "Messages"
        {
            Write-Verbose -Message "Retrieving emails based on {$period}"
            $url = "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='$period')"
            Invoke-MgGraphRequest -Uri $url `
                -OutputFilePath $tempReportPath | Out-Null
            $parsedReport = Import-Csv $tempReportPath

            $InboxEmailCount   = 0
            $DeletedEmailCount = 0

            # Count Emails for all users
            foreach ($user in $parsedReport)
            {
                if (($users.Length -gt 0 -and $users.Contains($user.'User Principal Name')) -or $users.Length -eq 0)
                {
                    $InboxEmailCount += $user.'Item Count'
                    $DeletedEmailCount += $user.'Deleted Item Count'
                }
            }

            $results = @{
                EmailsInInbox = $InboxEmailCount
                DeletedEmails = $DeletedEmailCount
                TotalItems = $InboxEmailCount + $DeletedEmailCount
            }
        }
    }
    return $results
}

function Get-MGdcEstimatedNumberOfItems2
{
    [OutputType([System.Collections.Hashtable])]
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true)]
        [System.String]
        $AppId,
        
        [parameter(Mandatory=$true)]
        [System.String]
        $TenantId,
        
        [parameter(Mandatory=$true)]
        [System.String]
        $Secret,
        
        [parameter(Mandatory = $true)]
        [System.String]
        [ValidateSet('Messages')]
        $Entity,

        [parameter()]
        [System.UInt32]
        $NumberOfDays = 7,

        [parameter()]
        [System.String[]]
        $GroupsID
    )
    $url = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{
        scope = "https://graph.microsoft.com/.default"
        grant_type = "client_credentials"
        client_secret = $Secret
        client_info = 1
        client_id = $AppId
    }
    Write-Verbose -Message "Requesting Access Token for Microsoft Graph"
    $OAuthReq = Invoke-RestMethod -Uri $url -Method Post -Body $body
    $AccessToken = $OAuthReq.access_token

    Write-Verbose -Message "Connecting to Microsoft Graph"
    Connect-MgGraph -AccessToken $AccessToken | Out-Null

    [array]$allUsers = Get-MgUser -All:$true

    $foldersToFilterOut = @("Conversation History", "Sent Items", "Deleted Items")

    $total = 0
    $index = 0
    foreach ($user in $allUsers)
    {
        $mailFolders = Get-MgUserMailFolder -UserId $user.Id -ErrorAction SilentlyContinue

        if ($null -ne $mailFolders)
        {
            foreach ($folder in $mailFolders)
            {
                if (-not $foldersToFilterOut.Contains($folder.DisplayName))
                {
                    Write-Progress -Activity "{$($user.UserPrincipalName)} -- {$($folder.DisplayName)}" -Status "Scanning" -PercentComplete ($index/$allUSers.Length*100)
                    Write-Verbose -Message "{$($user.UserPrincipalName)}::{$($folder.DisplayName)}---> $($folder.TotalItemCount)"
                    $total += $folder.TotalItemCount
                }
            }
        }
        $index++
    }

    return $total
}