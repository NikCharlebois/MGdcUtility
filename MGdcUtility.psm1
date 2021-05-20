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

    $allUsers = Get-MgUser -All:$true

    $foldersToFilterOut = @()#@("Conversation History", "Sent Items", "Deleted Items")

    $total = 0
    $jobName = "MGdcEstimateJob" + (New-Guid).ToString()

    $totalNumberOfUsers = $allUsers.Length
    $numberOfParallelThreads = 4
    $instances = Split-ArrayByParts -Array $allUsers `
        -Parts $numberOfParallelThreads

    foreach ($batch in $instances)
    {
        Start-Job -Name "$($jobName + $index)" -ScriptBlock {
            param(
                [System.Object[]]
                $Batch,

                [System.String]
                $AccessToken
            )            
            Connect-MgGraph -AccessToken $AccessToken | Out-Null
            $mailFolders = @()
            foreach ($user in $Batch)
            {
                $mailFolders += Get-MgUserMailFolder -UserId $User.Id -ErrorAction SilentlyContinue
            }
            return $mailFolders
        } -ArgumentList @($batch, $AccessToken) | Out-Null
    }

    do
    {
        [array]$pendingJobs = Get-Job | Where-Object -FilterScript { $_.Name -like 'MGdcEstimateJob*' -and $_.JobStateInfo.State -notin @('Complete', 'Blocked', 'Failed')}
        [array]$CompletedJobs = Get-Job | Where-Object -FilterScript { $_.Name -like 'MGdcEstimateJob*' -and $_.JobStateInfo.State -eq 'Complete'}

        foreach ($completedJob in $CompletedJobs)
        {
            $currentContent = Receive-Job -Name $completedJob.name
            $mailFolders = $currentContent
            Remove-Job -Name $completedJob.name -ErrorAction SilentlyContinue
            if ($null -ne $mailFolders)
            {
                foreach ($folder in $mailFolders)
                {
                    if (-not $foldersToFilterOut.Contains($folder.DisplayName))
                    {
                        $total += $folder.TotalItemCount
                    }
                }
            }
        }

    } while ($pendingJobs.Length -gt 0)
    
    return $total
}

function Split-ArrayByParts
{
    [OutputType([System.Object[]])]
    param(
        [Parameter(Mandatory = $true)]
        [System.Object[]]
        $Array,

        [Parameter(Mandatory = $true)]
        [System.Uint32]
        $Parts
    )

    if ($Parts)
    {
        $PartSize = [Math]::Ceiling($Array.Count / $Parts)
    }
    $outArray = New-Object 'System.Collections.Generic.List[PSObject]'

    for ($i = 1; $i -le $Parts; $i++)
    {
        $start = (($i - 1) * $PartSize)

        if ($start -lt $Array.Count)
        {
            $end = (($i) * $PartSize) - 1
            if ($end -ge $Array.count)
            {
                $end = $Array.count - 1
            }
            $outArray.Add(@($Array[$start..$end]))
        }
    }
    return , $outArray
}