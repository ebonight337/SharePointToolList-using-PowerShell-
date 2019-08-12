###################################################
# Teamsに紐づくチームサイトを取得します。
# https://blogs.technet.microsoft.com/teamsjp/2018/02/28/get-teamlist/
###################################################
Function GetTeamsSite(){
    $Credential = Get-Credential
    # Teamsに接続
    try{
        Write-Host "Teamsへ接続を行います。" -BackgroundColor Green -ForegroundColor Black
        Install-Module -Name MicrosoftTeams
        Connect-MicrosoftTeams -Credential $Credential
    }catch{
        $ErrorMessage = $_.Exception.Message
        throw "Failed Connect-Microsoft Teams command. `r`nPlease check the following error. `r`n $($ErrorMessage)"
    }

    # Exchangeに接続
    try{
        Write-Host "Exchangeへ接続を行います。" -BackgroundColor Green -ForegroundColor Black
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection
        $SessionId = Import-PSSession $Session -DisableNameChecking
    }catch{
        $ErrorMessage = $_.Exception.Message
        throw "Failed connect to MicrosoftExchange. `r`nPlease check the following error. `r`n $($ErrorMessage)"
    }

    # チームサイト取得
    # Get-Teamでは権限がないチームの情報が取得できない
    Write-Host "チームサイトの情報を取得します。" -BackgroundColor Green -ForegroundColor Black
    $o365groups = Get-UnifiedGroup -ResultSize Unlimited
    $TeamsList = @()

    foreach ($o365group in $o365groups) 
    {
        try{
            $teamschannels = Get-TeamChannel -GroupId $o365group.ExternalDirectoryObjectId
            $o365GroupMemberList = (Get-UnifiedGroupLinks -identity $o365group.ExternalDirectoryObjectId -LinkType Members) | select -expandproperty PrimarySmtpAddress
            $TeamsList = $TeamsList + [pscustomobject]@{GroupId = $o365group.ExternalDirectoryObjectId; GroupName = $o365group.DisplayName; `
            SMTP = $o365group.PrimarySmtpAddress; Members = $o365group.GroupMemberCount; MemberList = $o365GroupMemberList -join ', '; `
            SPOSite = $o365group.SharePointSiteUrl; TeamsEnabled = $true; Owners = $o365group.ManagedBy}
        } 
        catch{
            $ErrorCode = $_.Exception.ErrorCode
            switch ($ErrorCode) 
            {
                "404" 
                {
                    $TeamsList = $TeamsList + [pscustomobject]@{GroupId = $o365group.ExternalDirectoryObjectId; GroupName = $o365group.DisplayName; `
                    SMTP = $o365group.PrimarySmtpAddress; Members = $o365group.GroupMemberCount; MemberList = $o365GroupMemberList -join ', '; `
                    SPOSite = $o365group.SharePointSiteUrl; TeamsEnabled = $false; Owners = $o365group.ManagedBy}
                    break;
                }
                "403" 
                {
                    $TeamsList = $TeamsList + [pscustomobject]@{GroupId = $o365group.ExternalDirectoryObjectId; GroupName = $o365group.DisplayName; `
                    SMTP = $o365group.PrimarySmtpAddress; Members = $o365group.GroupMemberCount; MemberList = $o365GroupMemberList -join ', '; `
                    SPOSite = $o365group.SharePointSiteUrl; TeamsEnabled = $true; Owners = $o365group.ManagedBy}
                    break;
                }
                default 
                {
                    Write-Error ("Unknown ErrorCode trying to 'Get-TeamChannel -GroupId {0}' :: {1}" -f $o365group, $ErrorCode)
                }
            }
        }
    }
    Disconnect-MicrosoftTeams
    Get-PSSession | %{Remove-PSSession -Id $_.Id}
    if(Test-Path "c:\temp" -eq $False){
        New-Item -ItemType directory -Path c:\temp
    }
    $TeamsList | export-csv C:\temp\AllTeamsInTenant.csv -NoTypeInformation -Encoding UTF8

}

GetTeamsSite
