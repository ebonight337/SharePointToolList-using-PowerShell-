###################################################
# Teamsに紐づくチームサイトを取得します。
# https://blogs.technet.microsoft.com/teamsjp/2018/02/28/get-teamlist/
###################################################
Function GetTeamsSite(){
    $Credential = Get-Credential
    $ScriptFolder = Join-Path $PSScriptRoot TeamList.csv

    # Exchangeに接続
    try{
        Write-Host "Exchangeへ接続を行います。" -BackgroundColor Green -ForegroundColor Black
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection
        $SessionId = Import-PSSession $Session -DisableNameChecking
    }catch{
        $ErrorMessage = $_.Exception.Message
        throw "MicrosoftExchangeに接続できませんでした。 `r`nエラーを確認してください。 `r`n $($ErrorMessage)"
    }

    # チームサイト取得
    # Get-Teamでは権限がないチームの情報が取得できない
    Write-Host "チームサイトの情報を取得します。" -BackgroundColor Green -ForegroundColor Black
    $o365groups = Get-UnifiedGroup -ResultSize Unlimited
    $TeamsList = @()

    foreach ($o365group in $o365groups) 
    {
        try{
            # Teamのみ取得
            if ($o365group.ResourceProvisioningOptions -eq "Team"){
                $teamsList = $teamsList + [pscustomobject]@{TeamsID = $o365group.PrimarySmtpAddress; TeamsName = $o365group.DisplayName; SharePointUrl = $o365group.SharePointSiteUrl}
            }
        } catch{
            $ErrorMessage = $_.Exception.Message
            throw "Teamが取得できませんでした。 `r`nエラーを確認してください。 `r`n $($ErrorMessage)"
        }
    }
    Get-PSSession | %{Remove-PSSession -Id $_.Id}
    $TeamsList | export-csv $ScriptFolder -NoTypeInformation -Encoding UTF8

}

GetTeamsSite
