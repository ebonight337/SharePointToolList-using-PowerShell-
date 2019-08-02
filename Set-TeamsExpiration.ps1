########################################################
#   Description: Teamsのチームに有効期限ポリシーを適用する
#   Title: Set-TeamsExpiration.ps1
#   Created: 2019/08/02
#   Created by: Sendy
#   Ver: 1.0
########################################################
Param(
    [Parameter(Mandatory=$true)] [string] $TargetTeamsName
)

Function SetTeamsExpiration(){
    Param(
        [Parameter(Mandatory=$true)] [string] $TargetTeamsName
    )

    $Credential = Get-Credential

    # Exchangeに接続
    try{
        Write-Host "Exchangeへ接続を行います。" -BackgroundColor Green -ForegroundColor Black
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection
        $SessionId = Import-PSSession $Session -DisableNameChecking
    }catch{
        $ErrorMessage = $_.Exception.Message
        throw "Failed connect to MicrosoftExchange. `r`nPlease check the following error. `r`n $($ErrorMessage)"
    }

    # Graph APIでADに接続
    Install-Module -Name AzureAD
    Connect-AzureAD -Credential $Credential

    # 有効期限ポリシー適用処理
    Write-Host "有効期限ポリシー適用処理を行います。" -BackgroundColor Green -ForegroundColor Black
    $PolicyId = (Get-AzureADMSGroupLifecyclePolicy).Id
    $TargetTeam = Get-UnifiedGroup -ResultSize Unlimited | where {$_.DisplayName -eq $TargetTeamsName} 
    Add-AzureADMSLifecyclePolicyGroup -GroupId $TargetTeam.ExternalDirectoryObjectId -Id $PolicyId -ErrorAction SilentlyContinue

    # セッション削除
    Get-PSSession | %{Remove-PSSession $_.Id}
    Disconnect-AzureAD

}

SetTeamsExpiration -TargetTeamsName $TargetTeamsName
