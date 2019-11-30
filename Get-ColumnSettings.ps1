<#     
説明
	指定リストの以下の列情報をCSV出力します

例
　.\Get-FieldSettings.ps1 -WebUrl https://[tenant].sharepoint.com/sites/xxxxx -ListTitle "リスト名"
#>


# パラメータ定義
[CmdletBinding()]
Param(
	[Parameter(Mandatory=$true)]
	[string]$WebUrl,
	[Parameter(Mandatory=$true)]
	[string]$ListTitle
)

Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

$ErrorActionPreference = "Stop"

# 接続するユーザー名
$global:cstrUserName = "ここにユーザー名"

# 出力ファイルパス
$global:cstrOutputFilePath = ".\out.csv"


# メイン処理
function Start-MailProcessing ($objCredentials) {

    # パスワード入力
    $strPassword = Read-Host -Prompt "Enter password" -AsSecureString

    $objCred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($global:cstrUserName, $strPassword)

    $objCtx = New-Object Microsoft.SharePoint.Client.ClientContext($WebUrl);
    $objCtx.Credentials = $objCred;

    $objSite = $objCtx.Site;
    $objCtx.Load($objSite);
    $objCtx.ExecuteQuery();

    $objWeb = $objCtx.Web
    $objCtx.Load($objWeb)

    $objList = $objCtx.Web.Lists.GetByTitle($ListTitle)
    $fields = $objList.Fields
    $objCtx.Load($fields)
    $objCtx.ExecuteQuery()

    $ret = $fields | where { ($_.FromBaseType -eq $false) -or ($_.InternalName -eq "Title")} `
        | select InternalName,Title,TypeDisplayName,DefaultValue,Choices,ChoiceList,Required,Description

    foreach($row in $ret) {
        if ($row.Choices -ne $null) {
            foreach($ch in $row.Choices) {
                $row.ChoiceList += $ch + ","
            }
            if ($row.ChoiceList.EndsWith(",")) {
                $row.ChoiceList = $row.ChoiceList.Substring(0, $row.ChoiceList.Length-1)
            }
        }
    }

    $ret `
        | select InternalName,Title,TypeDisplayName,DefaultValue,ChoiceList,Required,Description `
        | Export-Csv -Path $global:cstrOutputFilePath -Encoding UTF8 -NoTypeInformation
}

# エントリーポイント
Start-MailProcessing
