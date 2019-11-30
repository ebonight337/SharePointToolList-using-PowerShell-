<#     
説明
  指定リストの以下の列情報をCSV出力します
例
　.\Get-ColumnSettings.ps1 -WebUrl https://[tenant].sharepoint.com/sites/xxxxx -ListTitle "リスト名" -UserName "ユーザー名"
#>

# パラメータ定義
[CmdletBinding()]
Param(
	[Parameter(Mandatory=$true)]
	[string]$WebUrl,
	[Parameter(Mandatory=$true)]
	[string]$ListTitle,
	[Parameter(Mandatory=$true)]
	[string]$UserName
)

Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

# 出力ファイルパス
$global:cstrOutputFilePath = ".\${ListTitle}.csv"


# メイン処理
function Get-ColumnFields ($objCredentials) {

    $strPassword = Read-Host -Prompt "Enter password" -AsSecureString
    $objCred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $strPassword)
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($WebUrl);
    $ctx.Credentials = $objCred;
    $objSite = $ctx.Site;
    $ctx.Load($objSite);
    $ctx.ExecuteQuery();
    
    $objWeb = $ctx.Web
    $ctx.Load($objWeb)

    $objList = $ctx.Web.Lists.GetByTitle($ListTitle)
    $fields = $objList.Fields
    $ctx.Load($fields)
    $ctx.ExecuteQuery()

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
Get-ColumnFields
