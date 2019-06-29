# 引数
Param(
    [string]$url,
    [string]$username,
    [string]$password,
    [string]$listTitle
)
$secpwd = $password |ConvertTo-SecureString -AsPlainText -force

# フォルダ作成関数
function CreateFolder($folderName){
    if(Test-Path $folderName){
    }else{
        New-Item $folderName -ItemType Directory | Out-Null
        return
    }
}

# コンポーネント読み込み
Set-Location $PSScriptRoot 
Write-Host "Load CSOM libraries" -foregroundcolor black -backgroundcolor yellow 
if((test-path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll") -and (test-path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll")){
    Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
    Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
    Write-Host "CSOM libraries loaded successfully" -foregroundcolor black -backgroundcolor Green 
}else{
    throw('sharepointclientcomponentsが端末にインストールされていることを確認してください。')
}

# サイトのコンテキスト読み込み
Write-Host "authenticate to SharePoint Online site collection $url and get ClientContext object" -foregroundcolor black -backgroundcolor yellow 
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($url)  
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $secpwd)
$Context.Credentials = $credentials 
$web = $context.Web 
$List = $web.Lists.GetByTitle($listTitle)
$fields = $List.Fields         
$site = $context.Site  
$context.Load($web)
$context.Load($List)

try{ 
    $context.ExecuteQuery() 
    Write-Host "authenticateed to SharePoint Online site collection $url and get ClientContext object succeefully" -foregroundcolor black -backgroundcolor Green 
}catch{ 
    Write-Host "Not able to authenticateed to SharePoint Online site collection $url $_.Exception.Message" -foregroundcolor black -backgroundcolor Red 
    return 
} 

# コンテンツタイプの読み込み
try { 
    $contentTypes = $List.ContentTypes
    $context.Load($contentTypes)
    $context.ExecuteQuery() 
}catch{ 
  Write-Host "Content Type" $contentType "not found in site content types collection. $_.Exception.Message"  -foregroundcolor red 
  return 
} 

# コンテンツタイプをすべてエクスポート
$now = get-date -Format "yyyyMMddhhmmss"
CreateFolder($listTitle)
$contentTypes | %{
    $Datas = @()
    $fieldRefCollection = $_.FieldLinks
    $context.Load($fieldRefCollection)
    $context.ExecuteQuery()
    $ctName = $_.Name
    $fieldRefCollection | %{
        $Data = New-Object PSObject | Select-Object Name, Required, hidden
        $Data.Name = $_.Name
        $Data.Required = $_.Required
        $Data.hidden = $_.hidden
        $Datas += $Data
        
        $Datas | Export-Csv -Encoding UTF8 -Path "./${listTitle}/${listTitle}_${ctName}_${now}.csv" -NoTypeInformation
    }
}

write-host "Contenttype is exist." -foregroundcolor black -backgroundcolor Green 
