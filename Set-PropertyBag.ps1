Param(
    [string]$targetSite = "https://hogehuga.sharepoint.com/sites/test1",
    [string]$registKey = "key",
    [string]$registVal = "val"
)

# コンポーネント読み込み
Function Load-Module () {
    if((test-path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll") -and (test-path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll")){
        Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
        Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
        return $true
    }else{
        Write-Host "SharePoint コンポーネントが端末にインストールされていることを確認してください。"
        return $false
    }
}

# 認証情報の取得
Function Get-Cred () {
    try{
	    $objCredential = Get-Credential
        $objCredential = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials ($objCredential.UserName, $objCredential.Password)
    }catch{
        Write-Host "認証情報の取得に失敗しました。"
        return
    }
	return $objCredential
}

# サイト コンテキストの取得
Function Get-SiteContext ($strUrl, $objCred)  {
	$objContext = New-Object Microsoft.SharePoint.Client.ClientContext ($strUrl)
	$objContext.Credentials = $objCred

	try {
        $objSite = $objContext.Site
        $objWeb = $objContext.Web
        $objContext.Load($objSite)
        $objContext.Load($objWeb)
		$objContext.ExecuteQuery()
		return $objContext
	} catch {
		Write-Host "SharePoint サイトの接続に失敗しました。"
		return $null
	}
}

#プロパティバッグの登録
Function Set-PropertyBag($objCtx, $Key, $Value){

    $objCtx.Web.AllProperties[$Key] = $Value
    $targetTeam = $targetList.TeamsName
    try{
        $objCtx.Web.Update()
        $objCtx.ExecuteQuery()
        return $true
    }catch{
		Write-Host "プロパティバッグの登録に失敗しました。"
        return $null
    }
}

# メイン処理
Function Start-Process ($targetSite) {
    # コンポーネント読み込み
    $mod = Load-Module
    if ($mod -eq $false) {
	return
    }

    # 認証情報の取得
    $objCred = Get-Cred
    if($objCred -eq $null){
        return
    }

    # サイト コンテキストの取得
    $objCtx = Get-SiteContext $targetSite $objCred

    #プロパティバッグの登録
    $propBg = Set-PropertyBag $objCtx $registKey $registVal
    if($propBg -eq $null){
        continue
    }
}

Start-Process $targetSite
