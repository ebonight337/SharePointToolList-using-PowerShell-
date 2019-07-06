# 引数
Param(
    [string]$importcsv,
    [string] $url,
    [string] $username,
    [string] $password,
    [string]$listTitle,
    [string]$contentTypeName,
    [string]$columnName,
    [bool]$isSiteColumn=$False,
    [bool]$Required=$True,
    [bool]$Hidden=$False
)
$secpwd = $password |ConvertTo-SecureString -AsPlainText -force

# コンポーネント読み込み
Set-Location $PSScriptRoot 
Write-Host "Load CSOM libraries" -foregroundcolor black -backgroundcolor yellow 
if((test-path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll") -and (test-path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll")){
    Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
    Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
    Write-Host "CSOM libraries loaded successfully" -foregroundcolor black -backgroundcolor Green 
}else{
    throw('Please confirm the sharepointclientcomponents is installed')
}

Function Add-columnToListContentType(){
    Param(
        [Parameter(Mandatory=$true)] [string]$url,
        [Parameter(Mandatory=$true)] [string]$username,
        [Parameter(Mandatory=$true)] [string]$password,
        [Parameter(Mandatory=$true)] [string]$listTitle,
        [Parameter(Mandatory=$true)] [string]$contentTypeName,
        [Parameter(Mandatory=$true)] [string]$columnName,
        [Parameter(Mandatory=$true)] [bool]$isSiteColumn,
        [Parameter(Mandatory=$true)] [bool]$Required,
        [Parameter(Mandatory=$true)] [bool]$Hidden
    )

    # サイトのコンテキスト読み込み
    try{ 
        Write-Host "authenticate to SharePoint Online site collection $url and get ClientContext object" -foregroundcolor black -backgroundcolor yellow 
        $context = New-Object Microsoft.SharePoint.Client.ClientContext($url)  
        $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $secpwd)
        $context.Credentials = $credentials 
        $list = $context.Web.Lists.GetByTitle($listTitle)
        $context.Load($list)
        $context.ExecuteQuery() 

        Write-Host "authenticateed to SharePoint Online site collection $url and get list object succeefully" -foregroundcolor black -backgroundcolor Green 
    }catch{ 
        Write-Host "Not able to authenticateed to SharePoint Online site collection $url $_" -foregroundcolor black -backgroundcolor Red 
        return 
    } 

    # コンテンツタイプの読み込み
    try { 
        $contentTypes = $list.ContentTypes
        $context.Load($contentTypes)
        $context.ExecuteQuery() 
    }catch{ 
        Write-Host "Content Type" $contentType "not found in site content types collection. $_" -foregroundcolor black -backgroundcolor Red 
        return 
    } 

    # コンテンツタイプのチェック
    $contentType = $contentTypes | Where {$_.Name -eq $contentTypeName}
    If($contentType -eq $Null)
    {
        Write-host "Content Type '$contentTypeName' doesn't exists in '$listTitle'" -foregroundcolor black -backgroundcolor Red 
        return
    }



    # 追加する列を取得します。サイト列または既存のリストの列のいずれかを判定
    If($isSiteColumn){
        $columnColl = $context.Web.Fields
    }else{
        $columnColl = $list.Fields
    }

    $context.Load($columnColl)
    $context.ExecuteQuery()
    $column = $columnColl | Where {$_.Title -eq $columnName}

    if($column -eq $Null)
    {
        Write-host "Column '$columnName' doesn't exists!" -foregroundcolor black -backgroundcolor Red 
        return
    }
    else
    {
        #コンテンツタイプに存在する列か確認
        $fieldCollection = $contentType.Fields
        $context.Load($fieldCollection)
        $context.ExecuteQuery()
        $field = $fieldCollection | Where {$_.Title -eq $columnName}
        if($Field -ne $Null)
        {
            Write-host "Column '$columnName' Already Exists in the content type!" -foregroundcolor black -backgroundcolor Red 
            return
        }

        #コンテンツタイプを追加
        Write-host "Column '$columnName' Added to '$contentTypeName'" -foregroundcolor black -backgroundcolor yellow 
        try{
            $fieldLink = New-Object Microsoft.SharePoint.Client.FieldLinkCreationInformation
            $fieldLink.Field = $column
            # Set "Required" and "Hidden" fileds.
            $fieldLink.Field.Required = $Required
            $fieldLink.Field.Hidden = $Hidden
            [Void]$ContentType.FieldLinks.Add($fieldLink)
            $contentType.Update($false)
            $context.ExecuteQuery() 

            Write-host "Column '$columnName' Add to '$contentTypeName' Successfully!" -foregroundcolor black -backgroundcolor Green
        }catch{
            Write-Host "Could not add " $contentType ". $_" -foregroundcolor black -backgroundcolor Red 
            return 
        }
    }
}

# ImportするCSVの行ごとに関数を実行
if(($importcsv) -and (Test-Path $importcsv)){
    $csv = Import-Csv $importcsv
    foreach($i in $csv){
        if($i.listTitle){$listTitle = $i.listTitle}
        if($i.contentTypeName){$contentTypeName = $i.contentTypeName}
        if($i.columnName){$columnName = $i.columnName}
        if($i.isSiteColumn){$isSiteColumn = [System.Convert]::ToBoolean($i.isSiteColumn)}
        if($i.Required){$Required = [System.Convert]::ToBoolean($i.Required)}
        if($i.Hidden){$Hidden = [System.Convert]::ToBoolean($i.Hidden)}

        Add-columnToListContentType -url $url -username $username -password $secpwd -listTitle $listTitle -contentTypeName $contentTypeName -columnName $columnName -isSiteColumn $isSiteColumn -Required $Required -Hidden $Hidden
    }
}elseif(!$importcsv){ # CSVを利用しない場合は引数そのまま実行
        Add-columnToListContentType -url $url -username $username -password $secpwd -listTitle $listTitle -contentTypeName $contentTypeName -columnName $columnName -isSiteColumn $isSiteColumn -Required $Required -Hidden $Hidden
}elseif(!(Test-Path $importcsv)){
    Write-Host "Could not load csv. check follow the error. $_" -foregroundcolor black -backgroundcolor Red 
}

write-host "Add Contenttype is exist." -foregroundcolor black -backgroundcolor Green 
