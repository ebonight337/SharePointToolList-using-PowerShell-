# 変数
$csvPath = "C:\temp\SiteList.csv"
$now = (Get-Date).ToString("yyyyMMddHHmmss")
$outPutFileName = "C:\temp\outputPage_${now}.csv"
$timeToAdd = "9:00"

# Read CSOM
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime") | Out-Null

# 認証情報ポップアップ
$cred = Get-Credential

$csv = Import-Csv -Path $csvPath
$csv | %{

    $targetURL = $_.URL
    $target = $_.ListName
    if($_.isList -eq $null){
        $isList = [System.Convert]::ToBoolean("false")
    }else{
        $isList = [System.Convert]::ToBoolean($_.isList)
    }
    Write-Host $targetURL

    # コンテキスト初期化
    $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($cred.UserName, $cred.Password)
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($targetURL)
    $ctx.Credentials = $credentials

    # Web初期化
    $ctx.Load($ctx.Web)
    $ctx.ExecuteQuery()

    $list = $ctx.Web.Lists.GetByTitle($target)
    $query = New-Object Microsoft.SharePoint.Client.CamlQuery
    $query.ViewXml="<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>0</Value></Eq></Where></Query></View>"
    $listItems = $list.GetItems($query)
    $ctx.Load($listItems)
    $ctx.ExecuteQuery()

    
    # CSVオブジェクト生成
    $outputDatas = New-Object System.Collections.ArrayList

    foreach($Item in $ListItems)
    {
        $outputData = New-Object PSObject | Select-Object SiteName, LibralyName, Author, Name, Location, Tag, TemporaryTag, ModifiedDate, ModifiedBy
        $outputData.SiteName = $ctx.Web.Title
        $outputData.Author = $Item.FieldValues.Author.Email
        $outputData.LibralyName = $target

        $fullFilePath = [regex]::Matches(($ctx.Web.URL), "https://.*?/") | foreach{$_.value}
        $fullFilePath = $fullFilePath.Substring(0, $fullFilePath.Length-1)
        if($isList){
            $outputData.Name = $Item["Title"]
            $outputData.Location = $fullFilePath + $Item.FieldValues.FileDirRef + "/DispForm.aspx?ID=" + $Item.FieldValues.ID
        }else{
            $outputData.Name = $Item.FieldValues.FileLeafRef
            $outputData.Location = $fullFilePath+$Item["FileRef"]
        }
        if($Item.FieldValues.Tag.Values.Values.Count -ne 0){
            $tags = for($i=0; $i -lt $Item.FieldValues.Tag.Values.Values.Count+1; $i += 4){$Item.FieldValues.Tag.Values.Values[$i+1]}
            $tagList = ""
            for($i=0; $i -ne $tags.Count; $i++)
            {
                if($i -eq 0){
                    $tagList += $tags[$i]
                }elseif($i -eq $tags.count -1){
                    $tagList += $tags[$i]
                }else{
                    $tagList += ";"+$tags[$i]
                }
            }
            $outputData.Tag = $tagList
        }
        $outputData.TemporaryTag = $Item.FieldValues.Temp_Tag
        
        $outputData.ModifiedDate = $Item.FieldValues.Modified + $timeToAdd
        $outputData.ModifiedBy = $Item.FieldValues.Editor.Email

        [void]$outputDatas.Add($outputData)
    }
    
    $outputDatas | ft -AutoSize
    $outputDatas | Export-Csv $OutputFilename -Encoding Default -NoTypeInformation -Append

    $ctx.Dispose()
}
