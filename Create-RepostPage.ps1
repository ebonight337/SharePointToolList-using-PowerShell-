
$sitePagesURL ="https://[tenant].sharepoint.com/sites/[siteName]"
$pageLib = "サイトのページ"
$targetPageName = "ホーム"
$MyCSVPublishingDate = Get-Date
$userName = "[Login user name]"
$password = "[pwd]"
$secPassword = ConvertTo-SecureString $password -AsPlainText -Force
$newPageFileName = "[Create page file name]"
$NewsPageFileName = "/sites/testSite/SitePages/"+ $newPageFileName
$originalSourceUrl = "/sites/testSite/SitePages/test11.aspx"

function Load-DLLandAssemblies{
	[string]$defaultDLLPath = ""

	# Load assemblies to PowerShell session 
	$defaultDLLPath = "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.SharePoint.Client.dll"
	[System.Reflection.Assembly]::LoadFile($defaultDLLPath)

	$defaultDLLPath = "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.SharePoint.Client.Runtime.dll"
	[System.Reflection.Assembly]::LoadFile($defaultDLLPath)

	$defaultDLLPath = "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.Online.SharePoint.Client.Tenant.dll"
	[System.Reflection.Assembly]::LoadFile($defaultDLLPath)

	$defaultDLLPath = "C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Microsoft.SharePoint.Client.Search\v4.0_16.0.0.0__71e9bce111e9429c\Microsoft.SharePoint.Client.Search.dll"
	#$defaultDLLPath = "D:\TOOLS\SHAREPOINT\SP_ONLINE\sharepointclientcomponents\microsoft.sharepointonline.csom.16.1.8119.1200\lib\net40-full\Microsoft.SharePoint.Client.Search.dll"
	[System.Reflection.Assembly]::LoadFile($defaultDLLPath)

}

function Create-RepostPage{
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SitePagesURL)
    $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userName,$secPassword)
    $mySiteCol = $ctx.Site
    $mySiteWeb = $ctx.web
    $myPageList = $ctx.web.Lists.GetByTitle($pageLib)
    $query = New-Object Microsoft.SharePoint.Client.CamlQuery
    $query.ViewXml="<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>0</Value></Eq></Where></Query></View>"

    $ctx.Load($mySiteCol)
    $ctx.Load($mySiteWeb)
    $ctx.Load($myPageList)
    $ctx.ExecuteQuery()
    
    $pageItems = $myPageList.GetItems($query)
    $MyEditoruserAccount = $mySiteWeb.EnsureUser('i:0#.f|membership|'+ $userName)
    $ctx.Load($pageItems)
    $ctx.load($MyEditoruserAccount)
	$ctx.executeQuery()


	$NewPageitem = $MyPagelist.RootFolder.Files.AddTemplateFile($NewsPageFileName, [Microsoft.SharePoint.Client.TemplateFileType]::ClientSidePage).ListItemAllFields
	# Make this page a "modern" page
	$NewPageitem["ContentTypeId"] = "0x0101009D1CB255DA76424F860D91F20E6C4118002A50BFCFB7614729B56886FADA02339B00AD3FEC8D5D31954785347F0053E63178";
	$NewPageitem["PageLayoutType"] = "RepostPage"
	$NewPageitem["PromotedState"] = "2"
	$NewPageitem["Title"] = "test10"
	$NewPageitem["ClientSideApplicationId"] = "b6917cb1-93a0-4b97-a84d-7cf49975d4ec"

	$NewPageitem["_OriginalSourceUrl"] =  $originalSourceUrl
	$NewPageitem["Editor"] = $MyEditoruserAccount.Id
	$NewPageitem["Author"] = $MyEditoruserAccount.Id
	$NewPageitem["Description"] = "コマンドでの実装テスト"
	$NewPageitem["BannerImageUrl"] = "/sites/testSite/SitePages"
	$NewPageitem["Modified"] = $MyCSVPublishingDate
	$NewPageitem["Created"] = $MyCSVPublishingDate
	$NewPageitem["Created_x0020_By"] = $MyEditoruserAccount.LoginName
	$NewPageitem["Modified_x0020_By"] = $MyEditoruserAccount.LoginName
	$NewPageitem["FirstPublishedDate"] = $MyCSVPublishingDate
	$NewPageitem.Update()
	$ctx.Load($NewPageitem)
	$ctx.ExecuteQuery()



}


$pageItems[11].File.Publish("test")
$pageItems[11].Update()
