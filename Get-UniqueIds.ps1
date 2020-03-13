#Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

#Variables for Processing
$SiteUrl = "SiteAddress"
$ListName="ドキュメント"
$UserName="hogehoge@hogehoge.onmicrosoft.com"
$Password ="******"

#Setup Credentials to connect
$Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName,(ConvertTo-SecureString $Password -AsPlainText -Force))

#Set up the context
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl) 
$Context.Credentials = $credentials

#Get the List
$List = $Context.web.Lists.GetByTitle($ListName)
#sharepoint online get list items powershell
$ListItems = $List.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()) 
$Context.Load($ListItems)
$Context.ExecuteQuery()

#Loop through each item
$ListItems | ForEach-Object {
    #Get the Title field value
    $val = $_["FileLeafRef"]
    write-host $_["FileLeafRef"] $_["UniqueId"]
}
