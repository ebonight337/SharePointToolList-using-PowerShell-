param (
    [parameter(mandatory=$true)][array]$csv,
    [parameter(mandatory=$true)][string]$url
)

#########################
# Valiables
#########################
$credential = Get-Credential

#########################
# ページの作成　Webパーツ含
#########################
function Create-Pages(){
    param (
        [string]$pageName,
        [string]$layoutType="Article",
        [string]$promoteAs="NewsArticle",
        [parameter(mandatory=$true)][array]$csv
    )
    Write-Output("${pageName}　Start processing")

    if($csv.pageName -eq ""){
        Write-Output("Not found pageName value.")
        return
    }


    $page = Add-PnPClientSidePage -Name $pageName -LayoutType $layoutType -PromoteAs $promoteAs
    if(!($csv.Property1 -eq "")){
        if(!($csv.Property2 -eq "")){
        Set-PnPListItem -List "サイトのページ" -Identity $page.PageListItem.Id -Values @{"Property1"=$csv.Property1; "Property2"=$csv.Property2}
        }else{
            Set-PnPListItem -List "サイトのページ" -Identity $page.PageListItem.Id -Values @{"Property1"=$csv.Property1}
        }
    }elseif(!($csv.Property2 -eq "")){
        Set-PnPListItem -List "サイトのページ" -Identity $page.PageListItem.Id -Values @{"Property2"=$csv.Property2}
    }
    
    if(!($csv.Thumbnail -eq "")){
        Set-PnPClientSidePage $page.PageListItem.FieldValues.FileLeafRef -ThumbnailUrl $csv.Thumbnail -Publish
    }
    if(!($csv.HeaderImage -eq "")){
        $createPage = Get-PnPClientSidePage $page.PageListItem.FieldValues.FileLeafRef
        $createPage.PageHeader.ImageServerRelativeUrl = $csv.HeaderImage
        $createPage.Save()
        $createPage.Publish()
    }
    if(($csv.Thumbnail -eq "") -and ($csv.HeaderImage -eq "")){
        Set-PnPClientSidePage $page.PageListItem.FieldValues.FileLeafRef -Publish
    }

    return
}

function Main(){
    Connect-PnPOnline $url -Credentials $credential
    $csv = Import-Csv $csv
    
    foreach($i in $csv){
        Create-Pages -pageName $i.pageName -csv $i
    }
    
    return
}

Main
