function DevOpsGetWorkItemsByPullRequestIds
(
    [String] $DevopsBaseUrl,
    [String] $DevopsProject,
    [String] $DevOpsPAT,
    [string] $DevOpsPullRequestIds
)
{
    write-host "******* Getting workitems from DevOps"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12
    $IDs = $DevOpsPullRequestIds.Split(",")
    if($DevOpsPAT){        
        $AdoAuthenicationHeader = @{Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($DevOpsPAT)")) }                
        
        $IDs | ForEach-Object { 
            write-host "Pull request ID $_"
            
            #Get pull request by ID
            $PullRequestUrl = "$DevopsBaseUrl/$DevopsProject/_apis/git/pullrequests/$($_)?api-version=5.1"                        
            $PullRequestResponse = Invoke-RestMethod -Uri $PullRequestUrl -Headers $AdoAuthenicationHeader -Method get        
    
            #Get work items list from pull request
            $RepositoryUrl = $PullRequestResponse.repository.url
            $RepositoryUrl += "/pullRequests/$_/workitems"    
            $response1 = Invoke-RestMethod -Uri $RepositoryUrl -Method Get -Headers $AdoAuthenicationHeader                                  

            $item = $response1.value | select-object -ExpandProperty id            
            $AllItems += $item
        }

        if ($AllItems){
            $AllItems = $AllItems | Select-Object -Unique
            $items = [system.String]::Join(",", $AllItems)                        
            
            #Get all items
            $url = "$DevopsBaseUrl/$DevopsProject/_apis/wit/workitems?ids=$items&fields=System.Id,System.Title,System.WorkItemType,System.Parent&api-version=5.1"                        
            $response2 = Invoke-RestMethod -Uri $url -Headers $AdoAuthenicationHeader -Method get
            
            #Get the parent ids of tasks
            $partenIds =@()
            $allWorkInfo = $response2.value 
            $allWorkInfo  | ForEach-Object {
                if(($_.fields.'System.WorkItemType' -eq 'Task') -and ($_.fields.'System.Parent')){
                    $partenIds += $_.fields.'System.Parent'                
                }            
            }
            
            #Get the PBI work items of tasks
            if($partenIds){
                $partenIds = $partenIds | Select-Object -Unique
                $itemsParentIds = [system.String]::Join(",", $partenIds)
                
                $url = "$DevopsBaseUrl/$DevopsProject/_apis/wit/workitems?ids=$itemsParentIds&fields=System.Id,System.Title,System.WorkItemType,System.Parent&api-version=5.1"                        
                $response3 = Invoke-RestMethod -Uri $url -Headers $AdoAuthenicationHeader -Method get
                $allWorkInfo += $response3.value
            }
    
            #Filter by PBI and Bugs only
            $allWorkInfo = $allWorkInfo | Where-Object {$_.fields.'System.WorkItemType' -in ('Product Backlog Item','Bug')}            
    
            #Remove duplicated
            $allWorkInfo = $allWorkInfo | Select-Object -Property id,fields -Unique            
        }
        return $allWorkInfo
    }else{
        Write-Error "PAT cannot be empty"
    }
}