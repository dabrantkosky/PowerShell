param (
    [string] $Organization,
    [string] $PAT,
    [string] $projectID
)

$AzureDevOpsAuthenicationHeader = @{Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($PAT)")) }
$UriOrganization = "https://dev.azure.com/$($Organization)/"
$UriOrganizationRelease = "https://vsrm.dev.azure.com/$($Organization)/"


   
# pipelines
$uriPipelines = $UriOrganization + $projectID + "/_apis/pipelines?api-version=7.0"
$PipelinesResult = Invoke-RestMethod -Uri $uriPipelines -Method get -Headers $AzureDevOpsAuthenicationHeader

# releases
$uriReleases = $UriOrganizationRelease + $projectID + "/_apis/release/definitions?api-version=7.0"
$ReleasesResult = Invoke-RestMethod -Uri $uriReleases -Method get -Headers $AzureDevOpsAuthenicationHeader


$Data = @{};

foreach ($pipeline in $PipelinesResult.value) { 
        
    $uriRuns = $UriOrganization + $projectID + "/_apis/pipelines/$($pipeline.id)/runs?api-version=7.0"
    $totalTime = "";
    $RunsResult = Invoke-RestMethod -Uri $uriRuns -Method get -Headers $AzureDevOpsAuthenicationHeader
      
    $PipelineDefinitions = $UriOrganization + $projectID + "/_apis/build/definitions/$($pipeline.id)?api-version=7.0"
    $DefinitionsResult = Invoke-RestMethod -Uri $PipelineDefinitions -Method get -Headers $AzureDevOpsAuthenicationHeader
        
    $succeededRuns = 0
    $runs = 0
    foreach ($run in $RunsResult.value) {
        if ($run.createdDate -gt "2022-00-00T00:00:00Z") {    
            $runs += 1
            if ($run.result -eq "succeeded") {
                $succeededRuns += 1
            }
        }
    }
    if ($null -ne ($RunsResult.value | Select-Object -First 1).createdDate) {
        $totalTime = NEW-TIMESPAN -Start ($RunsResult.value | Select-Object -First 1).createdDate -End ($RunsResult.value | Select-Object -First 1).finishedDate
        $etc = "{0:dd}d:{0:hh}h:{0:mm}m:{0:ss}s" -f $totalTime
    }

    if (($RunsResult.value | Select-Object -First 1).createdDate -lt "2023-01-01T00:00:00Z" -and $null -ne ($RunsResult.value | Select-Object -First 1).createdDate ) {
        $Data.Add($pipeline.name, @());
        $Data.($pipeline.name) = [PSCustomObject]@{
            "Pipeline Id"         = $pipeline.id
            "Pipeline Name"       = $pipeline.name
            "Type"                = "Pipeline"
            "Category"            = "Not run in 2 Months"
            "Last Run Date"       = ($RunsResult.value | Select-Object -First 1).createdDate
            "Last Run Outcome"    = ($RunsResult.value | Select-Object -First 1).result
            "Last Run Total Time" = $etc 
            "Pool"                = $DefinitionsResult.queue.pool.name
        }
            
    }
    else {
        if ($DefinitionsResult.queueStatus -ne "enabled" ) {
            $Data.Add($pipeline.name, @());
            $Data.($pipeline.name) = [PSCustomObject]@{
                "Pipeline Id"         = $pipeline.id
                "Pipeline Name"       = $pipeline.name
                "Type"                = "Pipeline"
                "Category"            = "Currently Disabled or Paused"
                "Last Run Date"       = ($RunsResult.value | Select-Object -First 1).createdDate
                "Last Run Outcome"    = ($RunsResult.value | Select-Object -First 1).result
                "Last Run Total Time" = $etc 
                "Pool"                = $DefinitionsResult.queue.pool.name
            }
                
        }
        else {
            if ($runs -ne 0 -and $succeededRuns / $runs -lt 0.7 ) {
                $Data.Add($pipeline.name, @());
                $Data.($pipeline.name) = [PSCustomObject]@{
                    "Pipeline Id"         = $pipeline.id
                    "Pipeline Name"       = $pipeline.name
                    "Type"                = "Pipeline"
                    "Category"            = "Success rate of less than 70%"
                    "Last Run Date"       = ($RunsResult.value | Select-Object -First 1).createdDate
                    "Last Run Outcome"    = ($RunsResult.value | Select-Object -First 1).result
                    "Last Run Total Time" = $etc 
                    "Pool"                = $DefinitionsResult.queue.pool.name
                }
                    
            }
        }
    }

}

foreach ($release in $ReleasesResult.value) {
    $uriReleaseDeployments = $UriOrganizationRelease + $projectID + "/_apis/Release/deployments?definitionId=$($release.id)"
    $ReleaseDeploymentRunsResult = Invoke-RestMethod -Uri $uriReleaseDeployments -Method get -Headers $AzureDevOpsAuthenicationHeader

    $totalTime = "";
    $succeededRuns = 0
    $runs = 0
    foreach ($run in $ReleaseDeploymentRunsResult.value) {
        if ($run.startedOn -gt "2022-00-00T00:00:00Z") {    
            $runs += 1
            if ($run.deploymentStatus -eq "succeeded") {
                $succeededRuns += 1
            }
        }
    }

    if ($null -ne ($ReleaseDeploymentRunsResult.value | Select-Object -First 1).startedOn) {
        $totalTime = NEW-TIMESPAN -Start ($ReleaseDeploymentRunsResult.value | Select-Object -First 1).startedOn -End ($ReleaseDeploymentRunsResult.value | Select-Object -First 1).completedOn
        $etc = "{0:dd}d:{0:hh}h:{0:mm}m:{0:ss}s" -f $totalTime
    }
    if (($ReleaseDeploymentRunsResult.value | Select-Object -First 1).startedOn -lt "2023-01-01T00:00:00Z" -and $null -ne ($ReleaseDeploymentRunsResult.value | Select-Object -First 1).startedOn ) {
           
        $Data.Add($release.name, @());
        $Data.($release.name) = [PSCustomObject]@{
            "Pipeline Id"         = $release.id
            "Pipeline Name"       = $release.name
            "Type"                = "Release"
            "Category"            = "Not run in 2 Months"
            "Last Run Date"       = ($ReleaseDeploymentRunsResult.value | Select-Object -First 1).startedOn
            "Last Run Outcome"    = ($ReleaseDeploymentRunsResult.value | Select-Object -First 1).deploymentStatus
            "Last Run Total Time" = $etc 
            "Pool"                = "Not known"
        }
            
    }
    else {
        if ($runs -ne 0 -and $succeededRuns / $runs -lt 0.7 ) {
            $Data.Add($release.name, @());
            $Data.($release.name) = [PSCustomObject]@{
                "Pipeline Id"         = $release.id
                "Pipeline Name"       = $release.name
                "Type"                = "Release"
                "Category"            = "Success rate of less than 70%"
                "Last Run Date"       = ($ReleaseDeploymentRunsResult.value | Select-Object -First 1).startedOn
                "Last Run Outcome"    = ($ReleaseDeploymentRunsResult.value | Select-Object -First 1).deploymentStatus
                "Last Run Total Time" = $etc 
                "Pool"                = "Not known"
            }
                
        }
    }
}
    
$Data.Keys | ForEach-Object { [PSCustomObject] @{Name = $Data[$_]."Pipeline Name"; Id = $Data[$_]."Pipeline Id"; Type = $Data[$_].Type; Category = $Data[$_].Category; "Last Run Date" = $Data[$_]."Last Run Date"; "Last Run Outcome" = $Data[$_]."Last Run Outcome"; "Last Run Total Time" = $Data[$_]."Last Run Total Time"; Pool = $Data[$_].Pool } } | ConvertTo-Csv | Out-File -FilePath "C:/temp/excel.csv"
