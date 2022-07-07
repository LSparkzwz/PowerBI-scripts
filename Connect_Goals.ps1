########## Params to configurate START ##########

# Access Token (found by checking the browser Network messages when connecting a goal on the PowerBI site)
$bearer = ""

# Set ValueConnection or TargetConnection
# $connectionType = "TargetConnection"
$connectionType = "ValueConnection"

# The scorecard you want to connect
$groupId = ""
$scorecardId = ""

### Query params (you need to find these paramaters by checking a manual query with Postman) START ###
# Goal path, starts from root goal to the goal you want to reach, with every parent inbetween
$columnNames = @("val1", "val2", "etc...")

# Get these params by doing a manual query once and checking the network payload
# ValueConnection
$vDatasetId = ""
$vUserId = ""
$vReportUrl = ""
$vOwner = ""
$vFrom = @(
    [ordered]@{
        Name   = ""
        Entity = ""
        Type   = 0
    },
    [ordered]@{
        Name   = ""
        Entity = ""
        Type   = 0
    }
) | ConvertTo-Json

# TargetConnection 
$tDatasetId = ""
$tUserId = ""
$tReportUrl = ""
$tOwner = ""
$tFrom = @(
    [ordered]@{
        Name   = ""
        Entity = ""
        Type   = 0
    },
    [ordered]@{
        Name   = ""
        Entity = ""
        Type   = 0
    }
) | ConvertTo-Json

### Query params END ###

########## Params to configurate END ##########

$api = "api.powerbi.com"
# Login-PowerBI
# $bearer = (Get-PowerBIAccessToken)["Authorization"]
$token = $bearer.SubString(6, $bearer.Length - 6)

while (-not $token) {
    $token = Read-Host -Prompt "Enter Power BI Access token"
}
while (-not $scorecardId) {
    $scorecardId = Read-Host -Prompt "Enter scorecard id"
}

# Goals that are already uploaded in Powerbi
$response = Invoke-WebRequest -Uri "https://$api/v1.0/myorg/groups/{$groupId}/scorecards($scorecardId)/goals" -Headers @{ "Authorization" = "Bearer $token" }
$scorecard = $response.Content | ConvertFrom-Json
# We will add newly created goals to this object to keep track of them, we do this locally because it's faster than doing the get every loop
$currentGoals = $scorecard.value | Select id, name, parentId, owner

$goalCount = $currentGoals.Count
$columnCount = $columnNames.Count
for ($i = 0; $i -le $goalCount - 1; $i++) {  
    $currentGoal = $currentGoals[$i]
    #Write-Host "Current Goal: " $currentGoal `n
    # We get the parent of the current goal and its parent's parent, and so on 
    # We count how many parents are found so we know in which column this goal belongs:
    # ex. we found 2 parents, this means that $currentGoal belongs to
    # the column Policy LOB because there are the WAMJLG and Company parents
    # ex2. we found 1 parent, this means that $currentGoal belongs to
    # the column WAMJLG because there's only the Company parent
    $goalValues = @($currentGoal.name)
    $cg = $currentGoal
 
    for ($j = 1; $j -le $columnCount - 1; $j++) {  
        if ($cg.parentId) {
            $parent = $currentGoals | where id -eq $cg.parentId
            $goalValues += @($parent[0].name)
            $cg = $parent[0]
        }
    }

    #Write-Host "Goals: " $goalValues `n


    $DatasetId = $vDatasetId
    $UserId = $vUserId
    $ReportUrl = $vReportUrl
    $From = $vFrom
    $Owner = $vOwner
    $UpsertAPI = "UpsertGoalCurrentValueConnection()"

    if ($connectionType -eq "TargetConnection") {
        $DatasetId = $tDatasetId
        $UserId = $tUserId
        $ReportUrl = $tReportUrl
        $From = $tFrom
        $Owner = $tOwner
        $UpsertAPI = "UpsertGoalTargetValueConnection()"
    }

    $query = @"
    {
        "type": "Current",
        "datasetId": "$DatasetId",
        "userId": "$UserId",
        "owner": "$Owner",
        "reportUrl": "$ReportUrl",
        "query": {
            "Commands": [
                {
                    "SemanticQueryDataShapeCommand": {
                        "Query": {
                            "Version": 2,
                            "From": $From,
                            "Select": [
                                {
                                    "Measure": {
                                        "Expression": {
                                            "SourceRef": {
                                                "Source": "c"
                                            }
                                        },
                                        "Property": "GWP YTD 2022"
                                    },
                                    "Name": "CRM.GWP YTD 2022"
                                }
                            ],
                            "Where": []
                        },
                        "Binding": {
                            "Primary": {
                                "Groupings": [
                                    {
                                        "Projections": [
                                            0
                                        ]
                                    }
                                ]
                            },
                            "DataReduction": {
                                "DataVolume": 3,
                                "Primary": {
                                    "Top": {
                                        "Count": 30000
                                    }
                                }
                            },
                            "Version": 1
                        }
                    }
                }
            ]
        },
        "shouldClearGoalValues": false
    }
"@

    $objectQuery = $query | ConvertFrom-Json
    $gvCount = $goalValues.Count

    # We get the column name of the current goal by checking how many parents we found
    for ($j = 0; $j -le $gvCount - 1; $j++) {  
        $goalColumnName = $columnNames[$gvCount - 1 - $j]
        #arrays containing the information about each goal found in the for before
        #goalColumnName = the name of the excel column (ex Company)
        #goalValue = the name value of the goal (ex Italy AG)
        #fromNames = the abbreviation of the excel column used in the powerbi query (ex c)
        #Write-Host $goalColumnName  :  $goalValues[$j] : $currentGoal

        # we add the conditions to the query
        $objectQuery.query.Commands[0].SemanticQueryDataShapeCommand.Query.Where += [ordered]@{
            Condition = @{
                In = [ordered]@{
                    Expressions = @(
                        [ordered]@{
                            Column = [ordered]@{
                                Expression = [ordered]@{
                                    SourceRef = [ordered]@{
                                        Source = "p"
                                    }
                                }
                                Property   = $goalColumnName #ex COMPANY
                            }
                        }
                    )
                    Values      = @(
                        , @(
                            [ordered]@{
                                Literal = [ordered]@{
                                    Value = "'" + $goalValues[$j] + "'" #ex Italy AG
                                }
                            }
                        )
                    )
                }
            }
        }
    }

    $query = $objectQuery | ConvertTo-Json -Depth 99
    #Write-Host $query
    #Gotta turn the query into a string because PowerBi wants it like that
    $q = $objectQuery.query 
    #Write-Host $query
    $qString = $q | ConvertTo-Json -Depth 99
    #remove newlines
    $qString = $qString -replace "`n|`r" 
    #keep spaces between values with spaces
    $qString = $qString -replace "(?<=[a-zA-Z0-9]) (?=[a-zA-Z0-9])", "@#@" 
    #remove extra spaces
    $qString = $qString -replace '\s+', ''
    #replace the @#@ with space
    $qString = $qString -replace '@#@', ' '
    #Write-Host $qString
            
    $objectQuery.query = $qString
    #Write-Host $objectQuery.query
    $query = $objectQuery | ConvertTo-Json -Depth 99
    Write-Host $query

    $api = "api.powerbi.com"
    $response = Invoke-WebRequest `
        -Method Post `
        -Uri "https://$api/v1.0/myOrg/internalScorecards($scorecardId)/goals($($currentGoal.id))/$UpsertAPI" `
        -Body ($query) `
        -ContentType "application/json" `
        -Headers @{ "Authorization" = "Bearer $token" }
}
