# Params to configurate
$groupId = ""
$scorecardId = ""

$api = "api.powerbi.com"
Login-PowerBI
$bearer = (Get-PowerBIAccessToken)["Authorization"]
$token = $bearer.SubString(6, $bearer.Length - 6)

while (-not $token) {
    $token = Read-Host -Prompt "Enter Power BI Access token"
}
while (-not $groupId) {
    $groupId = Read-Host -Prompt "Enter group ID"
}
while (-not $scorecardId) {
    $scorecardId = Read-Host -Prompt "Enter scorecard ID"
}

$response = Invoke-WebRequest -Uri "https://$api/v1.0/myorg/groups/{$groupId}/scorecards($scorecardId)/goals" -Headers @{ "Authorization" = "Bearer $token" }
$scorecard = $response.Content | ConvertFrom-Json
$currentGoals = $scorecard.value 

Write-Host $currentGoals[1]

function GetGoalSons($goal) {
    $goalList = @()
    $sons = $currentGoals | where parentId -eq $goal.id
    if ($sons.Count -gt 0) {
        Write-Host "`n------------------------------------------"
        Write-Host Sons of goal: $goal.name $goal.id
        Write-Host "------------------------------------------"

        foreach ($son in $sons) {
            Write-Host
            Write-Host Name: $son.name 
            Write-Host ID: $son.id 
            Write-Host Owner: $son.owner
            Write-Host Parent: $son.parentId
            $goalList += $son.id
        }
        # Breadth search instead of depth so it can be read better 
        foreach ($son in $sons) {
            $goalList += GetGoalSons -goal $son
        }
    }
    return $goalList
}

while ($true) {
    $goalId = Read-Host -Prompt "`nEnter goal id"
    # Get goal
    $goal = $currentGoals | where id -eq $goalId
    Write-Host `nInput Goal:`n
    Write-Host Name: $goal.name 
    Write-Host ID: $goal.id 
    Write-Host Owner: $goal.owner

    Write-Host `nSubtree of the requested Goal: $goal.name $goal.id`n
    # Get entire subtree of the requested goal
    $goalList = @()
    $goalList += $goalId
    $goalList += GetGoalSons -goal $goal -result $goalList

    # The list is not properly ordered
    Write-Host `nResult:
    Write-Host $goalList
}

