# Params to configurate
$scorecardId = ""
$newOwnerEmail = ""
# Goal path, starts from root goal to the goal you want to reach, with every parent inbetween
# Similar to create goal
$columnNames = @("")

$api = "api.powerbi.com"
Login-PowerBI
$bearer = (Get-PowerBIAccessToken)["Authorization"]
$token = $bearer.SubString(6, $bearer.Length - 6)

while (-not $token) {
    $token = Read-Host -Prompt "Enter Power BI Access token"
}
while (-not $scorecardId) {
    $scorecardId = Read-Host -Prompt "Enter scorecard id"
}
#$newOwnerEmail = Read-Host -Prompt "New goal owner email (leave empty to skip)"

$response = Invoke-WebRequest -Uri "https://$api/v1.0/myOrg/internalScorecards($scorecardId)?`$expand=goals" -Headers @{ "Authorization" = "Bearer $token" }
$scorecard = $response.Content | ConvertFrom-Json

$connections = $scorecard.goals

Write-Host Number of goal value connections to update: $connections.Count

function UpdateGoalOwner($connection, $newOwner) {
    Write-Host " - Updating owner: " -NoNewline
    Write-Host $connection
    $response = Invoke-WebRequest `
        -Method Patch `
        -Uri "https://$api/v1.0/myOrg/internalScorecards($scorecardId)/goals($($connection.id))" `
        -Body (@{ 
            additionalOwners = $connection.additionalOwners
            name             = $connection.name
            owner            = $newOwner 
            parentId         = $connection.parentId
            startDate        = $connection.startDate
            unit             = $connection.unit
        } | ConvertTo-Json) `
        -ContentType "application/json" `
        -Headers @{ "Authorization" = "Bearer $token" }
    Write-Host -ForegroundColor green OK
}

foreach ($connection in $connections) {
    Write-Host -NoNewLine "Updating goal ""$($connection.name)"" ($($connection.id)) "
    if ($newOwnerEmail) {
        UpdateGoalOwner -connection $connection -newOwner $newOwnerEmail
    } 
}
