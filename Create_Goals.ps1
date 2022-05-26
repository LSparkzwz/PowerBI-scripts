# Params to configurate
$groupId = ""
$scorecardId = ""
# Excel columns of your goals, the first column is a root goal, the second a son of that root goal, the third the son of the son etc..
# ex. A B C columns
# If there's an excel row with A = "1", B = "2", C = "3"
# The goals that will be created from that row are 1, 2 with 1 as parent, 3 with 2 as parent
#
# Duplicates with the same parent (or no parent if root) won't be created, for example
# If there's three rows:
# A = "1", B = "2", C = "3"
# A = "1", B = "2", C = "4"
# A = "5", B = "2", C = "3"
# The goals that will be created from that row are:
# 1, 2 with 1 as parent, 3 with 2 as parent, 4 with 2 as parent, (skipped creating 1 and 2 again because they're duplicates)
# 5, 2 with 5 as parent, 3 with 5 as parent (duplicates with different parents are permitted)
$columnNames = @("")
$excelPath = ""  
$sheetName = "" #Excel sheet 

Login-PowerBI
$bearer = (Get-PowerBIAccessToken)["Authorization"]
	
$api = "api.powerbi.com"
$token = $bearer.SubString(6, $bearer.Length - 6)
$rules

while (-not $token) {
	$token = Read-Host -Prompt "Enter Power BI Access token"
}
while (-not $groupId) {
	$groupId = Read-Host -Prompt "Enter group id"
}
while (-not $scorecardId) {
	$scorecardId = Read-Host -Prompt "Enter scorecard id"
}

Import-Module -Name ImportExcel

# Get a specific worksheet of an Excel file
$worksheetObj = Import-Excel -Path $excelPath -WorksheetName $sheetName
$totalNoOfRecords = $worksheetObj.Count

Write-Host Creating goals

# Goals that are already uploaded in Powerbi
$response = Invoke-WebRequest -Uri "https://$api/v1.0/myorg/groups/{$groupId}/scorecards($scorecardId)/goals" -Headers @{ "Authorization" = "Bearer $token" }
$scorecard = $response.Content | ConvertFrom-Json
# We will add newly created goals to this object to keep track of them, we do this locally because it's faster than doing the get every loop
$currentGoals = $scorecard.value | Select id, name, parentId

if (!$currentGoals) {
	Write-Host There is no already created goals
	$currentGoals = @()
}

if ($totalNoOfRecords -gt 1) {  
	# Loop to get values from excel file  
	for ($i = 0; $i -le $totalNoOfRecords - 1; $i++) {  
		$parentId = ""
		$rootGoalName = $worksheetObj[$i] | Select -ExpandProperty $columnNames[0]

		# Check if the root goal already exists (root goals = the goals at top level, which have no parents)
		$rootGoals = $currentGoals | where name -eq $rootGoalName

		# Root goal not created yet
		if ($rootGoals.Count -eq 0) {
			# Create Root goal
			$body = @{ name = $rootGoalName }  | ConvertTo-Json -Depth 10
			$response = Invoke-WebRequest -Method Post -Uri "https://api.powerbi.com/v1.0/myorg/groups/$groupId/scorecards($scorecardId)/goals" -ContentType "application/json" -Headers @{ "Authorization" = "Bearer $token" } -Body $body
			$response = $response.Content | ConvertFrom-Json

			# Update the collection of already created goals
			$currentGoals += [PSCustomObject]@{
				id       = $response.id
				name     = $response.name
				parentId = $response.parentId
			}
            
			# We save the parentId because we need it to create the sub goals found in this same Excel row
			$parentId = $response.id
		}
		else {
			# We get the id of the root goal so we can create the  subgoal
			# where returns a list of filtered values, we know that this filter will always return only one value because root names are unique
			$parentId = $rootGoals[0].id
		}

		# Create the subgoals

		for ($j = 1; $j -le $columnNames.Count - 1; $j++) {
			# Check if SubGoal already exists (we check its name and if it's parentId matches with the current parentId we have)
			$existsSubGoal = $false

			$subGoalName = $worksheetObj[$i] | Select -ExpandProperty $columnNames[$j]

			$subGoals = $currentGoals | Where-Object { ($_.name -eq $subGoalName) -and ($_.parentId -eq $parentId) }
			if ($subGoals.Count -ge 1) {
				$existsSubGoal = $true
				# We get the id of the  subgoal so we can create the next subgoal
				$parentId = $subGoals[0].id
			}

			#  subgoal not created yet
			if (!$existsSubGoal) {
				# Create  subgoal
				$body = @{ name = $subGoalName; parentId = $parentId }  | ConvertTo-Json -Depth 10
				$response = Invoke-WebRequest -Method Post -Uri "https://api.powerbi.com/v1.0/myorg/groups/$groupId/scorecards($scorecardId)/goals" -ContentType "application/json" -Headers @{ "Authorization" = "Bearer $token" } -Body $body
				$response = $response.Content | ConvertFrom-Json

				# Update the collection of already created goals
				$currentGoals += [PSCustomObject]@{
					id       = $response.id
					name     = $response.name
					parentId = $response.parentId
				}

				# We get the id of the  subgoal so we can create the next subgoal
				$parentId = $response.id
			}

		}
	}  
}  
Write-Host $subGoalName $parentId
Write-Host $response

Write-Host Goals Created

 

