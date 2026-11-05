Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$domainSettingsPath = Join-Path $env:APPDATA 'SOA Konwerter\domains.json'
$domainValues = [System.Collections.Generic.List[string]]::new()
$savedDomains = @()
if (Test-Path -LiteralPath $domainSettingsPath) {
	try { $savedDomains = @(Get-Content -LiteralPath $domainSettingsPath -Raw | ConvertFrom-Json -ErrorAction Stop) } catch { $savedDomains = @() }
}
foreach ($domainValue in $savedDomains) {
	$normalizedDomain = ([string]$domainValue).Trim().TrimStart('@').TrimEnd('.').ToLowerInvariant()
	if ($normalizedDomain -match '^(?=.{1,253}$)(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,63}$') {
		$domainWithAt = "@$normalizedDomain"
		if (-not $domainValues.Contains($domainWithAt)) { $domainValues.Add($domainWithAt) }
	}
}
$appState = @{
	Domains = $domainValues
	SelectedDomain = if ($domainValues.Count -gt 0) { $domainValues[0] } else { $null }
	DomainSettingsPath = $domainSettingsPath
}

function Save-Domains {
	param($State)

	$settingsDirectory = Split-Path -Parent $State.DomainSettingsPath
	[System.IO.Directory]::CreateDirectory($settingsDirectory) | Out-Null
	$json = ConvertTo-Json -InputObject @($State.Domains.ToArray()) -Depth 2
	Set-Content -LiteralPath $State.DomainSettingsPath -Value $json -Encoding UTF8
}

function Get-GraphContext {
	if (-not (Get-Command Get-MgContext -ErrorAction SilentlyContinue)) {
		Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
	}

	$context = Get-MgContext
	if (-not $context) {
		throw 'Connect to Microsoft Graph first.'
	}

	return $context
}

function New-ActionForm {
	param([string]$Title, [int]$Height = 270)

	$form = [System.Windows.Forms.Form]::new()
	$form.Text = $Title
	$form.StartPosition = 'CenterParent'
	$form.ClientSize = [System.Drawing.Size]::new(480, $Height)
	$form.FormBorderStyle = 'FixedDialog'
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	$form.BackColor = [System.Drawing.Color]::White
	return $form
}

function Add-FormLabel {
	param($Form, [string]$Text, [int]$X, [int]$Y, [int]$Width = 420, [int]$Height = 24)

	$label = [System.Windows.Forms.Label]::new()
	$label.Text = $Text
	$label.Location = [System.Drawing.Point]::new($X, $Y)
	$label.Size = [System.Drawing.Size]::new($Width, $Height)
	$label.Font = [System.Drawing.Font]::new('Segoe UI', 9)
	$Form.Controls.Add($label)
	return $label
}

function Add-FormTextBox {
	param($Form, [int]$X, [int]$Y, [int]$Width = 420)

	$textBox = [System.Windows.Forms.TextBox]::new()
	$textBox.Location = [System.Drawing.Point]::new($X, $Y)
	$textBox.Size = [System.Drawing.Size]::new($Width, 27)
	$textBox.Font = [System.Drawing.Font]::new('Segoe UI', 9)
	$Form.Controls.Add($textBox)
	return $textBox
}

function Add-FormButton {
	param($Form, [string]$Text, [int]$X, [int]$Y, [int]$Width = 150)

	$button = [System.Windows.Forms.Button]::new()
	$button.Text = $Text
	$button.Location = [System.Drawing.Point]::new($X, $Y)
	$button.Size = [System.Drawing.Size]::new($Width, 34)
	$button.Font = [System.Drawing.Font]::new('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
	$button.UseVisualStyleBackColor = $false
	$button.BackColor = [System.Drawing.Color]::FromArgb(0, 96, 128)
	$button.ForeColor = [System.Drawing.Color]::White
	$button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(0, 76, 102)
	$button.FlatAppearance.BorderSize = 1
	$button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(0, 116, 148)
	$button.Cursor = [System.Windows.Forms.Cursors]::Hand
	$Form.Controls.Add($button)
	return $button
}

function Add-FormProgressBar {
	param($Form, [int]$X, [int]$Y, [int]$Width = 420)

	$progressBar = [System.Windows.Forms.ProgressBar]::new()
	$progressBar.Location = [System.Drawing.Point]::new($X, $Y)
	$progressBar.Size = [System.Drawing.Size]::new($Width, 16)
	$progressBar.Minimum = 0
	$progressBar.Maximum = 100
	$progressBar.Style = 'Continuous'
	$Form.Controls.Add($progressBar)
	return $progressBar
}

function Update-SoaConversionProgress {
	param(
		[System.Windows.Forms.Label]$ProgressLabel,
		[System.Windows.Forms.ProgressBar]$ProgressBar,
		[int]$Converted,
		[int]$Total
	)

	$percentage = if ($Total -gt 0) { [int][Math]::Floor(($Converted * 100.0) / $Total) } else { 0 }
	$ProgressLabel.Text = "Converted $Converted/$Total ($percentage%)"
	$ProgressBar.Value = [Math]::Max(0, [Math]::Min($percentage, 100))
	[System.Windows.Forms.Application]::DoEvents()
}

function Confirm-SoaConversion {
	param($Form, [string]$TargetDescription)

	$message = "You are about to set the following objects to Cloud Managed:`r`n$TargetDescription`r`n`r`nThis may affect how the objects are managed and synchronized with on-premises Active Directory.`r`n`r`nDo you understand the impact and want to continue?"
	$confirmation = [System.Windows.Forms.MessageBox]::Show(
		$Form,
		$message,
		'Confirm Conversion',
		[System.Windows.Forms.MessageBoxButtons]::YesNo,
		[System.Windows.Forms.MessageBoxIcon]::Warning,
		[System.Windows.Forms.MessageBoxDefaultButton]::Button2
	)
	return $confirmation -eq [System.Windows.Forms.DialogResult]::Yes
}

function Get-CsvColumn {
	param([object[]]$Rows, [string[]]$Candidates, [string]$Purpose)

	if ($Rows.Count -eq 0) { throw 'Plik CSV jest pusty.' }
	$headers = @($Rows[0].PSObject.Properties.Name)
	foreach ($candidate in $Candidates) {
		if ($headers -contains $candidate) { return $candidate }
	}

	throw "Column $Purpose was not found. Available columns: $($headers -join ', ')"
}

function Invoke-SoaStatusBatch {
	param(
		[ValidateSet('users', 'groups')][string]$ObjectType,
		[string[]]$ObjectIds,
		[System.Windows.Forms.Label]$StatusLabel,
		[System.Windows.Forms.ProgressBar]$ProgressBar
	)

	$objectIds = @($ObjectIds | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique)
	$results = [System.Collections.Generic.List[object]]::new()
	$batchSize = 20
	$checkedCount = 0
	if ($ProgressBar) { $ProgressBar.Value = 0 }
	$StatusLabel.Text = "Checked 0/$($objectIds.Count) (0%)"
	[System.Windows.Forms.Application]::DoEvents()

	for ($offset = 0; $offset -lt $objectIds.Count; $offset += $batchSize) {
		$end = [Math]::Min($offset + $batchSize, $objectIds.Count)
		$batchIds = [System.Collections.Generic.List[string]]::new()
		for ($index = $offset; $index -lt $end; $index++) { $batchIds.Add($objectIds[$index]) }
		$requests = @(
			for ($slot = 0; $slot -lt $batchIds.Count; $slot++) {
				$escapedId = [uri]::EscapeDataString($batchIds[$slot])
				@{
					id = [string]$slot
					method = 'GET'
					url = "/$ObjectType/$escapedId/onPremisesSyncBehavior?`$select=isCloudManaged"
				}
			}
		)

		try {
			$response = Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/v1.0/$batch' `
				-Method POST -Body (@{ requests = $requests } | ConvertTo-Json -Depth 6 -Compress) -ErrorAction Stop
		} catch {
			$requestError = $_.Exception.Message
			foreach ($objectId in $batchIds) {
				$results.Add([pscustomobject]@{ ObjectType = $ObjectType; Identifier = $objectId; IsCloudManaged = $null; Status = 'Error'; Error = $requestError })
			}
			$checkedCount += $batchIds.Count
			$percentage = [int][Math]::Floor(($checkedCount * 100.0) / $objectIds.Count)
			$StatusLabel.Text = "Checked $checkedCount/$($objectIds.Count) ($percentage%)"
			if ($ProgressBar) { $ProgressBar.Value = [Math]::Min($percentage, 100) }
			[System.Windows.Forms.Application]::DoEvents()
			continue
		}

		$responded = [System.Collections.Generic.HashSet[int]]::new()
		foreach ($item in $response.responses) {
			$slot = [int]$item.id
			if ($slot -lt 0 -or $slot -ge $batchIds.Count) { continue }
			[void]$responded.Add($slot)
			$objectId = $batchIds[$slot]
			if ([int]$item.status -eq 200) {
				$isCloudManaged = $item.body.isCloudManaged
				$status = if ($isCloudManaged -eq $true) { 'ConvertedToSOA' } elseif ($isCloudManaged -eq $false) { 'NotConvertedToSOA' } else { 'Unknown' }
				$errorMessage = ''
			} else {
				$isCloudManaged = $null
				$status = if ([int]$item.status -eq 404) { 'NotFound' } else { 'Error' }
				$errorMessage = if ($item.body.error.message) { $item.body.error.message } else { "HTTP $($item.status)" }
			}
			$results.Add([pscustomobject]@{ ObjectType = $ObjectType; Identifier = $objectId; IsCloudManaged = $isCloudManaged; Status = $status; Error = $errorMessage })
		}

		for ($slot = 0; $slot -lt $batchIds.Count; $slot++) {
			if (-not $responded.Contains($slot)) {
				$objectId = $batchIds[$slot]
				$results.Add([pscustomobject]@{ ObjectType = $ObjectType; Identifier = $objectId; IsCloudManaged = $null; Status = 'Error'; Error = 'Graph batch did not return a response for this object.' })
			}
		}
		$checkedCount += $batchIds.Count
		$percentage = [int][Math]::Floor(($checkedCount * 100.0) / $objectIds.Count)
		$StatusLabel.Text = "Checked $checkedCount/$($objectIds.Count) ($percentage%)"
		if ($ProgressBar) { $ProgressBar.Value = [Math]::Min($percentage, 100) }
		[System.Windows.Forms.Application]::DoEvents()
	}

	return $results.ToArray()
}

function Invoke-SoaBatchPatch {
	param(
		[string]$ObjectType,
		[string[]]$ObjectIds,
		[System.Windows.Forms.Label]$StatusLabel,
		[System.Windows.Forms.Label]$ProgressLabel,
		[System.Windows.Forms.ProgressBar]$ProgressBar
	)

	$results = [System.Collections.Generic.List[object]]::new()
	$failed = [System.Collections.Generic.List[string]]::new()
	$batchSize = 20
	$objectIds = @($ObjectIds | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique)
	$successfulCount = 0
	if ($ProgressLabel -and $ProgressBar) { Update-SoaConversionProgress $ProgressLabel $ProgressBar 0 $objectIds.Count }

	for ($offset = 0; $offset -lt $objectIds.Count; $offset += $batchSize) {
		$end = [Math]::Min($offset + $batchSize, $objectIds.Count)
		$pending = [System.Collections.Generic.List[int]]::new()
		$completed = [System.Collections.Generic.HashSet[int]]::new()
		for ($index = $offset; $index -lt $end; $index++) { $pending.Add($index) }

		$attempt = 0
		while ($pending.Count -gt 0 -and $attempt -lt 8) {
			$requests = @(
				for ($slot = 0; $slot -lt $pending.Count; $slot++) {
					$objectId = [uri]::EscapeDataString($objectIds[$pending[$slot]])
					@{
						id      = [string]$slot
						method  = 'PATCH'
						url     = "/$ObjectType/$objectId/onPremisesSyncBehavior"
						headers = @{ 'Content-Type' = 'application/json'; Prefer = 'return=representation' }
						body    = @{ isCloudManaged = $true }
					}
				}
			)

			try {
				$response = Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/v1.0/$batch' `
					-Method POST -Body (@{ requests = $requests } | ConvertTo-Json -Depth 8 -Compress) `
					-ErrorAction Stop
			} catch {
				$attempt++
				$requestError = $_.Exception.Message
				$authorizationFailure = $requestError -match '(?i)forbidden|unauthorized|access denied|status code.*(401|403)'
				if ($attempt -ge 8 -or $authorizationFailure) {
					foreach ($index in $pending) {
						$failed.Add("$($objectIds[$index]): $requestError")
						$results.Add([pscustomobject]@{ ObjectType = $ObjectType; ObjectId = $objectIds[$index]; Status = 'Error'; Error = $requestError })
						[void]$completed.Add($index)
					}
					break
				}
				$StatusLabel.Text = "Graph request failed; retry $attempt/8 in 30 sec..."
				[System.Windows.Forms.Application]::DoEvents()
				Start-Sleep -Seconds 30
				continue
			}

			$retryIndices = [System.Collections.Generic.List[int]]::new()
			$retryAfter = 0
			$responded = [System.Collections.Generic.HashSet[int]]::new()
			foreach ($item in $response.responses) {
				$slot = [int]$item.id
				if ($slot -lt 0 -or $slot -ge $pending.Count) { continue }
				$sourceIndex = $pending[$slot]
				[void]$responded.Add($sourceIndex)
				if ([int]$item.status -in @(200, 204)) {
					$results.Add([pscustomobject]@{ ObjectType = $ObjectType; ObjectId = $objectIds[$sourceIndex]; Status = 'OK'; Error = '' })
					$successfulCount++
					if ($ProgressLabel -and $ProgressBar) { Update-SoaConversionProgress $ProgressLabel $ProgressBar $successfulCount $objectIds.Count }
					[void]$completed.Add($sourceIndex)
				} elseif ([int]$item.status -eq 429) {
					$retryIndices.Add($sourceIndex)
					$delay = 30
					if ($item.headers.'Retry-After') { $delay = [int]$item.headers.'Retry-After' }
					$retryAfter = [Math]::Max($retryAfter, $delay)
				} else {
					$message = if ($item.body.error.message) { $item.body.error.message } else { "HTTP $($item.status)" }
					$failed.Add("$($objectIds[$sourceIndex]): $message")
					$results.Add([pscustomobject]@{ ObjectType = $ObjectType; ObjectId = $objectIds[$sourceIndex]; Status = 'Error'; Error = $message })
					[void]$completed.Add($sourceIndex)
				}
			}
			foreach ($sourceIndex in $pending) {
				if (-not $responded.Contains($sourceIndex) -and $sourceIndex -notin $retryIndices) {
					$retryIndices.Add($sourceIndex)
					$retryAfter = [Math]::Max($retryAfter, 5)
				}
			}

			$pending = $retryIndices
			$attempt++
			if ($pending.Count -gt 0 -and $attempt -lt 8) {
				$waitSeconds = if ($retryAfter -gt 0) { $retryAfter + 5 } else { 5 }
				$StatusLabel.Text = "Graph throttling; retrying in $waitSeconds sec..."
				[System.Windows.Forms.Application]::DoEvents()
				Start-Sleep -Seconds $waitSeconds
			}
		}

		foreach ($index in $pending) {
			if (-not $completed.Contains($index)) {
				$message = 'Retry limit exceeded.'
				$failed.Add("$($objectIds[$index]): $message")
				$results.Add([pscustomobject]@{ ObjectType = $ObjectType; ObjectId = $objectIds[$index]; Status = 'Error'; Error = $message })
			}
		}
		$processed = [Math]::Min($end, $objectIds.Count)
		$StatusLabel.Text = "Processed $processed of $($objectIds.Count)..."
		[System.Windows.Forms.Application]::DoEvents()
	}

	$reportPath = Join-Path $env:TEMP "soa_$($ObjectType)_results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
	$results | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding UTF8
	return [pscustomobject]@{
		Successful = @($results | Where-Object Status -eq 'OK').Count
		Failed = @($failed)
		Resolved = $objectIds.Count
		Results = $results.ToArray()
		ReportPath = $reportPath
	}
}

function Invoke-SoaUsersFromUpns {
	param(
		[string[]]$UPNs,
		[System.Windows.Forms.Label]$StatusLabel,
		[System.Windows.Forms.Label]$ProgressLabel,
		[System.Windows.Forms.ProgressBar]$ProgressBar
	)

	$batchSize = 20
	$resolvedUsers = [System.Collections.Generic.List[object]]::new()
	$resolutionResults = [System.Collections.Generic.List[object]]::new()
	$failed = [System.Collections.Generic.List[string]]::new()
	$upns = @($UPNs | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique)
	$successfulCount = 0
	if ($ProgressLabel -and $ProgressBar) { Update-SoaConversionProgress $ProgressLabel $ProgressBar 0 $upns.Count }

	for ($offset = 0; $offset -lt $upns.Count; $offset += $batchSize) {
		$end = [Math]::Min($offset + $batchSize, $upns.Count)
		$pending = [System.Collections.Generic.List[int]]::new()
		$completed = [System.Collections.Generic.HashSet[int]]::new()
		for ($index = $offset; $index -lt $end; $index++) { $pending.Add($index) }
		$attempt = 0

		while ($pending.Count -gt 0 -and $attempt -lt 8) {
			$requests = @(
				for ($slot = 0; $slot -lt $pending.Count; $slot++) {
					$upn = [uri]::EscapeDataString($upns[$pending[$slot]])
					@{
						id = [string]$slot
						method = 'GET'
						url = "/users/$($upn)?`$select=id,userPrincipalName,onPremisesSyncEnabled"
					}
				}
			)
			try {
				$response = Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/v1.0/$batch' `
					-Method POST -Body (@{ requests = $requests } | ConvertTo-Json -Depth 6 -Compress) `
					-ErrorAction Stop
			} catch {
				$attempt++
				$requestError = $_.Exception.Message
				$authorizationFailure = $requestError -match '(?i)forbidden|unauthorized|access denied|status code.*(401|403)'
				if ($attempt -ge 8 -or $authorizationFailure) {
					foreach ($index in $pending) {
						$message = "GET batch failed: $requestError"
						$failed.Add("$($upns[$index]): $message")
						$resolutionResults.Add([pscustomobject]@{ UPN = $upns[$index]; Id = ''; Status = 'Error_GET'; Error = $message })
						[void]$completed.Add($index)
					}
					break
				}
				$StatusLabel.Text = "GET request failed; retry $attempt/8 in 30 sec..."
				[System.Windows.Forms.Application]::DoEvents()
				Start-Sleep -Seconds 30
				continue
			}

			$retryIndices = [System.Collections.Generic.List[int]]::new()
			$retryAfter = 0
			$responded = [System.Collections.Generic.HashSet[int]]::new()
			foreach ($item in $response.responses) {
				$slot = [int]$item.id
				if ($slot -lt 0 -or $slot -ge $pending.Count) { continue }
				$sourceIndex = $pending[$slot]
				[void]$responded.Add($sourceIndex)
				switch ([int]$item.status) {
					200 {
						if ($item.body.onPremisesSyncEnabled -eq $true) {
							$resolvedUsers.Add([pscustomobject]@{ UPN = $upns[$sourceIndex]; Id = $item.body.id; Status = 'Resolved' })
						} else {
							$message = 'The user is not currently synchronized from on-premises Active Directory; SOA conversion was skipped.'
							$failed.Add("$($upns[$sourceIndex]): $message")
							$resolutionResults.Add([pscustomobject]@{ UPN = $upns[$sourceIndex]; Id = $item.body.id; Status = 'NotSynced'; Error = $message })
						}
						[void]$completed.Add($sourceIndex)
					}
					404 {
						$message = 'User not found (404).'
						$failed.Add("$($upns[$sourceIndex]): $message")
						$resolutionResults.Add([pscustomobject]@{ UPN = $upns[$sourceIndex]; Id = ''; Status = 'NotFound_GET'; Error = $message })
						[void]$completed.Add($sourceIndex)
					}
					429 {
						$retryIndices.Add($sourceIndex)
						$delay = 30
						if ($item.headers.'Retry-After') { $delay = [int]$item.headers.'Retry-After' }
						$retryAfter = [Math]::Max($retryAfter, $delay)
					}
					default {
						$message = if ($item.body.error.message) { $item.body.error.message } else { "HTTP $($item.status)" }
						$failed.Add("$($upns[$sourceIndex]): $message")
						$resolutionResults.Add([pscustomobject]@{ UPN = $upns[$sourceIndex]; Id = ''; Status = 'Error_GET'; Error = $message })
						[void]$completed.Add($sourceIndex)
					}
				}
			}

			foreach ($sourceIndex in $pending) {
				if (-not $responded.Contains($sourceIndex) -and $sourceIndex -notin $retryIndices) {
					$retryIndices.Add($sourceIndex)
					$retryAfter = [Math]::Max($retryAfter, 5)
				}
			}

			$pending = $retryIndices
			$attempt++
			if ($pending.Count -gt 0 -and $attempt -lt 8) {
				$waitSeconds = if ($retryAfter -gt 0) { $retryAfter + 5 } else { 5 }
				$StatusLabel.Text = "Resolving UPN to object ID; Graph throttling, retrying in $waitSeconds sec..."
				[System.Windows.Forms.Application]::DoEvents()
				Start-Sleep -Seconds $waitSeconds
			}
		}
		foreach ($index in $pending) {
			if (-not $completed.Contains($index)) {
				$message = 'GET: retry limit exceeded.'
				$failed.Add("$($upns[$index]): $message")
				$resolutionResults.Add([pscustomobject]@{ UPN = $upns[$index]; Id = ''; Status = 'Error_GET'; Error = $message })
			}
		}

		$processed = [Math]::Min($end, $upns.Count)
		$StatusLabel.Text = "Phase 1/2: resolved UPNs for $processed of $($upns.Count)..."
		[System.Windows.Forms.Application]::DoEvents()
	}

	if ($resolvedUsers.Count -eq 0) {
		return [pscustomobject]@{ Successful = 0; Failed = @($failed); Resolved = 0; Results = @() }
	}

	$patchResults = [System.Collections.Generic.List[object]]::new()
	$resolvedArray = $resolvedUsers.ToArray()
	for ($offset = 0; $offset -lt $resolvedArray.Count; $offset += $batchSize) {
		$end = [Math]::Min($offset + $batchSize, $resolvedArray.Count)
		$pending = [System.Collections.Generic.List[int]]::new()
		for ($index = $offset; $index -lt $end; $index++) { $pending.Add($index) }
		$attempt = 0

		while ($pending.Count -gt 0 -and $attempt -lt 8) {
			$requests = @(
				for ($slot = 0; $slot -lt $pending.Count; $slot++) {
					$user = $resolvedArray[$pending[$slot]]
					@{
						id = [string]$slot
						method = 'PATCH'
						url = "/users/$($user.Id)/onPremisesSyncBehavior"
						headers = @{ 'Content-Type' = 'application/json'; Prefer = 'return=representation' }
						body = @{ isCloudManaged = $true }
					}
				}
			)
			try {
				$response = Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/v1.0/$batch' `
					-Method POST -Body (@{ requests = $requests } | ConvertTo-Json -Depth 8 -Compress) `
					-ErrorAction Stop
			} catch {
				$attempt++
				$requestError = $_.Exception.Message
				$authorizationFailure = $requestError -match '(?i)forbidden|unauthorized|access denied|status code.*(401|403)'
				if ($attempt -ge 8 -or $authorizationFailure) {
					foreach ($index in $pending) {
						$user = $resolvedArray[$index]
						$failed.Add("$($user.UPN): PATCH batch failed: $requestError")
						$patchResults.Add([pscustomobject]@{ UPN = $user.UPN; Id = $user.Id; Status = 'Error'; Error = $requestError })
					}
					break
				}
				$StatusLabel.Text = "PATCH request failed; retry $attempt/8 in 30 sec..."
				[System.Windows.Forms.Application]::DoEvents()
				Start-Sleep -Seconds 30
				continue
			}

			$retryIndices = [System.Collections.Generic.List[int]]::new()
			$retryAfter = 0
			$responded = [System.Collections.Generic.HashSet[int]]::new()
			foreach ($item in $response.responses) {
				$slot = [int]$item.id
				if ($slot -lt 0 -or $slot -ge $pending.Count) { continue }
				$sourceIndex = $pending[$slot]
				$user = $resolvedArray[$sourceIndex]
				[void]$responded.Add($sourceIndex)
				if ([int]$item.status -in @(200, 204)) {
					$patchResults.Add([pscustomobject]@{ UPN = $user.UPN; Id = $user.Id; Status = 'OK'; Error = '' })
					$successfulCount++
					if ($ProgressLabel -and $ProgressBar) { Update-SoaConversionProgress $ProgressLabel $ProgressBar $successfulCount $upns.Count }
				} elseif ([int]$item.status -eq 429) {
					$retryIndices.Add($sourceIndex)
					$delay = 30
					if ($item.headers.'Retry-After') { $delay = [int]$item.headers.'Retry-After' }
					$retryAfter = [Math]::Max($retryAfter, $delay)
				} else {
					$message = if ($item.body.error.message) { $item.body.error.message } else { "HTTP $($item.status)" }
					$failed.Add("$($user.UPN): $message")
					$patchResults.Add([pscustomobject]@{ UPN = $user.UPN; Id = $user.Id; Status = 'Error'; Error = $message })
				}
			}

			foreach ($sourceIndex in $pending) {
				if (-not $responded.Contains($sourceIndex) -and $sourceIndex -notin $retryIndices) {
					$retryIndices.Add($sourceIndex)
					$retryAfter = [Math]::Max($retryAfter, 5)
				}
			}

			$pending = $retryIndices
			$attempt++
			if ($pending.Count -gt 0 -and $attempt -lt 8) {
				$waitSeconds = if ($retryAfter -gt 0) { $retryAfter + 5 } else { 5 }
				$StatusLabel.Text = "PATCH isCloudManaged=true; Graph throttling, retrying in $waitSeconds sec..."
				[System.Windows.Forms.Application]::DoEvents()
				Start-Sleep -Seconds $waitSeconds
			}
		}

		foreach ($index in $pending) {
			$user = $resolvedArray[$index]
			$message = 'Retry limit exceeded.'
			$failed.Add("$($user.UPN): $message")
			$patchResults.Add([pscustomobject]@{ UPN = $user.UPN; Id = $user.Id; Status = 'Error'; Error = $message })
		}
		$processed = [Math]::Min($end, $resolvedArray.Count)
		$StatusLabel.Text = "Phase 2/2: updated $processed of $($resolvedArray.Count)..."
		[System.Windows.Forms.Application]::DoEvents()
	}

	$reportPath = Join-Path $env:TEMP "soa_user_results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
	$reportRows = @($resolutionResults.ToArray()) + @($patchResults.ToArray())
	$reportRows | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding UTF8
	return [pscustomobject]@{
		Successful = @($patchResults | Where-Object Status -eq 'OK').Count
		Failed = @($failed)
		Resolved = $resolvedUsers.Count
		Results = $patchResults.ToArray()
		ReportPath = $reportPath
	}
}

function Invoke-SetUsersAccountEnabled {
	param(
		[string[]]$UPNs,
		[bool]$Enabled,
		[System.Windows.Forms.Label]$StatusLabel
	)

	$batchSize = 20
	$maxAttempts = 8
	$targetState = if ($Enabled) { 'enabled' } else { 'disabled' }
	$upns = @($UPNs | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique)
	$resolved = [System.Collections.Generic.List[object]]::new()
	$results = [System.Collections.Generic.List[object]]::new()
	$failed = [System.Collections.Generic.List[string]]::new()

	for ($offset = 0; $offset -lt $upns.Count; $offset += $batchSize) {
		$end = [Math]::Min($offset + $batchSize, $upns.Count)
		$pending = [System.Collections.Generic.List[int]]::new()
		for ($index = $offset; $index -lt $end; $index++) { $pending.Add($index) }
		$attempt = 0

		while ($pending.Count -gt 0 -and $attempt -lt $maxAttempts) {
			$requests = @(
				for ($slot = 0; $slot -lt $pending.Count; $slot++) {
					$escapedUpn = [uri]::EscapeDataString($upns[$pending[$slot]])
					@{
						id = [string]$slot
						method = 'GET'
						url = "/users/$($escapedUpn)?`$select=id,userPrincipalName,accountEnabled"
					}
				}
			)
			try {
				$response = Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/v1.0/$batch' `
					-Method POST -Body (@{ requests = $requests } | ConvertTo-Json -Depth 6 -Compress) `
					-ErrorAction Stop
			} catch {
				if ($_.Exception.Message -match '(?i)forbidden|unauthorized|access denied|status code.*(401|403)') { throw }
				$attempt++
				if ($attempt -ge $maxAttempts) {
					foreach ($index in $pending) {
						$message = "GET batch: $($_.Exception.Message)"
						$failed.Add("$($upns[$index]): $message")
						$results.Add([pscustomobject]@{ UPN = $upns[$index]; ObjectId = ''; Status = 'Error'; Error = $message })
					}
					break
				}
				$StatusLabel.Text = "Odczyt UPN, ponowienie $attempt/$maxAttempts za 30 s..."
				[System.Windows.Forms.Application]::DoEvents()
				Start-Sleep -Seconds 30
				continue
			}

			$retry = [System.Collections.Generic.List[int]]::new()
			$responded = [System.Collections.Generic.HashSet[int]]::new()
			$retryAfter = 0
			foreach ($item in $response.responses) {
				$slot = [int]$item.id
				if ($slot -lt 0 -or $slot -ge $pending.Count) { continue }
				$sourceIndex = $pending[$slot]
				[void]$responded.Add($sourceIndex)
				switch ([int]$item.status) {
					200 {
						if (-not $item.body.id) {
							$message = 'Graph GET did not return an object ID.'
							$failed.Add("$($upns[$sourceIndex]): $message")
							$results.Add([pscustomobject]@{ UPN = $upns[$sourceIndex]; ObjectId = ''; Status = 'Error'; Error = $message })
						} elseif ([bool]$item.body.accountEnabled -eq $Enabled) {
						$results.Add([pscustomobject]@{ UPN = $upns[$sourceIndex]; ObjectId = $item.body.id; Status = 'NoChange'; Error = "The account is already $targetState." })
						} else {
						$resolved.Add([pscustomobject]@{ UPN = $upns[$sourceIndex]; Id = [string]$item.body.id })
						}
					}
					404 {
						$message = 'UPN not found (GET 404).'
						$failed.Add("$($upns[$sourceIndex]): $message")
						$results.Add([pscustomobject]@{ UPN = $upns[$sourceIndex]; ObjectId = ''; Status = 'NotFound'; Error = $message })
					}
					429 {
						$retry.Add($sourceIndex)
						$delay = 10
						if ($item.headers.'Retry-After') { $delay = [int]$item.headers.'Retry-After' }
						$retryAfter = [Math]::Max($retryAfter, $delay)
					}
					default {
						$message = if ($item.body.error.message) { $item.body.error.message } else { "HTTP $($item.status)" }
						$failed.Add("$($upns[$sourceIndex]): GET: $message")
						$results.Add([pscustomobject]@{ UPN = $upns[$sourceIndex]; ObjectId = ''; Status = 'Error'; Error = "GET: $message" })
					}
				}
			}
			foreach ($sourceIndex in $pending) {
				if (-not $responded.Contains($sourceIndex)) { $retry.Add($sourceIndex); $retryAfter = [Math]::Max($retryAfter, 5) }
			}
			$pending = $retry
			$attempt++
			if ($pending.Count -gt 0 -and $attempt -lt $maxAttempts) {
				$waitSeconds = if ($retryAfter -gt 0) { $retryAfter + 5 } else { 5 }
				$StatusLabel.Text = "Graph throttling; retrying GET in $waitSeconds sec..."
				[System.Windows.Forms.Application]::DoEvents()
				Start-Sleep -Seconds $waitSeconds
			}
		}
		foreach ($index in $pending) {
			$message = 'GET: retry limit exceeded.'
			$failed.Add("$($upns[$index]): $message")
			$results.Add([pscustomobject]@{ UPN = $upns[$index]; ObjectId = ''; Status = 'Error'; Error = $message })
		}
		$StatusLabel.Text = "Odczyt UPN: $([Math]::Min($end, $upns.Count))/$($upns.Count)..."
		[System.Windows.Forms.Application]::DoEvents()
	}

	$resolvedUsers = $resolved.ToArray()
	for ($offset = 0; $offset -lt $resolvedUsers.Count; $offset += $batchSize) {
		$end = [Math]::Min($offset + $batchSize, $resolvedUsers.Count)
		$pending = [System.Collections.Generic.List[int]]::new()
		for ($index = $offset; $index -lt $end; $index++) { $pending.Add($index) }
		$attempt = 0

		while ($pending.Count -gt 0 -and $attempt -lt $maxAttempts) {
			$requests = @(
				for ($slot = 0; $slot -lt $pending.Count; $slot++) {
					$user = $resolvedUsers[$pending[$slot]]
					@{
						id = [string]$slot
						method = 'PATCH'
						url = "/users/$($user.Id)"
						headers = @{ 'Content-Type' = 'application/json' }
						body = @{ accountEnabled = $Enabled }
					}
				}
			)
			try {
				$response = Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/v1.0/$batch' `
					-Method POST -Body (@{ requests = $requests } | ConvertTo-Json -Depth 6 -Compress) `
					-ErrorAction Stop
			} catch {
				if ($_.Exception.Message -match '(?i)forbidden|unauthorized|access denied|status code.*(401|403)') { throw }
				$attempt++
				if ($attempt -ge $maxAttempts) {
					foreach ($index in $pending) {
						$user = $resolvedUsers[$index]
						$message = "PATCH batch: $($_.Exception.Message)"
						$failed.Add("$($user.UPN): $message")
						$results.Add([pscustomobject]@{ UPN = $user.UPN; ObjectId = $user.Id; Status = 'Error'; Error = $message })
					}
					break
				}
				$StatusLabel.Text = "PATCH request failed; retry $attempt/$maxAttempts in 30 sec..."
				[System.Windows.Forms.Application]::DoEvents()
				Start-Sleep -Seconds 30
				continue
			}

			$retry = [System.Collections.Generic.List[int]]::new()
			$responded = [System.Collections.Generic.HashSet[int]]::new()
			$retryAfter = 0
			foreach ($item in $response.responses) {
				$slot = [int]$item.id
				if ($slot -lt 0 -or $slot -ge $pending.Count) { continue }
				$sourceIndex = $pending[$slot]
				$user = $resolvedUsers[$sourceIndex]
				[void]$responded.Add($sourceIndex)
				if ([int]$item.status -in @(200, 204)) {
					$results.Add([pscustomobject]@{ UPN = $user.UPN; ObjectId = $user.Id; Status = 'OK'; Error = '' })
				} elseif ([int]$item.status -eq 429) {
					$retry.Add($sourceIndex)
					$delay = 10
					if ($item.headers.'Retry-After') { $delay = [int]$item.headers.'Retry-After' }
					$retryAfter = [Math]::Max($retryAfter, $delay)
				} else {
					$message = if ($item.body.error.message) { $item.body.error.message } else { "HTTP $($item.status)" }
					$failed.Add("$($user.UPN): PATCH: $message")
					$results.Add([pscustomobject]@{ UPN = $user.UPN; ObjectId = $user.Id; Status = 'Error'; Error = "PATCH: $message" })
				}
			}
			foreach ($sourceIndex in $pending) {
				if (-not $responded.Contains($sourceIndex)) { $retry.Add($sourceIndex); $retryAfter = [Math]::Max($retryAfter, 5) }
			}
			$pending = $retry
			$attempt++
			if ($pending.Count -gt 0 -and $attempt -lt $maxAttempts) {
				$waitSeconds = if ($retryAfter -gt 0) { $retryAfter + 5 } else { 5 }
				$StatusLabel.Text = "Graph throttling; retrying PATCH in $waitSeconds sec..."
				[System.Windows.Forms.Application]::DoEvents()
				Start-Sleep -Seconds $waitSeconds
			}
		}
		foreach ($index in $pending) {
			$user = $resolvedUsers[$index]
			$message = 'PATCH: retry limit exceeded.'
			$failed.Add("$($user.UPN): $message")
			$results.Add([pscustomobject]@{ UPN = $user.UPN; ObjectId = $user.Id; Status = 'Error'; Error = $message })
		}
		$StatusLabel.Text = "Zmiana stanu kont: $([Math]::Min($end, $resolvedUsers.Count))/$($resolvedUsers.Count)..."
		[System.Windows.Forms.Application]::DoEvents()
	}

	$reportPath = Join-Path $env:TEMP "user_account_enabled_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
	$results | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding UTF8
	return [pscustomobject]@{
		Successful = @($results | Where-Object Status -eq 'OK').Count
		NoChange = @($results | Where-Object Status -eq 'NoChange').Count
		Failed = @($failed)
		Resolved = $upns.Count
		ReportPath = $reportPath
	}
}

function Get-UsersForDomain {
	param([string]$Domain)

	$nextUri = 'https://graph.microsoft.com/v1.0/users?$select=id,userPrincipalName,userType,onPremisesSyncEnabled&$top=999'
	$users = [System.Collections.Generic.List[object]]::new()
	while ($nextUri) {
		$page = Invoke-MgGraphRequest -Uri $nextUri -Method GET -ErrorAction Stop
		foreach ($user in $page.value) {
			if ($user.userType -eq 'Member' -and $user.onPremisesSyncEnabled -eq $true -and $user.userPrincipalName -and $user.userPrincipalName.EndsWith($Domain, [StringComparison]::OrdinalIgnoreCase)) {
				$users.Add($user)
			}
		}
		$nextUri = $page.'@odata.nextLink'
	}

	return $users.ToArray()
}

function Get-MemberUserUpnsForDomain {
	param([string]$Domain, [System.Windows.Forms.Label]$StatusLabel)

	$nextUri = 'https://graph.microsoft.com/v1.0/users?$select=userPrincipalName,mail,userType,onPremisesSyncEnabled&$top=999'
	$matchedUpns = [System.Collections.Generic.List[string]]::new()
	$scannedCount = 0
	while ($nextUri) {
		$page = Invoke-MgGraphRequest -Uri $nextUri -Method GET -ErrorAction Stop
		foreach ($user in $page.value) {
			$scannedCount++
			if ($user.userType -ne 'Member' -or $user.onPremisesSyncEnabled -ne $true -or -not $user.userPrincipalName) { continue }
			$upnMatches = ([string]$user.userPrincipalName).EndsWith($Domain, [StringComparison]::OrdinalIgnoreCase)
			$mailMatches = $user.mail -and ([string]$user.mail).EndsWith($Domain, [StringComparison]::OrdinalIgnoreCase)
			if ($upnMatches -or $mailMatches) { $matchedUpns.Add([string]$user.userPrincipalName) }
		}
		$StatusLabel.Text = "Scanned $scannedCount users; found $($matchedUpns.Count)..."
		[System.Windows.Forms.Application]::DoEvents()
		$nextUri = $page.'@odata.nextLink'
	}

	return @($matchedUpns | Sort-Object -Unique)
}

function Get-EntraGroupsForDomainAndSid {
	param(
		[string]$Domain,
		[string]$DomainSid,
		[string]$ExactMail,
		[bool]$RestrictMailToDomain,
		[System.Windows.Forms.Label]$StatusLabel
	)

	$nextUri = 'https://graph.microsoft.com/v1.0/groups?$select=id,displayName,mail,onPremisesSecurityIdentifier&$top=999'
	$matchedGroups = [System.Collections.Generic.List[object]]::new()
	$scannedCount = 0
	$sidPrefix = "$DomainSid-"
	while ($nextUri) {
		$page = Invoke-MgGraphRequest -Uri $nextUri -Method GET -ErrorAction Stop
		foreach ($group in $page.value) {
			$scannedCount++
			$groupSid = [string]$group.onPremisesSecurityIdentifier
			$groupMail = [string]$group.mail
			if (-not $group.id -or -not $groupSid.StartsWith($sidPrefix, [StringComparison]::OrdinalIgnoreCase)) { continue }
			if ($groupSid.Substring($sidPrefix.Length) -notmatch '^\d+$') { continue }
			if ($RestrictMailToDomain -and -not $groupMail.EndsWith($Domain, [StringComparison]::OrdinalIgnoreCase)) { continue }
			if ($ExactMail -and $groupMail -ine $ExactMail) { continue }
			$matchedGroups.Add([pscustomobject]@{
				GroupId = [string]$group.id
				DisplayName = [string]$group.displayName
				Mail = $groupMail
				OnPremisesSecurityIdentifier = $groupSid
			})
		}
		$StatusLabel.Text = "Odczytano $scannedCount grup; znaleziono $($matchedGroups.Count)..."
		[System.Windows.Forms.Application]::DoEvents()
		$nextUri = $page.'@odata.nextLink'
	}

	return @($matchedGroups | Sort-Object DisplayName, GroupId)
}

function Show-OperationResult {
	param($Form, $StatusLabel, $Result)

	$summary = "Completed. UPNs resolved: $($Result.Resolved); successful: $($Result.Successful); errors: $($Result.Failed.Count)."
	$StatusLabel.Text = $summary
	if ($Result.ReportPath) { $summary += "`r`nRaport CSV: $($Result.ReportPath)" }
	if ($Result.Failed.Count -gt 0) {
		$details = ($Result.Failed | Select-Object -First 12) -join "`r`n"
		[System.Windows.Forms.MessageBox]::Show(
			$Form,
			"$summary`r`n`r`nFirst errors:`r`n$details",
			'SOA Result',
			[System.Windows.Forms.MessageBoxButtons]::OK,
			[System.Windows.Forms.MessageBoxIcon]::Warning
		) | Out-Null
	} else {
		[System.Windows.Forms.MessageBox]::Show($Form, $summary, 'SOA Result') | Out-Null
	}
}

function Show-CsvUsersForm {
	$domain = $appState.SelectedDomain
	$form = New-ActionForm 'Users from CSV' 250
	Add-FormLabel $form 'Select a CSV file with a UPN or UserPrincipalName column:' 20 20 | Out-Null
	$pathBox = Add-FormTextBox $form 20 50 330
	$browse = Add-FormButton $form 'Browse...' 360 48 100
	$status = Add-FormLabel $form "Target domain: $domain" 20 94 440 25
	$progressBar = Add-FormProgressBar $form 20 126 440
	$progressLabel = Add-FormLabel $form 'Converted 0/0 (0%)' 20 148 440 22
	$run = Add-FormButton $form 'Convert Users to SOA' 20 178 190
	$browse.Add_Click({
		$dialog = [System.Windows.Forms.OpenFileDialog]::new()
		$dialog.Filter = 'Pliki CSV (*.csv)|*.csv|Wszystkie pliki (*.*)|*.*'
		if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) { $pathBox.Text = $dialog.FileName }
	}.GetNewClosure())
	$run.Add_Click({
		try {
			if (-not (Test-Path -LiteralPath $pathBox.Text)) { throw 'Select an existing CSV file.' }
			$null = Get-GraphContext
			$rows = @(Import-Csv -LiteralPath $pathBox.Text -ErrorAction Stop)
			$column = Get-CsvColumn $rows @('UserPrincipalName', 'UPN', 'userPrincipalName', 'Email') 'UPN'
			$ids = @($rows | ForEach-Object { ([string]$_.$column).Trim() } | Where-Object { $_ -and $_.EndsWith($domain, [StringComparison]::OrdinalIgnoreCase) } | Sort-Object -Unique)
			if ($ids.Count -eq 0) { throw "The CSV contains no UPNs ending in $domain." }
			$targetDescription = "$($ids.Count) users in domain $domain from CSV file: $($pathBox.Text)"
			if (-not (Confirm-SoaConversion $form $targetDescription)) { return }
			$result = Invoke-SoaUsersFromUpns -UPNs $ids -StatusLabel $status -ProgressLabel $progressLabel -ProgressBar $progressBar
			Show-OperationResult $form $status $result
		} catch {
			$status.Text = 'The operation failed.'
			[System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'Error', 'OK', 'Error') | Out-Null
		}
	}.GetNewClosure())
	$form.ShowDialog($mainForm) | Out-Null
}

function Show-ExportMemberUsersForm {
	$domain = $appState.SelectedDomain
	$form = New-ActionForm 'Export Member User UPNs' 190
	Add-FormLabel $form "Exports Member users whose UPN or mail ends with $domain." 20 20 440 42 | Out-Null
	$status = Add-FormLabel $form 'The output CSV will contain a UserPrincipalName column.' 20 70 440 35
	$export = Add-FormButton $form 'Browse and Export CSV' 20 120 190
	$export.Add_Click({
		try {
			$null = Get-GraphContext
			$dialog = [System.Windows.Forms.SaveFileDialog]::new()
			$dialog.Filter = 'CSV files (*.csv)|*.csv'
			$dialog.FileName = "users_$($domain.TrimStart('@'))_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
			if ($dialog.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) { return }

			$export.Enabled = $false
			$upns = @(Get-MemberUserUpnsForDomain -Domain $domain -StatusLabel $status)
			if ($upns.Count -eq 0) { throw "No Member users found with a UPN or mail address ending in $domain." }
			$upns |
				ForEach-Object { [pscustomobject]@{ UserPrincipalName = $_ } } |
				Export-Csv -LiteralPath $dialog.FileName -NoTypeInformation -Encoding UTF8
			$status.Text = "Saved $($upns.Count) users."
			[System.Windows.Forms.MessageBox]::Show($form, "Saved $($upns.Count) UPNs to:`r`n$($dialog.FileName)", 'Export Complete', 'OK', 'Information') | Out-Null
		} catch {
			$status.Text = 'The export failed.'
			[System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'Export Error', 'OK', 'Error') | Out-Null
		} finally {
			$export.Enabled = $true
		}
	}.GetNewClosure())
	$form.ShowDialog($mainForm) | Out-Null
}

function Show-ExportEntraGroupsForm {
	$domain = $appState.SelectedDomain
	$form = New-ActionForm 'Get Groups from EntraID' 280
	Add-FormLabel $form "Selected domain: $domain. Enter the domain SID without the final RID:" 20 18 440 35 | Out-Null
	$sidBox = Add-FormTextBox $form 20 55 420
	$domainFilter = [System.Windows.Forms.CheckBox]::new()
	$domainFilter.Text = "Restrict results to mail addresses ending in $domain"
	$domainFilter.Location = [System.Drawing.Point]::new(20, 88)
	$domainFilter.Size = [System.Drawing.Size]::new(440, 24)
	$domainFilter.Checked = $false
	$form.Controls.Add($domainFilter)
	Add-FormLabel $form 'Exact mail address (optional):' 20 118 440 22 | Out-Null
	$mailBox = Add-FormTextBox $form 20 142 420
	$status = Add-FormLabel $form '' 20 174 440 30
	$export = Add-FormButton $form 'Browse and Export CSV' 20 210 190
	$export.Add_Click({
		try {
			$domainSid = $sidBox.Text.Trim().TrimEnd('-')
			if ($domainSid -notmatch '^S-1-5-21-\d+-\d+-\d+$') {
				throw 'Enter the domain SID in the format S-1-5-21-<number>-<number>-<number>, without the group RID.'
			}
			$exactMail = $mailBox.Text.Trim()
			if ($domainFilter.Checked -and $exactMail -and -not $exactMail.EndsWith($domain, [StringComparison]::OrdinalIgnoreCase)) {
				throw "The mail address must end in the selected domain $domain."
			}
			$null = Get-GraphContext
			$dialog = [System.Windows.Forms.SaveFileDialog]::new()
			$dialog.Filter = 'CSV files (*.csv)|*.csv'
			$dialog.FileName = "groups_$($domain.TrimStart('@'))_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
			if ($dialog.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) { return }

			$export.Enabled = $false
			$groups = @(Get-EntraGroupsForDomainAndSid -Domain $domain -DomainSid $domainSid -ExactMail $exactMail -RestrictMailToDomain $domainFilter.Checked -StatusLabel $status)
			if ($groups.Count -eq 0) { throw 'No groups matched the SID and selected mail filters.' }
			$groups | Export-Csv -LiteralPath $dialog.FileName -NoTypeInformation -Encoding UTF8
			$status.Text = "Saved $($groups.Count) groups."
			[System.Windows.Forms.MessageBox]::Show($form, "Saved $($groups.Count) groups to:`r`n$($dialog.FileName)", 'Export Complete', 'OK', 'Information') | Out-Null
		} catch {
			$status.Text = 'The export failed.'
			[System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'Export Error', 'OK', 'Error') | Out-Null
		} finally {
			$export.Enabled = $true
		}
	}.GetNewClosure())
	$form.ShowDialog($mainForm) | Out-Null
}

function Show-AccountStateActionForm {
	param(
		[bool]$Enabled,
		[bool]$FromCsv
	)

	$domain = $appState.SelectedDomain
	$stateText = if ($Enabled) { 'enabled' } else { 'disabled' }
	$form = New-ActionForm "User Accounts - $stateText" $(if ($FromCsv) { 250 } else { 220 })
	$browse = $null
	if ($FromCsv) {
		Add-FormLabel $form "CSV with a UPN, UserPrincipalName, or Email column; domain: $domain" 20 20 440 40 | Out-Null
		$pathBox = Add-FormTextBox $form 20 65 330
		$browse = Add-FormButton $form 'Browse...' 360 63 100
		$browse.Add_Click({
			$dialog = [System.Windows.Forms.OpenFileDialog]::new()
			$dialog.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
			if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) { $pathBox.Text = $dialog.FileName }
		}.GetNewClosure())
		$buttonY = 145
	} else {
		Add-FormLabel $form "User UPN; must belong to domain $domain" 20 20 440 28 | Out-Null
		$upnBox = Add-FormTextBox $form 20 53 420
		$buttonY = 105
	}
	$status = Add-FormLabel $form '' 20 ($buttonY - 34) 440 28
	$run = Add-FormButton $form "Set account to $stateText" 20 $buttonY 220
	$run.Add_Click({
		try {
			$null = Get-GraphContext
			if ($FromCsv) {
				if (-not (Test-Path -LiteralPath $pathBox.Text)) { throw 'Select an existing CSV file.' }
				$rows = @(Import-Csv -LiteralPath $pathBox.Text -ErrorAction Stop)
				$column = Get-CsvColumn $rows @('UserPrincipalName', 'UPN', 'userPrincipalName', 'Email') 'UPN'
				$upns = @(
					$rows |
						ForEach-Object { ([string]$_.$column).Trim() } |
						Where-Object { $_ -and $_.EndsWith($domain, [StringComparison]::OrdinalIgnoreCase) } |
						Sort-Object -Unique
				)
				if ($upns.Count -eq 0) { throw "The CSV contains no UPNs ending in $domain." }
			} else {
				$upn = $upnBox.Text.Trim()
				if (-not $upn -or -not $upn.EndsWith($domain, [StringComparison]::OrdinalIgnoreCase)) {
					throw "Enter a UPN ending in $domain."
				}
				$upns = @($upn)
			}

			$actionVerb = if ($Enabled) { 'enable' } else { 'disable' }
			$confirmation = [System.Windows.Forms.MessageBox]::Show(
				$form,
				"Are you sure you want to $actionVerb $($upns.Count) user account(s) in domain $domain?",
				'Confirm Change',
				[System.Windows.Forms.MessageBoxButtons]::YesNo,
				[System.Windows.Forms.MessageBoxIcon]::Warning
			)
			if ($confirmation -ne [System.Windows.Forms.DialogResult]::Yes) { return }

			$run.Enabled = $false
			$result = Invoke-SetUsersAccountEnabled -UPNs $upns -Enabled $Enabled -StatusLabel $status
			$summary = "Completed. Changed: $($result.Successful); unchanged: $($result.NoChange); errors: $($result.Failed.Count).`r`nReport: $($result.ReportPath)"
			$status.Text = "Changed: $($result.Successful); unchanged: $($result.NoChange); errors: $($result.Failed.Count)."
			$messageIcon = if ($result.Failed.Count -gt 0) { [System.Windows.Forms.MessageBoxIcon]::Warning } else { [System.Windows.Forms.MessageBoxIcon]::Information }
			[System.Windows.Forms.MessageBox]::Show($form, $summary, 'Account State Update Result', [System.Windows.Forms.MessageBoxButtons]::OK, $messageIcon) | Out-Null
		} catch {
			$status.Text = 'The operation failed.'
			$errorMessage = $_.Exception.Message
			if ($errorMessage -match '(?i)forbidden|insufficient privileges|access denied|status code.*403') {
				$errorMessage += "`r`n`r`nVerify that the signed-in identity has permission to update users. For app-only authentication, configure the required application permissions and grant admin consent. Sign in again after changing permissions."
			}
			[System.Windows.Forms.MessageBox]::Show($form, $errorMessage, 'Account State Update Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
		} finally {
			$run.Enabled = $true
		}
	}.GetNewClosure())
	$form.ShowDialog($mainForm) | Out-Null
}

function Show-AccountStateMenu {
	$form = New-ActionForm 'Enable or Disable User Accounts' 330
	$actions = @(
		@{ Text = 'Disable Single User'; Enabled = $false; FromCsv = $false },
		@{ Text = 'Disable Users from CSV'; Enabled = $false; FromCsv = $true },
		@{ Text = 'Enable Single User'; Enabled = $true; FromCsv = $false },
		@{ Text = 'Enable Users from CSV'; Enabled = $true; FromCsv = $true }
	)
	$top = 20
	foreach ($action in $actions) {
		$button = Add-FormButton $form $action.Text 20 $top 420
		$button.Height = 48
		$enabledState = [bool]$action.Enabled
		$csvMode = [bool]$action.FromCsv
		$button.Add_Click({ Show-AccountStateActionForm -Enabled $enabledState -FromCsv $csvMode }.GetNewClosure())
		$top += 65
	}
	$form.ShowDialog($mainForm) | Out-Null
}

function Show-CsvGroupsForm {
	$form = New-ActionForm 'Groups from CSV' 250
	Add-FormLabel $form 'Select a CSV file with a GroupId, Id, or ObjectId column:' 20 20 | Out-Null
	$pathBox = Add-FormTextBox $form 20 50 330
	$browse = Add-FormButton $form 'Browse...' 360 48 100
	$status = Add-FormLabel $form 'Entra ID group object IDs are required.' 20 94 440 32
	$progressBar = Add-FormProgressBar $form 20 130 440
	$progressLabel = Add-FormLabel $form 'Skonwertowano 0/0 (0%)' 20 152 440 22
	$run = Add-FormButton $form 'Convert Groups to SOA' 20 182 190
	$browse.Add_Click({
		$dialog = [System.Windows.Forms.OpenFileDialog]::new()
		$dialog.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
		if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) { $pathBox.Text = $dialog.FileName }
	}.GetNewClosure())
	$run.Add_Click({
		try {
			if (-not (Test-Path -LiteralPath $pathBox.Text)) { throw 'Select an existing CSV file.' }
			$null = Get-GraphContext
			$rows = @(Import-Csv -LiteralPath $pathBox.Text -ErrorAction Stop)
			$column = Get-CsvColumn $rows @('GroupId', 'Id', 'ObjectId') 'group ID'
			$ids = @($rows | ForEach-Object { ([string]$_.$column).Trim() } | Where-Object { $_ } | Sort-Object -Unique)
			if ($ids.Count -eq 0) { throw 'The CSV contains no group IDs.' }
			$invalidIds = @($ids | Where-Object { $_ -notmatch '^[0-9a-fA-F]{8}-(?:[0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$' })
			if ($invalidIds.Count -gt 0) { throw "Invalid Group Object ID(s): $($invalidIds -join ', ')" }
			$targetDescription = "$($ids.Count) grup z pliku CSV: $($pathBox.Text)"
			if (-not (Confirm-SoaConversion $form $targetDescription)) { return }
			$result = Invoke-SoaBatchPatch -ObjectType 'groups' -ObjectIds $ids -StatusLabel $status -ProgressLabel $progressLabel -ProgressBar $progressBar
			Show-OperationResult $form $status $result
		} catch {
			$status.Text = 'The operation failed.'
			[System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'Error', 'OK', 'Error') | Out-Null
		}
	}.GetNewClosure())
	$form.ShowDialog($mainForm) | Out-Null
}

function Show-SingleObjectForm {
	param([ValidateSet('users', 'groups')][string]$ObjectType)

	$domain = $appState.SelectedDomain
	$noun = if ($ObjectType -eq 'users') { 'User' } else { 'Group' }
	$fieldName = if ($ObjectType -eq 'users') { 'UPN' } else { 'Group Object ID' }
	$fieldHint = if ($ObjectType -eq 'users') { "selected domain: $domain" } else { 'GUID format' }
	$form = New-ActionForm "Convert Single $noun to SOA" 230
	Add-FormLabel $form "$fieldName ($fieldHint):" 20 20 440 25 | Out-Null
	$inputBox = Add-FormTextBox $form 20 52 420
	$status = Add-FormLabel $form '' 20 92 440 25
	$run = Add-FormButton $form 'Convert to SOA' 20 132 190
	$run.Add_Click({
		try {
			$objectId = $inputBox.Text.Trim()
			if (-not $objectId) { throw "Enter $fieldName." }
			if ($ObjectType -eq 'users' -and -not $objectId.EndsWith($domain, [StringComparison]::OrdinalIgnoreCase)) {
				throw "The UPN must end in $domain."
			}
			if ($ObjectType -eq 'groups' -and $objectId -notmatch '^[0-9a-fA-F]{8}-(?:[0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') {
				throw 'Enter a valid Group Object ID in GUID format.'
			}
			$null = Get-GraphContext
			$targetDescription = if ($ObjectType -eq 'users') { "user $objectId" } else { "group $objectId" }
			if (-not (Confirm-SoaConversion $form $targetDescription)) { return }
			if ($ObjectType -eq 'users') {
				$result = Invoke-SoaUsersFromUpns @($objectId) $status
			} else {
				$result = Invoke-SoaBatchPatch $ObjectType @($objectId) $status
			}
			Show-OperationResult $form $status $result
		} catch {
			$status.Text = 'The operation failed.'
			[System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'Error', 'OK', 'Error') | Out-Null
		}
	}.GetNewClosure())
	$form.ShowDialog($mainForm) | Out-Null
}

function Show-AllDomainForm {
	$domain = $appState.SelectedDomain
	$form = New-ActionForm 'Convert All Users to SOA' 270
	Add-FormLabel $form "Selected domain: $domain" 20 20 440 25 | Out-Null
	Add-FormLabel $form 'Preview the number of eligible users before confirming the conversion.' 20 52 440 42 | Out-Null
	$status = Add-FormLabel $form '' 20 103 440 42
	$preview = Add-FormButton $form 'Load Preview' 20 158 150
	$execute = Add-FormButton $form 'Convert All Users' 190 158 180
	$execute.Enabled = $false
	$form.Tag = @{}
	$preview.Add_Click({
		try {
			$null = Get-GraphContext
			$status.Text = 'Loading users from Microsoft Graph...'
			[System.Windows.Forms.Application]::DoEvents()
			$domainUsers = @(Get-UsersForDomain $domain)
			$form.Tag.Users = @($domainUsers | ForEach-Object { $_.userPrincipalName })
			$status.Text = "Found $($domainUsers.Count) eligible users in domain $domain."
			$execute.Enabled = $domainUsers.Count -gt 0
		} catch {
			$status.Text = 'Failed to load the user list.'
			[System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'Error', 'OK', 'Error') | Out-Null
		}
	}.GetNewClosure())
	$execute.Add_Click({
		try {
			$ids = @($form.Tag.Users)
			if ($ids.Count -eq 0) { throw 'Load the preview first.' }
			$targetDescription = "$($ids.Count) users in domain $domain"
			if (-not (Confirm-SoaConversion $form $targetDescription)) { return }
			$execute.Enabled = $false
			$result = Invoke-SoaBatchPatch 'users' $ids $status
			Show-OperationResult $form $status $result
		} catch {
			[System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'Error', 'OK', 'Error') | Out-Null
		} finally { $execute.Enabled = $true }
	}.GetNewClosure())
	$form.ShowDialog($mainForm) | Out-Null
}

function Show-PermissionForm {
	$domain = $appState.SelectedDomain
	$form = New-ActionForm 'Validate Graph Connection and Permissions' 270
	Add-FormLabel $form 'Test user UPN (read-only):' 20 20 | Out-Null
	$upnBox = Add-FormTextBox $form 20 50 420
	$status = Add-FormLabel $form 'This test does not make any changes.' 20 90 440 40
	$check = Add-FormButton $form 'Check Access' 20 145 150
	$check.Add_Click({
		try {
			$upn = $upnBox.Text.Trim()
			if (-not $upn.EndsWith($domain, [StringComparison]::OrdinalIgnoreCase)) { throw "Enter a UPN ending in $domain." }
			$context = Get-GraphContext
			$escapedUpn = [uri]::EscapeDataString($upn)
			$state = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/users/$escapedUpn/onPremisesSyncBehavior?`$select=isCloudManaged" -Method GET -ErrorAction Stop
			$status.Text = "Read succeeded. Tenant: $($context.TenantId); isCloudManaged: $($state.isCloudManaged)."
			[System.Windows.Forms.MessageBox]::Show($form, "Read access to the SOA endpoint succeeded.`r`nTenant: $($context.TenantId)`r`nAuthentication type: $($context.AuthType)`r`nApplication: $($context.AppName)`r`n`r`nThis confirms read access only; it does not guarantee write permission.", 'Validation Result', 'OK', 'Information') | Out-Null
		} catch {
			$errorMessage = $_.Exception.Message
			$statusCode = $null
			if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
				$statusCode = [int]$_.Exception.Response.StatusCode
			}

			if ($statusCode -eq 403 -or $errorMessage -match '(?i)forbidden|access denied|status code.*403') {
				$context = Get-MgContext
				$identity = if ($context.Account) { $context.Account } elseif ($context.AppName) { "application $($context.AppName)" } else { 'unknown identity' }
				$grantedScopes = if ($context.Scopes) { $context.Scopes -join ', ' } else { 'no delegated scopes in the context' }
				$status.Text = "Signed in as $identity, but access to the SOA endpoint was denied."
				$guidance = @"
The Microsoft Graph connection is active, but the SOA endpoint returned HTTP 403 Forbidden.
Identity: $identity
Tenant: $($context.TenantId)
Scopes in context: $grantedScopes

For app-only sign-in, configure the Application permissions User-OnPremisesSyncBehavior.ReadWrite.All and User.Read.All, then grant admin consent.
For delegated sign-in, the appropriate Microsoft Graph delegated permission and administrator consent are required.
Connect-MgGraph without -Scopes authenticates the user but does not grant access to the SOA endpoint.

Microsoft Graph response: $errorMessage
"@
				[System.Windows.Forms.MessageBox]::Show($form, $guidance, 'Microsoft Graph Permission Denied', 'OK', 'Warning') | Out-Null
			} else {
				$status.Text = 'Validation failed.'
				[System.Windows.Forms.MessageBox]::Show($form, $errorMessage, 'Validation Error', 'OK', 'Error') | Out-Null
			}
		}
	}.GetNewClosure())
	$form.ShowDialog($mainForm) | Out-Null
}

function Show-GroupPermissionForm {
	$form = New-ActionForm 'Validate Group Permissions' 250
	Add-FormLabel $form 'Group Object ID (read-only test):' 20 20 440 25 | Out-Null
	$groupIdBox = Add-FormTextBox $form 20 50 420
	$status = Add-FormLabel $form 'Test verifies read access only; it does not change the group.' 20 90 440 40
	$check = Add-FormButton $form 'Validate Group Permissions' 20 145 230
	$check.Add_Click({
		try {
			$groupId = $groupIdBox.Text.Trim()
			if ($groupId -notmatch '^[0-9a-fA-F]{8}-(?:[0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') {
				throw 'Enter a valid Group Object ID in GUID format.'
			}
			$context = Get-GraphContext
			$escapedGroupId = [uri]::EscapeDataString($groupId)
			$groupState = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/groups/$escapedGroupId/onPremisesSyncBehavior?`$select=isCloudManaged" -Method GET -ErrorAction Stop
			$identity = if ($context.Account) { $context.Account } elseif ($context.AppName) { "application $($context.AppName)" } else { 'unknown identity' }
			$status.Text = "SOA endpoint read succeeded. isCloudManaged: $($groupState.isCloudManaged)."
			$message = "Read access to the group SOA endpoint succeeded.`r`nGroup ID: $groupId`r`nisCloudManaged: $($groupState.isCloudManaged)`r`nIdentity: $identity`r`nTenant: $($context.TenantId)`r`nAuthentication: $($context.AuthType)`r`n`r`nThis confirms read access only, not permission to modify the group."
			[System.Windows.Forms.MessageBox]::Show($form, $message, 'Group permission validation', 'OK', 'Information') | Out-Null
		} catch {
			$errorMessage = $_.Exception.Message
			if ($errorMessage -match '(?i)forbidden|access denied|status code.*403') {
				$context = Get-MgContext
				$identity = if ($context.Account) { $context.Account } elseif ($context.AppName) { "application $($context.AppName)" } else { 'unknown identity' }
				$scopes = if ($context.Scopes) { $context.Scopes -join ', ' } else { 'no delegated scopes in context' }
				$errorMessage += "`r`n`r`nIdentity: $identity`r`nTenant: $($context.TenantId)`r`nScopes: $scopes`r`nCheck that this identity has permission to read groups."
			}
			$status.Text = 'Group permission validation failed.'
			[System.Windows.Forms.MessageBox]::Show($form, $errorMessage, 'Group permission validation failed', 'OK', 'Error') | Out-Null
		}
	}.GetNewClosure())
	$form.ShowDialog($mainForm) | Out-Null
}

function Show-CheckSingleSoaStatusForm {
	param([ValidateSet('users', 'groups')][string]$ObjectType)

	$isUser = $ObjectType -eq 'users'
	$objectLabel = if ($isUser) { 'user UPN' } else { 'Group Object ID' }
	$form = New-ActionForm "Check SOA status for single $ObjectType" 230
	Add-FormLabel $form "Enter ${objectLabel}:" 20 20 440 25 | Out-Null
	$identifierBox = Add-FormTextBox $form 20 50 420
	$status = Add-FormLabel $form 'Checks isCloudManaged without changing the object.' 20 90 440 35
	$check = Add-FormButton $form 'Check SOA Status' 20 135 190
	$check.Add_Click({
		try {
			$identifier = $identifierBox.Text.Trim()
			if ($isUser -and $identifier -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
				throw 'Enter a valid user UPN.'
			}
			if (-not $isUser -and $identifier -notmatch '^[0-9a-fA-F]{8}-(?:[0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') {
				throw 'Enter a valid Group Object ID in GUID format.'
			}
			$null = Get-GraphContext
			$result = @(Invoke-SoaStatusBatch -ObjectType $ObjectType -ObjectIds @($identifier) -StatusLabel $status)[0]
			$status.Text = "$($result.Status): isCloudManaged=$($result.IsCloudManaged)"
			$message = if ($result.Error) { "$($result.Status): $($result.Error)" } else { "$($result.Status)`r`n${objectLabel}: $identifier`r`nisCloudManaged: $($result.IsCloudManaged)" }
			$icon = if ($result.Status -in @('Error', 'NotFound')) { [System.Windows.Forms.MessageBoxIcon]::Warning } else { [System.Windows.Forms.MessageBoxIcon]::Information }
			[System.Windows.Forms.MessageBox]::Show($form, $message, 'SOA status', [System.Windows.Forms.MessageBoxButtons]::OK, $icon) | Out-Null
		} catch {
			$status.Text = 'SOA status check failed.'
			[System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'SOA status check failed', 'OK', 'Error') | Out-Null
		}
	}.GetNewClosure())
	$form.ShowDialog($mainForm) | Out-Null
}

function Show-CheckSingleUserSoaStatusForm {
	Show-CheckSingleSoaStatusForm -ObjectType 'users'
}

function Show-CheckSingleGroupSoaStatusForm {
	Show-CheckSingleSoaStatusForm -ObjectType 'groups'
}

function Show-CheckSoaStatusCsvForm {
	param([ValidateSet('users', 'groups')][string]$ObjectType)

	$isUser = $ObjectType -eq 'users'
	$objectLabel = if ($isUser) { 'users' } else { 'groups' }
	$form = New-ActionForm "Check SOA status for $objectLabel from CSV" 270
	$columnHint = if ($isUser) { 'UPN, UserPrincipalName, or Email' } else { 'GroupId, Id, or ObjectId' }
	Add-FormLabel $form "Select CSV with a $columnHint column:" 20 20 440 35 | Out-Null
	$pathBox = Add-FormTextBox $form 20 57 330
	$browse = Add-FormButton $form 'Browse...' 360 55 100
	$status = Add-FormLabel $form 'A CSV report will be generated for every checked object.' 20 95 440 38
	$progressBar = Add-FormProgressBar $form 20 140 440
	$progressLabel = Add-FormLabel $form 'Checked 0/0 (0%)' 20 162 440 22
	$run = Add-FormButton $form 'Check SOA Status' 20 192 190
	$browse.Add_Click({
		$dialog = [System.Windows.Forms.OpenFileDialog]::new()
		$dialog.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
		if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) { $pathBox.Text = $dialog.FileName }
	}.GetNewClosure())
	$run.Add_Click({
		try {
			if (-not (Test-Path -LiteralPath $pathBox.Text)) { throw 'Select an existing CSV file.' }
			$null = Get-GraphContext
			$rows = @(Import-Csv -LiteralPath $pathBox.Text -ErrorAction Stop)
			$candidates = if ($isUser) { @('UserPrincipalName', 'UPN', 'userPrincipalName', 'Email') } else { @('GroupId', 'Id', 'ObjectId') }
			$column = Get-CsvColumn $rows $candidates $(if ($isUser) { 'user UPN' } else { 'group ID' })
			$objectIds = @($rows | ForEach-Object { ([string]$_.$column).Trim() } | Where-Object { $_ } | Sort-Object -Unique)
			if ($objectIds.Count -eq 0) { throw 'CSV contains no object identifiers.' }

			$saveDialog = [System.Windows.Forms.SaveFileDialog]::new()
			$saveDialog.Filter = 'CSV files (*.csv)|*.csv'
			$saveDialog.FileName = "soa_${objectLabel}_status_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
			if ($saveDialog.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) { return }

			$run.Enabled = $false
			$results = @(Invoke-SoaStatusBatch -ObjectType $ObjectType -ObjectIds $objectIds -StatusLabel $progressLabel -ProgressBar $progressBar)
			$report = if ($isUser) {
				$results | Select-Object @{ Name = 'UserPrincipalName'; Expression = { $_.Identifier } }, IsCloudManaged, Status, Error
			} else {
				$results | Select-Object @{ Name = 'GroupId'; Expression = { $_.Identifier } }, IsCloudManaged, Status, Error
			}
			$report | Export-Csv -LiteralPath $saveDialog.FileName -NoTypeInformation -Encoding UTF8
			$converted = @($results | Where-Object Status -eq 'ConvertedToSOA').Count
			$status.Text = "Checked $($results.Count) $objectLabel; converted to SOA: $converted."
			[System.Windows.Forms.MessageBox]::Show($form, "$($status.Text)`r`nReport: $($saveDialog.FileName)", 'SOA status report', 'OK', 'Information') | Out-Null
		} catch {
			$status.Text = 'SOA status check failed.'
			[System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'SOA status check failed', 'OK', 'Error') | Out-Null
		} finally {
			$run.Enabled = $true
		}
	}.GetNewClosure())
	$form.ShowDialog($mainForm) | Out-Null
}

function Show-CheckGroupsFromCsvSoaStatusForm {
	Show-CheckSoaStatusCsvForm -ObjectType 'groups'
}

function Show-CheckUsersFromCsvSoaStatusForm {
	Show-CheckSoaStatusCsvForm -ObjectType 'users'
}

$mainForm = [System.Windows.Forms.Form]::new()
$mainForm.Text = 'SOA Converter'
$mainForm.StartPosition = 'CenterScreen'
$mainForm.ClientSize = [System.Drawing.Size]::new(760, 630)
$mainForm.FormBorderStyle = 'FixedDialog'
$mainForm.MaximizeBox = $false
$mainForm.BackColor = [System.Drawing.Color]::White

$domainLabel = [System.Windows.Forms.Label]::new()
$domainLabel.Text = 'Select Domain'
$domainLabel.Location = [System.Drawing.Point]::new(20, 18)
$domainLabel.Size = [System.Drawing.Size]::new(720, 24)
$domainLabel.Font = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$domainLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 112, 192)
$mainForm.Controls.Add($domainLabel)

$domainPicker = [System.Windows.Forms.ComboBox]::new()
$domainPicker.Location = [System.Drawing.Point]::new(20, 48)
$domainPicker.Size = [System.Drawing.Size]::new(420, 30)
$domainPicker.DropDownStyle = 'DropDownList'
$domainPicker.Font = [System.Drawing.Font]::new('Segoe UI', 10)
$domainPicker.ForeColor = [System.Drawing.Color]::FromArgb(0, 112, 192)
foreach ($domainValue in $appState.Domains) { [void]$domainPicker.Items.Add($domainValue) }
$domainPicker.SelectedItem = $appState.SelectedDomain
$domainPicker.Add_SelectedIndexChanged({ $appState.SelectedDomain = [string]$domainPicker.SelectedItem }.GetNewClosure())
$mainForm.Controls.Add($domainPicker)

$addDomainButton = Add-FormButton $mainForm 'Add Domain' 455 46 120
$addDomainButton.Add_Click({
	$form = New-ActionForm 'Add Domain' 190
	Add-FormLabel $form 'Enter an external or internal domain:' 20 20 | Out-Null
	$domainBox = Add-FormTextBox $form 20 50 420
	$saveButton = Add-FormButton $form 'Add Domain' 20 95 150
	$saveButton.Add_Click({
		try {
			$normalizedDomain = $domainBox.Text.Trim().TrimStart('@').TrimEnd('.').ToLowerInvariant()
			if ($normalizedDomain -notmatch '^(?=.{1,253}$)(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,63}$') {
				throw 'Enter a valid domain, for example, example.com.'
			}
			$domainWithAt = "@$normalizedDomain"
			if ($appState.Domains.Contains($domainWithAt)) { throw 'This domain is already in the list.' }
			$appState.Domains.Add($domainWithAt)
			Save-Domains $appState
			$domainPicker.Items.Clear()
			foreach ($domainValue in $appState.Domains) { [void]$domainPicker.Items.Add($domainValue) }
			$domainPicker.SelectedItem = $domainWithAt
			$form.Close()
		} catch {
			[System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'Domain Not Added', 'OK', 'Warning') | Out-Null
		}
	}.GetNewClosure())
	$form.ShowDialog($mainForm) | Out-Null
}.GetNewClosure())

$removeDomainButton = Add-FormButton $mainForm 'Remove Domain' 585 46 155
$removeDomainButton.BackColor = [System.Drawing.Color]::FromArgb(184, 35, 45)
$removeDomainButton.FlatAppearance.BorderColor = $removeDomainButton.BackColor
$removeDomainButton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(206, 57, 67)
$removeDomainButton.Add_Click({
	$selectedDomain = [string]$domainPicker.SelectedItem
	if (-not $selectedDomain) {
		[System.Windows.Forms.MessageBox]::Show($mainForm, 'Select a domain to remove.', 'No Domain Selected', 'OK', 'Information') | Out-Null
		return
	}

	$confirmation = [System.Windows.Forms.MessageBox]::Show(
		$mainForm,
		"Are you sure you want to remove $selectedDomain from the saved list?",
		'Confirm Domain Removal',
		[System.Windows.Forms.MessageBoxButtons]::YesNo,
		[System.Windows.Forms.MessageBoxIcon]::Warning,
		[System.Windows.Forms.MessageBoxDefaultButton]::Button2
	)
	if ($confirmation -ne [System.Windows.Forms.DialogResult]::Yes) { return }

	[void]$appState.Domains.Remove($selectedDomain)
	Save-Domains $appState
	$domainPicker.Items.Clear()
	foreach ($domainValue in $appState.Domains) { [void]$domainPicker.Items.Add($domainValue) }
	if ($domainPicker.Items.Count -gt 0) {
		$domainPicker.SelectedIndex = 0
	} else {
		$appState.SelectedDomain = $null
	}
}.GetNewClosure())

$connectButton = Add-FormButton $mainForm 'Connect Microsoft Graph' 20 91 720
$connectButton.Height = 43
$connectButton.Font = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$connectButton.Add_Click({
	$form = New-ActionForm 'Connect to Microsoft Graph' 340
	$tabs = [System.Windows.Forms.TabControl]::new()
	$tabs.Location = [System.Drawing.Point]::new(15, 15)
	$tabs.Size = [System.Drawing.Size]::new(450, 235)
	$userPage = [System.Windows.Forms.TabPage]::new('User (UPN)')
	$appPage = [System.Windows.Forms.TabPage]::new('Application')
	$userPage.AutoScroll = $true
	$appPage.AutoScroll = $true
	$tabs.TabPages.Add($userPage)
	$tabs.TabPages.Add($appPage)
	Add-FormLabel $userPage 'User account UPN:' 12 20 | Out-Null
	$accountBox = Add-FormTextBox $userPage 12 48 405
	Add-FormLabel $userPage 'Interactive sign-in using delegated permissions.' 12 88 405 42 | Out-Null
	Add-FormLabel $appPage 'Tenant ID:' 12 10 | Out-Null
	$tenantBox = Add-FormTextBox $appPage 12 35 405
	Add-FormLabel $appPage 'Application (client) ID:' 12 68 | Out-Null
	$clientBox = Add-FormTextBox $appPage 12 93 405
	Add-FormLabel $appPage 'Client secret:' 12 126 | Out-Null
	$secretBox = Add-FormTextBox $appPage 12 151 405
	$secretBox.UseSystemPasswordChar = $true
	Add-FormLabel $appPage 'The client secret is not saved.' 12 179 405 22 | Out-Null
	$form.Controls.Add($tabs)
	$login = Add-FormButton $form 'Connect' 20 260 130
	$login.Add_Click({
		$secureSecret = $null
		try {
			if (-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)) {
				Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
			}
			Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
			if ($tabs.SelectedTab -eq $userPage) {
				$expectedAccount = $accountBox.Text.Trim()
				if (-not $expectedAccount -or $expectedAccount -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { throw 'Enter a valid UPN.' }
				Connect-MgGraph -Scopes 'User.ReadWrite.All', 'Directory.ReadWrite.All', 'Group.ReadWrite.All' -ContextScope Process -NoWelcome -ErrorAction Stop | Out-Null
				$context = Get-MgContext
				if (-not $context.Account -or $context.Account -ine $expectedAccount) {
					Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
					throw "Signed in as '$($context.Account)', but expected '$expectedAccount'. Sign in again with the correct account."
				}
				$message = "Connected as $($context.Account) to tenant $($context.TenantId)."
			} else {
				$tenantId = $tenantBox.Text.Trim()
				$clientId = $clientBox.Text.Trim()
				if (-not $tenantId -or -not $clientId -or -not $secretBox.Text) { throw 'Enter the Tenant ID, Client ID, and client secret.' }
				$secureSecret = ConvertTo-SecureString $secretBox.Text -AsPlainText -Force
				$credential = [System.Management.Automation.PSCredential]::new($clientId, $secureSecret)
				Connect-MgGraph -TenantId $tenantId -ClientSecretCredential $credential -ContextScope Process -NoWelcome -ErrorAction Stop | Out-Null
				$context = Get-MgContext
				$message = "Connected application $($context.AppName) to tenant $($context.TenantId)."
			}
			[System.Windows.Forms.MessageBox]::Show($form, $message, 'Microsoft Graph') | Out-Null
			$form.Close()
		} catch {
			[System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'Sign-in Error', 'OK', 'Error') | Out-Null
		} finally {
			$secretBox.Clear()
			if ($secureSecret) { $secureSecret.Dispose() }
		}
	}.GetNewClosure())
	$form.ShowDialog($mainForm) | Out-Null
}.GetNewClosure())

$buttonSpecs = @(
	@{ Text = 'Convert Groups to SOA from CSV'; Column = 1; Color = [System.Drawing.Color]::FromArgb(184, 35, 45); Action = { Show-CsvGroupsForm } },
	@{ Text = 'Convert Users to SOA from CSV'; Column = 1; Color = [System.Drawing.Color]::FromArgb(184, 35, 45); Action = { Show-CsvUsersForm } },
	@{ Text = 'Convert Single User to SOA'; Column = 1; Color = [System.Drawing.Color]::FromArgb(184, 35, 45); Action = { Show-SingleObjectForm 'users' } },
	@{ Text = 'Convert Single Group to SOA'; Column = 1; Color = [System.Drawing.Color]::FromArgb(184, 35, 45); Action = { Show-SingleObjectForm 'groups' } },
	@{ Text = 'Convert All Users to SOA'; Column = 1; Color = [System.Drawing.Color]::FromArgb(184, 35, 45); Action = { Show-AllDomainForm } },
	@{ Text = 'Enabled / Disabled Users'; Column = 1; Color = [System.Drawing.Color]::FromArgb(184, 35, 45); Action = { Show-AccountStateMenu } },
	@{ Text = 'Get Users from EntraID'; Column = 2; Color = [System.Drawing.Color]::FromArgb(0, 104, 145); Action = { Show-ExportMemberUsersForm } },
	@{ Text = 'Get Groups from EntraID'; Column = 2; Color = [System.Drawing.Color]::FromArgb(0, 104, 145); Action = { Show-ExportEntraGroupsForm } },
	@{ Text = 'Validate User Permissions'; Column = 2; Color = [System.Drawing.Color]::FromArgb(75, 85, 95); Action = { Show-PermissionForm } },
	@{ Text = 'Validate Group Permissions'; Column = 2; Color = [System.Drawing.Color]::FromArgb(75, 85, 95); Action = { Show-GroupPermissionForm }; RequiresDomain = $false },
	@{ Text = 'Check SOA status for single user'; Column = 2; Color = [System.Drawing.Color]::FromArgb(126, 196, 235); TextColor = [System.Drawing.Color]::FromArgb(15, 43, 64); Action = { Show-CheckSingleUserSoaStatusForm }; RequiresDomain = $false },
	@{ Text = 'Check SOA status for single group'; Column = 2; Color = [System.Drawing.Color]::FromArgb(126, 196, 235); TextColor = [System.Drawing.Color]::FromArgb(15, 43, 64); Action = { Show-CheckSingleGroupSoaStatusForm }; RequiresDomain = $false },
	@{ Text = 'Check SOA status for Groups from CSV'; Column = 2; Color = [System.Drawing.Color]::FromArgb(126, 196, 235); TextColor = [System.Drawing.Color]::FromArgb(15, 43, 64); Action = { Show-CheckGroupsFromCsvSoaStatusForm }; RequiresDomain = $false },
	@{ Text = 'Check SOA status for Users from CSV'; Column = 2; Color = [System.Drawing.Color]::FromArgb(126, 196, 235); TextColor = [System.Drawing.Color]::FromArgb(15, 43, 64); Action = { Show-CheckUsersFromCsvSoaStatusForm }; RequiresDomain = $false }
)

$conversionHeader = [System.Windows.Forms.Label]::new()
$conversionHeader.Text = 'CONVERSION AND ACCOUNT MANAGEMENT'
$conversionHeader.Location = [System.Drawing.Point]::new(20, 145)
$conversionHeader.Size = [System.Drawing.Size]::new(350, 20)
$conversionHeader.Font = [System.Drawing.Font]::new('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$conversionHeader.ForeColor = [System.Drawing.Color]::FromArgb(75, 85, 95)
$mainForm.Controls.Add($conversionHeader)

$statusHeader = [System.Windows.Forms.Label]::new()
$statusHeader.Text = 'DIRECTORY AND STATUS CHECKS'
$statusHeader.Location = [System.Drawing.Point]::new(390, 145)
$statusHeader.Size = [System.Drawing.Size]::new(350, 20)
$statusHeader.Font = [System.Drawing.Font]::new('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$statusHeader.ForeColor = [System.Drawing.Color]::FromArgb(75, 85, 95)
$mainForm.Controls.Add($statusHeader)

$leftTop = 174
$rightTop = 174
foreach ($spec in $buttonSpecs) {
	if ($spec.Column -eq 1) {
		$buttonX = 20
		$buttonY = $leftTop
		$leftTop += 51
	} else {
		$buttonX = 390
		$buttonY = $rightTop
		$rightTop += 51
	}
	$button = Add-FormButton $mainForm $spec.Text $buttonX $buttonY 350
	$button.Height = 43
	$button.Font = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
	$button.BackColor = $spec.Color
	$button.ForeColor = if ($spec.ContainsKey('TextColor')) { $spec.TextColor } else { [System.Drawing.Color]::White }
	$button.FlatAppearance.BorderColor = $spec.Color
	$button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(
		[Math]::Min([int]$spec.Color.R + 22, 255),
		[Math]::Min([int]$spec.Color.G + 22, 255),
		[Math]::Min([int]$spec.Color.B + 22, 255)
	)
	$action = $spec.Action
	$requiresDomain = if ($spec.ContainsKey('RequiresDomain')) { [bool]$spec.RequiresDomain } else { $true }
	$button.Add_Click({
		if ($requiresDomain -and -not $appState.SelectedDomain) {
			[System.Windows.Forms.MessageBox]::Show($mainForm, 'Add and select a domain first.', 'No Domain Selected', 'OK', 'Information') | Out-Null
			return
		}
		& $action
	}.GetNewClosure())
}

$mainForm.Add_Shown({ $mainForm.Activate() })
[void]$mainForm.ShowDialog()
