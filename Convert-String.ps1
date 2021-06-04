<#
.SYNOPSIS
all kinds of string and/or data manipulation functions
.DESCRIPTION
created over a period of time when dealing with mundane task of cleansing data input
.NOTES
FileName:	Convert-Strings.ps1
CoAuthor:	Chris Kenis
Contact: 	@KICTS
Created: 	04-06-2020
#>
function Convert-WideSpaceToColumnWidth {
	param (
		[string]$InputValue,
		[byte]$NumberOfSpaces = 7
	)
	#find number of consecutive spaces between word boundaries = not at start or end of inputvalue
	$indices = @([regex]::Matches($InputValue, "(?<=\b|\B)[\s|\t]{$($NumberOfSpaces),}(?=\b|\B)").Index | Sort-Object)
	if (-not $indices) { $indices = $InputValue.Length } #short string parsed, probably singular value
	return $indices
}
function ConvertTo-KeyValue {
	[OutputType([System.Collections.ArrayList])]
	param(
		[string[]]$FieldValues,
		[string]$Separator = ':',
		[int]$StartFromIndex
	)
	[System.Collections.ArrayList]$KeyValues = @()
	for ([int]$i = 0; $i -lt $FieldValues.Count; $i++) {
		[string]$FieldString = ""
		if ($i -lt $StartFromIndex) { $FieldString = $FieldValues[$i] }
		else {
			do { $FieldString = $FieldString.Trim(), $FieldValues[$i].Trim() -join ' '; $i++ }
			until (($FieldValues[$i] -match $Separator) -or ($i -ge $FieldValues.Count))
			$i -= 1 #go back 1 column for next iteration
		}
		$KeyValues.Add($FieldString.Trim()) | Out-Null
	}
	return $KeyValues
}
function Split-KeyValue ([string]$InputValue) {
	$return = ($InputValue -split ":", 2)[-1]
	if (-not [string]::IsNullOrWhiteSpace($return)) { return $return.Trim() }
}
function Split-FixedWidth {
	param(
		[Parameter(Mandatory, ValueFromPipeline)][AllowEmptyString()][string]$InputValue,
		[Parameter(Mandatory)][int[]]$Columns
	)
	process {
		[System.Collections.ArrayList]$return = @()
		$Columns = @($Columns | Where-Object { $_ -le $InputValue.TrimEnd().Length } | Sort-Object ) #prevent upperbound violation
		if (-not $Columns) { [void]$return.Add($InputValue.Trim()) } #first field could be shorter than minimum column width
		else {
			[void]$return.Add($InputValue.SubString(0, $Columns[0]).Trim()) #add first field from start of inputvalue
			for ($i = 1; $i -lt $Columns.Count; $i++) {
				[int]$Pos = $Columns[($i - 1)] #starting position of substring
				[int]$Len = $Columns[$i] - $Columns[($i - 1)] #number of characters to return
				[void]$return.Add($InputValue.SubString($Pos, $Len).Trim())
			}
			[void]$return.Add($InputValue.SubString($Columns[-1]).Trim()) #add last field until end of inputvalue
		}
		return $return
	}
}
function Split-WideSpaced {
	param(
		[Parameter(Mandatory, ValueFromPipeline)][AllowEmptyString()][string]$FieldValue,
		[byte]$SpaceWidth = 3,
		[ValidateSet(';', ',', '|', ':', '-')][string]$Delimiter,
		[switch]$KeyValueCorrection
	)
	#use of variable widespace and not \s+ is a tradeoff for predictability
	process {
		[System.Collections.ArrayList]$return = @()
		if (-not [string]::IsNullOrWhiteSpace($FieldValue)) {
			#sometimes the normal delimiter is missing from the input and is replaced by a space resulting in a widespace between words
			if ($PSBoundParameters.ContainsKey('Delimiter')) { [string[]]$return = @(($FieldValue -replace $WideSpace, $Delimiter) -split $Delimiter) }
			else { [string[]]$return = @($FieldValue -split $WideSpace) }
			$return = $return.Where( { -not [string]::IsNullOrWhiteSpace($_) }).Trim()
			if ($KeyValueCorrection.IsPresent) { $return = ConvertTo-KeyValue -FieldValues $return -StartFromIndex 1 }
		}
		return $return
	}
	begin {
		$WideSpace = "(?<=\b|\B)[\s|\t]{$($SpaceWidth),}(?=\b|\B)"
	}
}
function Format-CardNumber ([string]$InputValue) {
	if (-not [string]::IsNullOrWhiteSpace($InputValue)) {
		#strip prefix and all non digit or letter characters
		$CardNumber = ($InputValue -replace "Kaart(nummer)?:?\s+" -replace '\W')
		switch ($CardNumber.Length) {
			22 { $Columns = @(4, 8, 12, 16, 20); break }
			18 { $Columns = @(4, 8, 12, 16); break }
			17 { $Columns = @(4, 8, 12, 16); break }
			16 { $Columns = @(4, 8, 12); break }
			12 { $Columns = @(3, 7, 10); break }
			10 { $Columns = @(8, 10); break }
		}
		$Return = (Split-FixedWidth -InputValue $CardNumber -Columns $Columns) -join '-'
	}
	return $Return
}
function ConvertTo-Time ([string]$InputValue) {
	[cultureinfo]$script:MyCulture = Get-Culture
	if (-not [string]::IsNullOrWhiteSpace($InputValue)) {
		$Separators = '\.|-|u'
		try {
			$TimeVal = $InputValue.Trim() -replace $Separators, ':'
			switch ($TimeVal.Length ) {
				( { $_ -gt 6 }) { $TimeFormat = 'H:mm:ss'; break }
				( { $_ -le 5 }) { $TimeFormat = 'H:mm'; break }
			}
			$Time = [datetime]::ParseExact($TimeVal, $TimeFormat, $script:MyCulture)
		}
		catch { Write-LogEntry -Message "Could not convert inputvalue: '$($InputValue)' to Time" -Severity 2 }
	}
	return $Time
}
function ConvertTo-Date {
	param (
		[string]$InputValue,
		[DateTime]$ReferenceDate,
		[switch]$CorrectYearSequence
	)
	[cultureinfo]$script:MyCulture = Get-Culture
	if (-not [string]::IsNullOrWhiteSpace($InputValue)) {
		$Separators = ':|\.|-'
		try {
			$DateVal = $InputValue.Trim() -replace $Separators, '/'
			$DateVal = ($DateVal -split '/').ForEach( { $_.Padleft(2, '0') }) -join ('/')
			$DateFormat = 'dd/MM/yyyy'
			#add Year part from ReferenceDate if missing from inpuntvalue
			switch ($DateVal.Length ) {
				( { $_ -lt 6 }) {
					$DateVal = $DateVal, $ReferenceDate.Year -join '/'
					$Date = [datetime]::ParseExact($DateVal, $DateFormat, $script:MyCulture)
					if ($CorrectYearSequence.IsPresent) {
						#account for year end transition when more than 11 (absolute) months difference
						#comes in handy when processing sequential dates
						if (($Date.Month - $ReferenceDate.Month) -ge 11) { $Date = $Date.AddYears(-1) }
						if (($Date.Month - $ReferenceDate.Month) -le -11) { $Date = $Date.AddYears(1) }
					}
					break 
				}
				( { $_ -le 8 }) {
					$DateFormat = 'dd/MM/yy'
					$Date = [datetime]::ParseExact($DateVal, $DateFormat, $script:MyCulture)
					break
				}
				( { $_ -ge 8 }) { $Date = [datetime]::ParseExact($DateVal, $DateFormat, $script:MyCulture) }
			}
		}
		catch { Write-Warning -Message "Could not convert '$($InputValue)' to Date" -Severity 2 }
	}
	return $Date
}
function Convert-Amount {
	param (
		[string]$InputValue,
		$ThousandSeparator = '.'
	)
	[cultureinfo]$script:MyCulture = Get-Culture
	if (-not [string]::IsNullOrWhiteSpace($InputValue)) {
		try {
			#strip all unnecessary characters from numeric value and convert to decimal for further calculation(s)
			$Amount = ($InputValue -replace '\s+' -replace [Regex]::Escape($ThousandSeparator))
			$Amount = "{0:f2}" -f ([System.Convert]::ToDecimal($Amount, $script:MyCulture))
		}
		catch { Write-LogEntry -Message "Could not convert inputvalue: '$($InputValue)' to Decimal" -Severity 2 ; $Amount = $null }
	}
	return $Amount
}
function Select-FieldValue {
	param (
		[string]$Value1,
		[string]$Value2,
		[ValidateSet('First', 'Last', 'Left', 'Right', 'Both', 'Prompt')]$Preference
	)
	try {
		$Returns = @(($Value1, $Value2).where( { -not [string]::IsNullOrWhiteSpace($_) }))
		if ($Returns) {
			switch ($Preference) {
				{ @('First', 'Left') -contains $_ } { $Return = $Returns[0] }
				{ @('Last', 'Right') -contains $_ } { $Return = $Returns[-1] }
				'Prompt' { $Return = Out-GridView -InputObject $Returns -Title "Select value to be returned:" -OutputMode Single }
				'Both' { $Return = $Returns -join '|' }
				default { $Return = $Returns[0] }
			}
		}
	}
	catch { Write-LogEntry -Message "Error comparing/merging values" }
	return $Return
}
