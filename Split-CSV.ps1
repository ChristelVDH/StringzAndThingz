[cmdletbinding()]

param (
    [System.IO.FileInfo]$InFile,
    [System.IO.FileInfo]$OutFile,
    [ValidateNotNullOrEmpty][string]$SearchString,
    [ValidateSet('First', 'Last')]$SplitAtOccurrence = 'First',
    [switch]$Append
)
switch ($SplitAtOccurrence) {
    'First' { $SplitIndex = 0 }
    'Last' { $SplitIndex = -1 }
}
$text = Get-Content -Path $InFile.FullName
#split lines at <First> or <Last> occurrence and subtract -1 from linenumber as index starts at 0
$MatchingLineNumber = ($text | Select-String -Pattern $SearchString)[$SplitIndex].LineNumber - 1
if ($MatchingLineNumber) {
    if ($Append.IsPresent){ Add-Content -Value $text[1..$MatchingLineNumber] -Path $OutFile.FullName }
    else { Set-Content -Value $text[0..$MatchingLineNumber] -Path $OutFile.FullName }
    Write-Verbose -Message "Entries until $($SplitAtOccurrence) match pattern <$($SearchString)> have been moved to: $($OutFile.FullName)"
    try {
        Remove-Item -Path $InFile.Directory -Include "remove*.prev" -Force
        Rename-Item -Path $InFile -NewName "remove_$($InFile.BaseName).prev"
        Write-Verbose -Message "Original file has been renamed to: remove_$($InFile.BaseName).prev"
        Set-Content -Path $InFile.FullName -Value $text[0] #create new logfile and write header row
        Add-Content -Path $InFile.FullName -Value $text[($MatchingLineNumber+1)..-1] #copy lines after oldest entry until end of text
    }
    catch [System.IO.FileNotFoundException] { Write-Warning -Message "Failed to clean up existing file: $($InFile.FullName)" }
    catch { Write-Error -Message "Unknown error when trying to (re)create new file!!!" }
}
else { Write-Warning -Message "SearchString <$($SearchString)> did not occur in file $($InFile.Name)" }
