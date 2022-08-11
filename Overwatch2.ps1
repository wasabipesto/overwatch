# Overwatch Script v2
# Written by Justin Dickerson
# If you are having trouble running this script, try using the following:
#   Powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Overwatch2.ps1 -help

param (
	# Get launch arguments
	[Switch]
		$Help = $false,
	[Switch]
		$Debug = $false,
	[ValidateRange(0,9999)][Int]
		$History = 1,
	[ValidateScript({Test-Path $_})][String]
		$DirectPath = "",
	[ValidateScript({Test-Path $_})][String]
		$LinksPath = "",
	[Switch]
		$LinksAscend = $false,
	[Switch]
		$IgnoreBackups = $false,
	[String]
		$EmailRecipient = "",
	[ValidateScript({Test-Path $_})][String]
		$OutputPath = "",
	[Switch]
		$OutputPretty = $true
)

if ($Help) {
	# Print the help dialog
	echo "  Overwatch Script v2"
	echo "  Written by Justin Dickerson"
	echo ""
	echo "  The purpose of this script is to recursively check all files in a path and"
	echo "  notify a user via email which files have changed in the past few days."
	echo "  Multiple target directories can be provided by creating a shortcut to each"
	echo "  and placing them all in a single folder."
	echo ""
	echo "  Usage:"
	echo "    -Help: Display this message"
	echo "    -Debug (default false): Display additional information, and output"
	echo "       certain state elements to disk"
	echo "    -History (integer 0-9999, default 1): Files modified within this number"
	echo "       of days will be displayed in the final report"
	echo "    -DirectPath (path): Scan the given path recursively for any files last"
	echo "       modified within 'History' days"
	echo "    -LinksPath (path): Check the destination of any *.lnk files in the given"
	echo "       path, then scans each recursively for any files last modified within"
	echo "       'History' days"
	echo "    -LinksAscend (default false): If LinksPath is specified, ascend from the"
	echo "       link destination by one level before scanning recursively"
	echo "    -IgnoreBackups (default false): Ignore all files with 'backup' in the path"
	echo "    -EmailRecipient (email address): If specified, send a final report to"
	echo "       the specified user through the currently-active Outlook session"
	echo "    -OutputPath (path): If specified, output the final report to the disk at"
	echo "       the specified location"
	echo "    -OutputPretty (default true): Make the report pretty with HTML, otherwise"
	echo "       the report is sent as plaintext"
	echo ""
	echo "  If you want this script to run regularly, use the following in Task Scheduler:"
	echo "    Start a program"
	echo "    Program: Powershell.exe"
	echo "    Add arguments: -NoProfile -ExecutionPolicy Bypass -File (path to this script) (arguments)"
	echo "    Start in: (leave blank)"
  exit
}

# Get the shell set up
$WScript = New-Object -ComObject WScript.Shell

# Get some basic information
$DateTime = Get-Date
$DateOnly = Get-Date -Format "MM/dd/yyyy"

echo "Overwatch starting at $DateTime"
echo ""

if ($Debug) { # Print some general information
	echo "Starting Arguments:"
	echo "Debug: $Debug"
	echo "History: $History"
	echo "DirectPath: $DirectPath"
	echo "LinksPath: $LinksPath"
	echo "LinksAscend: $LinksAscend"
	echo "IgnoreBackups: $IgnoreBackups"
	echo "EmailRecipient: $EmailRecipient"
	echo "OutputPath: $OutputPath"
	echo "OutputPretty: $OutputPretty"
	echo ""
}

if ($LinksPath) { # Extract a list of places to search from the provided links
	echo "Looking for links in $LinksPath..."
	$LinksInPath = Get-ChildItem -Path $LinksPath\*.lnk
	$LinksTarget = $LinksInPath | ForEach-Object {$WScript.CreateShortcut($_.FullName).TargetPath}
	
	if ($LinksAscend) {
		echo "Ascending to parent folder..."
		$LinksTarget = $LinksTarget | Split-Path -Parent
	}
	
	if ($Debug) {
		echo "LinksPath: $LinksPath"
		echo "LinksInPath: $LinksInPath"
		echo "LinksTarget: $LinksTarget"
	}
	
	echo "Done!"
	echo ""
}

if ($DirectPath -or $LinksPath) { # Get all ietms in the indicated paths and filter the ones we want
	if ($LinksPath) { 
		echo "Searching for files modified in the last $History days within extracted paths..." 
		$PathsToCheck = $LinksTarget
	} else { 
		echo "Searching for files modified in the last $History days within $DirectPath..."
		$PathsToCheck = $DirectPath
	}
	
	$FilesInPath = Get-ChildItem -Recurse -File -Path $PathsToCheck
	
	if ($IgnoreBackups) { 
		echo "Ignoring backups..." 
		$FilesInPath = $FilesInPath | Where {$_.FullName -notlike "*backup*"} 
	}
	
	$FilesFiltered = $FilesInPath | ? {  $_.LastWriteTime -gt (Get-Date).AddDays($History * -1) }
	
	if ($Debug) {
		echo "PathsToCheck: $PathsToCheck"
		echo "Dumping FilesInPath to log..."
		echo $FilesInPath > FilesInPath.txt     # this will take lots of disk space
		echo "Dumping FilesFiltered to log..."
		echo $FilesFiltered > FilesFiltered.txt
	}
	
	# Move objects into string for report generation
	$ReportContent = $FilesFiltered | Out-String
	
	echo "Done!"
	echo ""
}

if ($OutputPretty) { # Prettify report for email or browser
	echo "Prettifying report..."
	$ReportContent = $ReportContent -replace '\nMode\s+.*', ''								#wipe header line
	$ReportContent = $ReportContent -replace '\n\-{4}\s+.*', ''								#wipe divider line
	$ReportContent = $ReportContent -replace '(\s?) *\r\n +(\S)', '$1$2'					#fix line breaks	
	$ReportContent = $ReportContent -replace '\S{6}\s+(\d.*M)\s+\d+\s+(\S.*\S)\s+(\n)', '<tr><td>$1</td><td>$2</td></tr>$3'
																							#tabify lastmod and name	
	$ReportContent = $ReportContent -replace 'Directory: (\S.*\S)\s+(\n)', '</table><br /><a href="$1"><b>$1</b></a><table>$2'
																							#linkify directory & encase tables	
	$ReportContent = "Overwatch report started: " + $DateTime + "<br />" + "All files modified in the last " + $History + " days.<br />" + $ReportContent	
																							#add header
	$ReportContent = $ReportContent + "</table><br /><br />"								#add footer
	
	echo "Done!"
	echo ""
} 

if ($EmailRecipient) { # Connect to outlook, create email and send it
	echo "Connecting to Outlook..."
	$Outlook = New-Object -ComObject Outlook.Application
	
	echo "Getting ready to send email to $EmailRecipient..."
	$Mail = $Outlook.CreateItem(0)
	$Mail.To = $EmailRecipient
	$Mail.Subject = "Overwatch Report " + $DateOnly 
	
	if ($OutputPretty) { 
		$Mail.BodyFormat = 2
		$Mail.HTMLBody = $ReportContent 
	} else {
		$Mail.Body = $ReportContent 
	}
	
	echo "Sending email..."
	$Mail.Send()
	
	echo "Done!"
	echo ""
}

if ($OutputPath) { # Dump report to disk
	echo "Saving report to $OutputPath..."
	if ($OutputPretty) { 
		echo $ReportContent > OverwatchReport.html
	} else {
		echo $ReportContent > OverwatchReport.txt
	}
	
	echo "Done!"
	echo ""
}