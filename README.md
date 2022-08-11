# Overwatch
The purpose of this script is to recursively check all files in a path and notify a user via email which files have changed in the past few days. Multiple target directories can be provided by creating a shortcut to each and placing them all in a single folder.

Usage:

- -Help: Display this message
- -Debug (default false): Display additional information, and output certain state elements to disk
- -History (integer 0-9999, default 1): Files modified within this number of days will be displayed in the final report
- -DirectPath (path): Scan the given path recursively for any files last modified within 'History' days
- -LinksPath (path): Check the destination of any *.lnk files in the given path, then scans each recursively for any files last modified within 'History' days
- -LinksAscend (default false): If LinksPath is specified, ascend from the link destination by one level before scanning recursively
- -IgnoreBackups (default false): Ignore all files with 'backup' in the path
- -EmailRecipient (email address): If specified, send a final report to the specified user through the currently-active Outlook session
- -OutputPath (path): If specified, output the final report to the disk at the specified location
- -OutputPretty (default true): Make the report pretty with HTML, otherwise the report is sent as plaintext

If you want this script to run regularly, use the following in Task Scheduler:
- Start a program
- Program: Powershell.exe
- Add arguments: -NoProfile -ExecutionPolicy Bypass -File (path to this script) (arguments)
- Start in: (leave blank)
