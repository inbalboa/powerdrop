# Powerdrop

Dropbox dir uploader on pure Powershell.
Even Powershell 1.0 is upported! 

You need a Dropbox access token for this script.
Go to https://www.dropbox.com/developers/apps/, create your own app and generate the access token.

Example usage:
`C:\PS> .\powerdrop.ps1 -SourceDirPath localedir -TargetDirPath /remotedir -Token "00000-00000_00000-0" -Log dbox.log`

Be careful! Files will be removed after successful upload to Dropbox.