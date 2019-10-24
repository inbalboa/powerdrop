<#
.SYNOPSIS
    Dropbox upload tool.
.DESCRIPTION
    Uploads spesific directory to Dropbox.
.PARAMETER SourceDirPath
    Local directory for upload.
.PARAMETER TargetDirPath
    Remote Dropbox directory. May start from /.
.PARAMETER Token
    Remote Dropbox token.
.PARAMETER Log
    Path to log file.
.EXAMPLE
    C:\PS> .\powerdrop.ps1 -SourceDirPath localedir -TargetDirPath /remotedir -Token "00000-00000_00000-0" -Log dbox.log
.NOTES
    Author: Sergey Shlyapugin
    Date: Oct 24, 2019
    Version: 0.0.1
#>

Param(
	[CmdletBinding()]
    [Parameter(Mandatory=$true)]
    [string]$SourceDirPath,
    [Parameter(Mandatory=$true)]
    [string]$TargetDirPath,
    [Parameter(Mandatory=$true)]
    [string]$Token,
    [Parameter(Mandatory=$false)]
    [string]$Log
)

Function GetFiles {
	Param(
		[string]$SourceDirPath
	)
	
	Get-ChildItem -Path $SourceDirPath | Where-Object {!$_.PSisContainer}
}

Function UploadFileSess {
	Param(
		[string]$SourceFilePath,
		[string]$TargetFilePath
	)
	
	$file = Get-Item $SourceFilePath
	$flen = $file.Length
	$fname = $file.Name
	$uri = "https://content.dropboxapi.com/2/files/upload_session/"
	
	$a = $false
	Try {
		$argStr = @{"close" = $false} | ConvertTo-JSON
		$webRequest = [System.Net.WebRequest]::Create($uri + "start")
		$webRequest.Headers.Add("Authorization", "Bearer " + $Script:Token)
		$webRequest.Headers.Add("Dropbox-API-Arg", $argStr)
		$webRequest.Method = "POST"
		$WebRequest.ContentType = "application/octet-stream"
		[System.Net.WebResponse] $resp = $webRequest.GetResponse()
		$rs = $resp.GetResponseStream()
		[System.IO.StreamReader] $sr = New-Object System.IO.StreamReader -argumentList $rs
		[string] $results = $sr.ReadToEnd()
		$session_id = (ConvertFrom-JSON $results).session_id
	
		$bufSize = 100MB
		$bytesWritten = 0

		$fileStream = [System.IO.File]::Open($file, "Open", "Read", "None")
		$chunk = New-Object byte[] $bufSize
		while ($bytesRead = $fileStream.Read($chunk, 0, $bufsize)) {
			for ($i = 0; $i -lt 10; $i++) {
				try {
					if (($bytesWritten + $bytesRead) -ge $flen) {
						$webRequest = [System.Net.WebRequest]::Create($uri + "finish")
						$argStr = @{"cursor" = @{"session_id" = $session_id; "offset" = $bytesWritten}; "commit" = @{ "path" = $TargetFilePath; "mode" = "add"; "autorename" = $true; "mute" = $false; strict_conflict = $false}} | ConvertTo-JSON
					}
					else {
						$webRequest = [System.Net.WebRequest]::Create($uri + "append_v2")
						$argStr = @{"cursor" = @{"session_id" = $session_id; "offset" = $bytesWritten;}; "close" = $false} | ConvertTo-JSON
					}
					break
				}
				catch {
					Log("Error uploading chunk for file {0}: {1}" -f $file, $_.Exception.Message)
					Log("Sleeping for 1 min...")
					start-sleep -s 60
				}
			}
			$webRequest.Headers.Add("Dropbox-API-Arg", $argStr)
			$webRequest.Headers.Add("Authorization", "Bearer " + $Script:Token)
			$webRequest.Method = "POST"
			$webRequest.ContentType = "application/octet-stream"
			$webRequest.ProtocolVersion = [System.Net.HttpVersion]::Version11
			$webRequest.ContentLength = $bytesRead

			$requestStream = $webRequest.GetRequestStream()
			$progressActivityMessage = ("Sending file... {0} - {1} bytes" -f $file.Name, $flen)
			
			$requestStream.write($chunk, 0, $bytesRead)
			$requestStream.Flush()
			$bytesWritten += $bytesRead
			$progressStatusMessage = ("Sent {0} bytes - {1:N0} MB" -f $bytesWritten, ($bytesWritten / 1MB))
			Write-Progress -Activity $progressActivityMessage -Status $progressStatusMessage -PercentComplete ($bytesWritten/$flen*100)
			
			if ($requestStream) { $requestStream.Close() }
			$response = $webRequest.GetResponse()
			$response.Close();
		}
		
		Log("{0}: uploaded successfully" -f $fname)
		$a = $true
	}
	Catch {
		Log("{0}: error uploading {1}" -f $fname, $_.Exception.Message)
	}
	Finally {
		if ($fileStream) { $FileStream.Close() }
	}
	
	return $a
}

Function UploadFiles {
	Param(
		[array]$Files,
		[string]$TargetDirPath
	)	
	
	New-Variable -Name "Txt" -Value "" -Scope Script
	$Files | Foreach-Object {
		$Script:Txt = "<b>{0}</b>`n`n" -f $env:computername
		if (UploadFileSess $_.FullName ($TargetDirPath + "/" + $_.Name)) { DelFile $_.FullName }
	}
}

Function DelFile {
	Param(
		[string]$FilePath
	)

	$fname = (Get-Item $FilePath).Name
	try {
		Remove-Item $FilePath
		Log("{0}: removed successfully" -f $fname)
	}
	catch {
		Log("{0}: error removing {1}" -f $fname, $_.Exception.Message)
	}
}

Function Escape-JSONString($str){
	if ($str -eq $null) { return "" }
	$str = $str.ToString().Replace('"','\"').Replace('\','\\').Replace("`n",'\n').Replace("`r",'\r').Replace("`t",'\t')
	$str;
}

Function ConvertTo-JSON($maxDepth = 4, $forceArray = $false) {
	begin {
		$data = @()
	}
	process{
		$data += $_
	}
	
	end{	
		if ($data.length -eq 1 -and $forceArray -eq $false) { $value = $data[0] }
		else {	$value = $data }

		if ($value -eq $null) { return "null" }

	$dataType = $value.GetType().Name
	switch -regex ($dataType) {
	            'String'  {
					return  "`"{0}`"" -f (Escape-JSONString $value )
				}
	            '(System\.)?DateTime'  {return  "`"{0:yyyy-MM-dd}T{0:HH:mm:ss}`"" -f $value}
	            'Int32|Double' {return  "$value"}
				'Boolean' {return  "$value".ToLower()}
	            '(System\.)?Object\[\]' { # array
					
					if ($maxDepth -le 0){return "`"$value`""}
					
					$jsonResult = ''
					foreach($elem in $value){
						if ($jsonResult.Length -gt 0) {$jsonResult +=', '}				
						$jsonResult += ($elem | ConvertTo-JSON -maxDepth ($maxDepth -1))
					}
					return "[" + $jsonResult + "]"
	            }
				'(System\.)?Hashtable' { # hashtable
					$jsonResult = ''
					foreach($key in $value.Keys){
						if ($jsonResult.Length -gt 0) {$jsonResult +=', '}
						$jsonResult += 
@"
	"{0}": {1}
"@ -f $key , ($value[$key] | ConvertTo-JSON -maxDepth ($maxDepth -1) )
					}
					return "{" + $jsonResult + "}"
				}
	            default { #object
					if ($maxDepth -le 0){return  "`"{0}`"" -f (Escape-JSONString $value)}
					
					return "{" +
						(($value | Get-Member -MemberType *property | % { 
@"
	"{0}": {1}
"@ -f $_.Name , ($value.($_.Name) | ConvertTo-JSON -maxDepth ($maxDepth -1) )			
					
					}) -join ', ') + "}"
	    		}
		}
	}
}

Function ConvertFrom-JSON([object] $item) { 
    add-type -assembly system.web.extensions
    $ps_js=new-object system.web.script.serialization.javascriptSerializer

    #The comma operator is the array construction operator in PowerShell
    return $ps_js.DeserializeObject($item)
}

Function Log {
	Param([string]$Txt)
	
	Write-Host $Txt
	
	if ($Script:Log) { Out-File -FilePath $Script:Log -InputObject ("[" + (Get-Date).ToString() + "] " + $Txt) -Append -encoding unicode }		
}

$Files = GetFiles $SourceDirPath
if ($Files) { UploadFiles $Files $TargetDirPath }
