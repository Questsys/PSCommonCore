<#	
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2019 v5.6.166
	 Created on:   	8/6/2019 11:59 AM
	 Created by:   	Gary Cook
	 Organization: 	Quest
	 Filename:     	PSCommonCore.psm1
	-------------------------------------------------------------------------
	 Module Name: PSCommonCore
	===========================================================================
	#requires the PShellLogging module found on the Powershell Gallery
#>




Function Test-Module
{
	param
	(
		[parameter(Mandatory = $true)]
		[string]$Module,
		[parameter(Mandatory = $true)]
		[bool]$Load = $false
		
	)
	[CmdletBinding]
	$Return = New-Object System.Management.Automation.PSObject
	$Return | Add-Member -MemberType NoteProperty -Name Status -Value $null
	$Return | Add-Member -MemberType NoteProperty -Name Message -Value $null
	
	if (!(get-module $Module))
	{
		If (!(Get-Module -ListAvailable $Module))
		{
			$Return.Status = "Not Loaded"
			$Return.Message = "The Module $($Module) was not available to load"
			
		}
		else
		{
			if ($Load -eq $true)
			{
				Import-Module ActiveDirectory -ErrorAction SilentlyContinue
				if (!(Get-Module activedirectory))
				{
					$Return.Status = "Not Loaded"
					$Return.Message = "The Module $($Module) was available but failed to load"
				}
				else
				{
					$Return.Status = "Loaded"
					$Return.Message = "Module $($Module) was loaded successfully"
				}
			}
			else
			{
				$Return.Status = "Not Loaded"
				$Return.Message = "The Module $($Module) was available but not loaded due to switch"
			}
			
		}
		
	}
	else
	{
		$Return.Status = "Loaded"
		$Return.Message = "Module $($Module) was already loaded"
	}
	return $Return
}

function Test-PSVersion 
{
	param
	(
		[parameter(Mandatory = $true)]
		[string]$Min,
		[parameter(Mandatory = $true)]
		[ValidateSet("Desktop","Core")]
		[string]$Edition = "Desktop"
	)
	
	
	if ($Min.contains(".") -eq $true)
	{
		#Write-Host "Checking Powershell against $($Min)"
		$Major = ($Min -split "\.")[0]
		#Write-Host "Major Version: $($Major)"
		$Minor = ($Min -split "\.")[1]
		#Write-Host "Minor Version: $($Minor)"
	}
	else
	{
		#Write-Host "Checking Powershell against $($Min)"
		$Major = $Min
		#Write-Host "Major Version: $($Major)"
		$Minor = "0"
		#Write-Host "Minor Version: $($Minor)"
	}
	$Version = $PSVersionTable.pscompatibleversions
	$Return = $false
	#Write-Host "Processing PowerShell Compatable Versions"
	foreach ($V in $Version)
	{
		#Write-Host "Checking Version Major:$($V.major) Minor:$($V.minor)"
		if ($V.major -eq $Major -and $V.minor -eq $Minor)
		{
			#Write-Host "Matches Minimum Version"
			$Return = $true
		}
	}
	if ($Edition -eq $PSVersionTable.PSEdition)
	{
		return $Return
	}
	else
	{
		return $false
	}
	
	
}

function Write-Color([String[]]$Text, [ConsoleColor[]]$Color = (get-host).ui.rawui.ForegroundColor, [ConsoleColor[]]$BackColor = (get-host).ui.rawui.BackgroundColor, [int]$StartTab = 0, [int]$LinesBefore = 0, [int]$LinesAfter = 0)
{
	$DefaultColor = $Color[0]
	$DefaultBackColor = $BackColor[0]
	if ($LinesBefore -ne 0) { for ($i = 0; $i -lt $LinesBefore; $i++) { Write-Host "`n" -NoNewline } } # Add empty line before
	if ($StartTab -ne 0) { for ($i = 0; $i -lt $StartTab; $i++) { Write-Host "`t" -NoNewLine } } # Add TABS before text
	if ($Color.Count -ge $Text.Count -and $BackColor.count -ge $Text.count)
	{
		for ($i = 0; $i -lt $Text.Length; $i++) { Write-Host $Text[$i] -ForegroundColor $Color[$i] -backgroundcolor $BackColor[$i] -NoNewLine }
	}
	else
	{
		if ($Color.Count -ge $Text.Count -and $BackColor.Count -lt $Text.Count)
		{
			for ($i = 0; $i -lt $BackColor.Length; $i++) { Write-Host $Text[$i] -ForegroundColor $Color[$i] -BackgroundColor $BackColor[$i] -NoNewLine }
			for ($i = $BackColor.Length; $i -lt $Text.Length; $i++) { Write-Host $Text[$i] -ForegroundColor $Color[$i] -BackgroundColor $DefaultBackColor -NoNewLine }
		}
		if ($Color.Count -lt $Text.Count -and $BackColor.Count -ge $Text.Count)
		{
			for ($i = 0; $i -lt $Color.Length; $i++) { Write-Host $Text[$i] -ForegroundColor $Color[$i] -BackgroundColor $BackColor[$i] -NoNewLine }
			for ($i = $BackColor.Length; $i -lt $Text.Length; $i++) { Write-Host $Text[$i] -ForegroundColor $DefaultColor -BackgroundColor $BackColor[$i] -NoNewLine }
		}
		if ($Color.Count -lt $Text.Count -and $BackColor.Count -lt $Text.Count)
		{
			if ($Color.Count -lt $BackColor.count)
			{
				for ($i = 0; $i -lt $Color.Length; $i++) { Write-Host $Text[$i] -ForegroundColor $Color[$i] -BackgroundColor $BackColor[$i] -NoNewLine }
				for ($i = $Color.Length; $i -lt $BackColor.length; $i++) { Write-Host $Text[$i] -ForegroundColor $DefaultColor -BackgroundColor $BackColor[$i] -NoNewLine }
				for ($i = $BackColor.Length; $i -lt $Text.length; $i++) { Write-Host $Text[$i] -ForegroundColor $DefaultColor -BackgroundColor $DefaultBackColor -NoNewLine }
			}
			if ($Color.Count -gt $BackColor.count)
			{
				for ($i = 0; $i -lt $BackColor.Length; $i++) { Write-Host $Text[$i] -ForegroundColor $Color[$i] -BackgroundColor $BackColor[$i] -NoNewLine }
				for ($i = $BackColor.Length; $i -lt $Color.length; $i++) { Write-Host $Text[$i] -ForegroundColor $Color[$i] -BackgroundColor $DefaultBackColor -NoNewLine }
				for ($i = $Color.Length; $i -lt $Text.length; $i++) { Write-Host $Text[$i] -ForegroundColor $DefaultColor -BackgroundColor $DefaultBackColor -NoNewLine }
			}
			if ($Color.Count -eq $BackColor.count)
			{
				for ($i = 0; $i -lt $BackColor.Length; $i++) { Write-Host $Text[$i] -ForegroundColor $Color[$i] -BackgroundColor $BackColor[$i] -NoNewLine }
				for ($i = $BackColor.Length; $i -lt $text.length; $i++) { Write-Host $Text[$i] -ForegroundColor $DefaultColor -BackgroundColor $DefaultBackColor -NoNewLine }
			}
		}
		
		
	}
	Write-Host
	if ($LinesAfter -ne 0) { for ($i = 0; $i -lt $LinesAfter; $i++) { Write-Host "`n" } } # Add empty line after
}

function Connect-EOL
{
	param
	(
		[parameter(Mandatory = $false)]
		[System.Management.Automation.PSCredential]
		$Credential
	)
	if ($Credential -eq $null)
	{
		$Credential = Get-Credential -Message "Please Enter your Exchange Online Admin Credential."
	}
	$Session = New-PSSession -ConfigurationName "Microsoft.Exchange" -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $Credential -Authentication Basic -AllowRedirection
	Import-PSSession $Session -DisableNameChecking -AllowClobber
	return $session
}

function Disconnect-EOL
{
	param
	(
		[parameter(Mandatory = $true,ValueFromPipeline = $true)]
		$Session
	)
	Remove-PSSession -Session $Session
}

function Start-Log
{
	<#
        .SYNOPSIS
			Creates the supplied log file $Log.
        .DESCRIPTION
        .PARAMETER
			$Log
				the complete path to the log to write to. required.
			$type
				the type of log file to generate TXT is assumed.  Possible Values are TXT, CSV, JSON.
        .EXAMPLE
			creates a log at location and returns object representing the log and type
			Start-Log -Log "C:\applog.txt" -Type CSV
        .NOTES
            FunctionName :	Write-Log
            Created by   :	Gary Cook
            Date Coded   : 	07/26/2019
		.OUTPUTS
			Returns and object containing the path to the log and the type of the log.
    #>
	[CmdletBinding()]
	Param
	(
		[parameter (position = 0, Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
		[string]$Log,
		[Parameter (Position = 1, Mandatory = $false, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
		[ValidateSet("TXT", "CSV", "JSON")]
		[string]$Type = "TXT"
	)
	Begin
	{
	}
	Process
	{
		
		# Format Date for our Log File 
		$FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
		#if log does not exists create log with application runtime banner
		if (!(Test-Path $Log -PathType Leaf))
		{
			if (!(Test-Path $Log))
			{
				#create file including path if the path does not exist
				$NewLogFile = New-Item $Log -Force -ItemType File
			}
			if ($Type -eq "TXT")
			{
				#create file with banner
				$Banner = "*************************************************"
				$Banner | Out-File -FilePath $Log -Append -force
				$Banner = "Application log created $($FormattedDate) on computer $($env:COMPUTERNAME)"
				$Banner | Out-File -FilePath $Log -Append
				$Banner = "*************************************************"
				$Banner | Out-File -FilePath $Log -Append
			}
			if ($Type -eq "CSV")
			{
				#open out file with headder
				$Banner = "Date,Level,Message"
				$Banner | Out-File -FilePath $Log -Append -force
				$Banner = "$($FormattedDate),INFO:,Application Log file Created for computer $($env:COMPUTERNAME)"
				$Banner | Out-File -FilePath $Log -Append
			}
			if ($Type -eq "JSON")
			{
				$Banner = "{`"DATE`": `"$($FormattedDate)`",`"LEVEL`": `"INFO:`",`"MESSAGE`": `"Application Log file Created for computer $($env:COMPUTERNAME)`"}"
				$Banner | Out-File -FilePath $Log -Append -force
			}
			
		}
		$obj = new-object System.Management.Automation.PSObject
		$obj | Add-Member -MemberType NoteProperty -Name Log -Value (get-item $log).VersionInfo.filename
		$obj | Add-Member -MemberType NoteProperty -Name Type -Value $Type
		
		return $obj
		
	}
	end
	{
		
	}
	
	
}



Function Write-Log
{
    <#
        .SYNOPSIS
			Writes the Entry in $Line to the supplied log file $Log.  Built to take pipeline input from object returned from start-log.
        .DESCRIPTION
        .PARAMETER
			$Line
				String of data to write to the log file. required.
			$Log
				the complete path to the log to write to. required.
			$Level
				The type of line to write to the log.  Valid vales are Error,Warn,Info. Default is Info.
			$Type
				the type of log file to generate TXT is assumed.  Possible Values are TXT, CSV, JSON.
        .EXAMPLE
			$mylog | Write-Log -Line "This is an entry for the log" -level Info
        .NOTES
            FunctionName :	Write-Log
            Created by   :	Gary Cook
            Date Coded   : 	07/26/2019
		.OUTPUTS
			Returns 0 if log exists or -1 if the log file does not exist
    #>
	[CmdletBinding()]
	Param
	(
		[parameter (position = 0, Mandatory = $true)]
		[string]$Line,
		[parameter (position = 1, Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
		[string]$Log,
		[Parameter (position = 2, Mandatory = $false)]
		[ValidateSet("Error", "Warn", "Info")]
		[string]$Level = "Info",
		[Parameter (Position = 3, Mandatory = $false, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
		[ValidateSet("TXT", "CSV", "JSON")]
		[string]$Type = "TXT"
	)
	Begin
	{
	}
	Process
	{
		
		# Format Date for our Log File 
		$FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
		# Write message to error, warning, or verbose pipeline and specify $LevelText 
		switch ($Level)
		{
			'Error' {
				
				$LevelText = 'ERROR:'
			}
			'Warn' {
				
				$LevelText = 'WARNING:'
			}
			'Info' {
				
				$LevelText = 'INFO:'
			}
		}
		#if log does not exists reutrn -1 else return 0
		if (!(Test-Path $Log -PathType Leaf))
		{
			if (!(Test-Path $Log))
			{
				return -1
				break
				
			}
			
			
		}
		# Write message to proper log type 
		switch ($Type)
		{
			'TXT' {
				"$($FormattedDate) $($LevelText) $($Line)" | Out-File -FilePath $Log -Append
			}
			'CSV' {
				"$($FormattedDate),$($LevelText),$($Line)" | Out-File -FilePath $Log -Append
			}
			'JSON' {
				"{`"DATE`": `"$($FormattedDate)`",`"LEVEL`": `"$($LevelText)`",`"MESSAGE`": `"$($Line)`"}" | Out-File -FilePath $Log -Append
			}
		}
		
		
		return 0
	}
	End
	{
	}
}


Function Close-Log
{
	<#
        .SYNOPSIS
			Closes the supplied log file $Log.  Built to take pipeline input from object returned from start-log.
        .DESCRIPTION
        .PARAMETER
			$Log
				the complete path to the log to write to. required.
			$Type
				the type of log file to generate TXT is assumed.  Possible Values are TXT, CSV, JSON.
        .EXAMPLE
			$mylog | Close-Log 
        .NOTES
            FunctionName :	Write-Log
            Created by   :	Gary Cook
            Date Coded   : 	07/26/2019
    #>
	[CmdletBinding()]
	Param
	(
		[parameter (position = 0, Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
		[string]$Log,
		[Parameter (Position = 1, Mandatory = $false, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
		[ValidateSet("TXT", "CSV", "JSON")]
		[string]$Type = "TXT"
	)
	Begin
	{
	}
	Process
	{
		
		# Format Date for our Log File 
		$FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
		if ($Type -eq "TXT")
		{
			#close out file with footer
			$Footer = "*************************************************"
			$Footer | Out-File -FilePath $Log -Append
			$Footer = "Application log end $($FormattedDate) on computer $($env:COMPUTERNAME)"
			$Footer | Out-File -FilePath $Log -Append
			$Footer = "*************************************************"
			$Footer | Out-File -FilePath $Log -Append
		}
		if ($Type -eq "CSV")
		{
			#close out file with footer
			$Footer = "$($FormattedDate),INFO:,Application Log file end for computer $($env:COMPUTERNAME)"
			$Footer | Out-File -FilePath $Log -Append
		}
		if ($Type -eq "JSON")
		{
			$Footer = "{`"DATE`": `"$($FormattedDate)`",`"LEVEL`": `"INFO:`",`"MESSAGE`": `"Application Log file end for computer $($env:COMPUTERNAME)`"}"
			$Footer | Out-File -FilePath $Log -Append
		}
		
		
		
	}
	End
	{
	}
	
	
}

function Scramble-String([string]$inputString)
{
	$characterArray = $inputString.ToCharArray()
	$scrambledStringArray = $characterArray | Get-Random -Count $characterArray.Length
	$outputString = -join $scrambledStringArray
	return $outputString
}

function Get-RandomCharacters($length, $characters)
{
	$random = 1 .. $length | ForEach-Object { Get-Random -Maximum $characters.length }
	$private:ofs = ""
	return [String]$characters[$random]
}

function create-password
{
	param
	(
		[parameter(Mandatory = $true)]
		[int]$Length,
		[parameter(Mandatory = $false)]
		[switch]$Upper = $false,
		[parameter(Mandatory = $false)]
		[switch]$Lower = $false,
		[parameter(Mandatory = $false)]
		[switch]$Number = $false,
		[parameter(Mandatory = $false)]
		[switch]$Symbol = $false,
		[parameter(Mandatory = $false)]
		[int]$MinLower = 1,
		[parameter(Mandatory = $false)]
		[int]$MinUpper = 1,
		[parameter(Mandatory = $false)]
		[int]$MinNumber = 1,
		[parameter(Mandatory = $false)]
		[int]$MinSymbol = 1,
		[parameter(Mandatory = $false)]
		[switch]$NonAmbiguous = $false
	)
	# Fix password length if minchars is more than length
	$BaseCount = 1
	if ($Upper)
	{
		$BaseCount += $MinUpper
	}
	if ($Lower)
	{
		$BaseCount += $MinLower
	}
	if ($Number)
	{
		$BaseCount += $MinNumber
	}
	if ($Symbol)
	{
		$BaseCount += $MinSymbol
	}
	if ($BaseCount -gt $Length)
	{
		$Length = $BaseCount
	}
	$remaining = $Length - $BaseCount
	
	# Creates character arrays for the different character classes, based on ASCII character values.
	[string]$charsLower = (97 .. 122 | %{ [Char]$_ })
	$charsLower = $charsLower -replace " ", ""
	[string]$charsUpper = (65 .. 90 | %{ [Char]$_ })
	$charsUpper = $charsUpper -replace " ", ""
	[string]$charsNumber = (48 .. 57 | %{ [Char]$_ })
	$charsNumber = $charsNumber -replace " ", ""
	[string]$charsSymbol = (35, 36, 40, 41, 42, 44, 45, 46, 47, 58, 59, 63, 64, 92, 95 | %{ [Char]$_ })
	$charsSymbol = $charsSymbol -replace " ", ""
	
	# Create character arrays for non ambiguous characters l (ell), 1 (one), I (capital i), O (capital o), 0 (zero), B (capital b), 8 (eight), q (queue), g (gee), | (pipe)
	[string]$charsLowera = "abcdefhijkmnoprstuvwxyz"
	$charsLowera = $charsLowera -replace " ", ""
	[string]$charsUppera = "ACDEFGHJKLMNPQRTUVWXYZ"
	$charsUppera = $charsUppera -replace " ", ""
	[string]$charsNumbera = "234679"
	$charsNumbera = $charsNumbera -replace " ", ""
	[string]$charsSymbola = "!@#$%^&*()-=+:,?_.~"
	$charsSymbola = $charsSymbola -replace " ", ""
	
	
	if ($NonAmbiguous -eq $false)
	{
		if ($Upper)
		{
			$RandomUString = Get-RandomCharacters -length $MinUpper -characters $charsUpper
		}
		else
		{
			$RandomUString = ""
		}
		if ($Lower)
		{
			$RandomLString = Get-RandomCharacters -length $MinLower -characters $charsLower
		}
		else
		{
			$RandomLString = ""
		}
		if ($Number)
		{
			$RandomNString = Get-RandomCharacters -length $MinNumber -characters $charsNumber
		}
		else
		{
			$RandomNString = ""
		}
		if ($Symbol)
		{
			$RandomSString = Get-RandomCharacters -length $MinSymbol -characters $charsSymbol
		}
		else
		{
			$RandomSString = ""
		}
		
		
		
		if ($remaining -gt 0)
		{
			$Tempc = ""
			if ($Upper)
			{
				$Tempc += $charsUpper
			}
			if ($Lower)
			{
				$Tempc += $charsLower
			}
			if ($Number)
			{
				$Tempc += $charsNumber
			}
			if ($Symbol)
			{
				$Tempc += $charsSymbol
			}
			$RandomString = Get-RandomCharacters -length $remaining -characters $Tempc
		}
		else
		{
			$RandomString = ""
		}
	}
	else
	{
		if ($Upper)
		{
			$RandomUString = Get-RandomCharacters -length $MinUpper -characters $charsUppera
		}
		else
		{
			$RandomUString = ""
		}
		if ($Lower)
		{
			$RandomLString = Get-RandomCharacters -length $MinLower -characters $charsLowera
		}
		else
		{
			$RandomLString = ""
		}
		if ($Number)
		{
			$RandomNString = Get-RandomCharacters -length $MinNumber -characters $charsNumbera
		}
		else
		{
			$RandomNString = ""
		}
		if ($Symbol)
		{
			$RandomSString = Get-RandomCharacters -length $MinSymbol -characters $charsSymbola
		}
		else
		{
			$RandomSString = ""
		}
		if ($remaining -gt 0)
		{
			$Tempc = ""
			if ($Upper)
			{
				$Tempc += $charsUppera
			}
			if ($Lower)
			{
				$Tempc += $charsLowera
			}
			if ($Number)
			{
				$Tempc += $charsNumbera
			}
			if ($Symbol)
			{
				$Tempc += $charsSymbola
			}
			$RandomString = Get-RandomCharacters -length $remaining -characters $Tempc
		}
		else
		{
			$RandomString = ""
		}
	}
	
	#combine all into a single string
	$Return = $RandomUString + $RandomLString + $RandomNString + $RandomSString + $RandomString
	
	#scramble the scring
	$Return = Scramble-String -inputString $Return
	
	return $Return
	
	

}

