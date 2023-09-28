<#	
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2019 v5.6.166
	 Created on:   	8/6/2019 11:59 AM
	 Modified on:	1/21/2022 10:16 AM
	 Created by:   	Gary Cook
	 Organization: 	Quest
	 Filename:     	PSCommonCore.psm1
	-------------------------------------------------------------------------
	 Module Name: PSCommonCore
	===========================================================================
	#requires the PShellLogging module found on the Powershell Gallery
#>




#.EXTERNALHELP PSCommonCore.psm1-Help.xml
function Measure-IOPS
{
	param
	(
		[parameter(Mandatory = $true)]
		[string]$Speed,
		[parameter(Mandatory = $false)]
		[string]$SectorSize = "4K"
	)
	BEGIN
	{
		
	}
	PROCESS
	{
		# Convert Speed to value and type
		[double]$Value = $Speed -replace '[^0-9,.]', ''
		$type = $Speed -replace '[0-9,., ]', ''
		$size = $type.substring(0, 1)
		#check to see if the B is Capital
		if ($type -cmatch "B")
		{
			[int]$factor = 1
		}
		else
		{
			[int]$factor = 8
		}
		
		switch -casesensitive ($size)
		{
			b {
				$Value = $Value /$factor / 1024 / 1024
			}
			B {
				$value = $Value /$factor / 1024 / 1024
			}
			k {
				$Value = $Value /$factor / 1024
			}
			K {
				$Value = $Value /$factor / 1024
			}
			M {
				$Value = $Value /$factor
			}
			m {
				$Value = $Value /$factor
			}
			g {
				$Value = $Value /$factor * 1024
			}
			G {
				$Value = $Value /$factor * 1024
			}
			
		}
		
		[int]$Sector = $SectorSize -replace '[^0-9]', ''
		$Sector = $Sector * 1024
		Write-Host "Calulating for Sector size of $($Sector)"
		Write-Host "A throughput of $($Speed) in IOPS is:"
		[double]$IOPS = ($Value / ($sector / 1024)) * 1024
		Write-Host "$([math]::Round($IOPS, 2))"
		
	}
	END
	{
		
	}
	
}

#.EXTERNALHELP PSCommonCore.psm1-Help.xml
Function Test-Module
{
	param
	(
		[parameter(Mandatory = $true)]
		[string]$Module,
		[parameter(Mandatory = $true)]
		[bool]$Load = $false,
		[parameter(Mandatory = $true)]
		[bool]$Install = $false
		
	)
	[CmdletBinding]
	$Return = New-Object System.Management.Automation.PSObject
	$Return | Add-Member -MemberType NoteProperty -Name Status -Value $null
	$Return | Add-Member -MemberType NoteProperty -Name Message -Value $null
	
	$exit = $false
	
	If (!(Get-Module -ListAvailable $Module) -and $Install -eq $false)
	{
		$Return.Status = "Not Loaded"
		$Return.Message = "The Module $($Module) was not available to load"
		$exit = $true
	}
	if (!(Get-Module -ListAvailable $module) -and $Install -eq $true)
	{
		Install-Module $Module -Force -AllowClobber
		Import-Module $module -ea SilentlyContinue
		$Return.Status = "Loaded"
		$Return.Message = "The Module $($Module) was Installed and Loaded"
		$exit = $true
	}
	
	
	if (!(get-module $Module) -and $exit -eq $false)
	{
		if ($Load -eq $true)
		{
			Import-Module $module -ErrorAction SilentlyContinue
			if (!(Get-Module $module))
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
	else
	{
		$Return.Status = "Loaded"
		$Return.Message = "Module $($Module) was already loaded"
	}
	return $Return
}

#.EXTERNALHELP PSCommonCore.psm1-Help.xml
function Test-PSVersion
{
	param
	(
		[parameter(Mandatory = $true)]
		[string]$Min,
		[parameter(Mandatory = $true)]
		[ValidateSet("Desktop", "Core")]
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

#.EXTERNALHELP PSCommonCore.psm1-Help.xml
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

#.EXTERNALHELP PSCommonCore.psm1-Help.xml
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

#.EXTERNALHELP PSCommonCore.psm1-Help.xml
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

#.EXTERNALHELP PSCommonCore.psm1-Help.xml
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

#.EXTERNALHELP PSCommonCore.psm1-Help.xml
function New-RandomString([string]$inputString)
{
	$characterArray = $inputString.ToCharArray()
	$scrambledStringArray = $characterArray | Get-Random -Count $characterArray.Length
	$outputString = -join $scrambledStringArray
	return $outputString
}

#.EXTERNALHELP PSCommonCore.psm1-Help.xml
function Get-RandomCharacters($length, $characters)
{
	$random = 1 .. $length | ForEach-Object { Get-Random -Maximum $characters.length }
	$private:ofs = ""
	return [String]$characters[$random]
}

#.EXTERNALHELP PSCommonCore.psm1-Help.xml
function New-password
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
	if ($Upper -eq $false -and $Lower -eq $false -and $Number -eq $false -and $Symbol -eq $false)
	{
		Write-Error "Must Specify at least one character type for password i.e. -lower in call to new-password"
		
		
	}
	else
	{
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
		$Return = New-RandomString -inputString $Return
		
		return $Return
		
	}
	
}

#.EXTERNALHELP PSCommonCore.psm1-Help.xml
function get-loggedonuser ()
{
	[CmdletBinding()]
	Param
	(
		[parameter (position = 0, Mandatory = $false, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
		[string]$computername,
		[Parameter (Position = 1, Mandatory = $false, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
		[ValidateSet($true, $false)]
		[bool]$local = $false,
		[Parameter (Position = 2, Mandatory = $false, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
		[System.Management.Automation.PSCredential]$credential
	)
	Begin
	{
	}
	Process
	{
		
		
		if ($local -eq $false -and $computername -eq $null)
		{
			$computername = Read-Host "Remote computername was not provided please enter remote computername"
			if ($computername -eq "")
			{
				Write-Host "Computername not provided exiting"
				end
			}
		}
		if ($credential -eq $null)
		{
			write-host "User Credential not provided using logged on user"
			$localuser = $true
		}
		
		$regexa = '.+Domain="(.+)",Name="(.+)"$'
		$regexd = '.+LogonId="(\d+)"$'
		
		$logontype = @{
			"0"  = "Local System"
			"2"  = "Interactive" #(Local logon)
			"3"  = "Network" # (Remote logon)
			"4"  = "Batch" # (Scheduled task)
			"5"  = "Service" # (Service account logon)
			"7"  = "Unlock" #(Screen saver)
			"8"  = "NetworkCleartext" # (Cleartext network logon)
			"9"  = "NewCredentials" #(RunAs using alternate credentials)
			"10" = "RemoteInteractive" #(RDP\TS\RemoteAssistance)
			"11" = "CachedInteractive" #(Local w\cached credentials)
		}
		switch ($local)
		{
			$true {
				$logon_sessions = @(gwmi win32_logonsession -ComputerName "localhost")
				$logon_users = @(gwmi win32_loggedonuser -ComputerName "localhost")
			}
			$false {
				if ($localuser -ne $true)
				{
					$logon_sessions = @(gwmi win32_logonsession -ComputerName $computername -Credential $credential)
					$logon_users = @(gwmi win32_loggedonuser -ComputerName $computername -Credential $credential)
				}
				else
				{
					$logon_sessions = @(gwmi win32_logonsession -ComputerName $computername)
					$logon_users = @(gwmi win32_loggedonuser -ComputerName $computername)
				}
				
			}
			
		}
		
		
		$session_user = @{ }
		
		$logon_users | % {
			$_.antecedent -match $regexa > $nul
			$username = $matches[1] + "\" + $matches[2]
			$_.dependent -match $regexd > $nul
			$session = $matches[1]
			$session_user[$session] += $username
		}
		
		
		$logon_sessions | %{
			$starttime = [management.managementdatetimeconverter]::todatetime($_.starttime)
			
			$loggedonuser = New-Object -TypeName psobject
			$loggedonuser | Add-Member -MemberType NoteProperty -Name "Session" -Value $_.logonid
			$loggedonuser | Add-Member -MemberType NoteProperty -Name "User" -Value $session_user[$_.logonid]
			$loggedonuser | Add-Member -MemberType NoteProperty -Name "Type" -Value $logontype[$_.logontype.tostring()]
			$loggedonuser | Add-Member -MemberType NoteProperty -Name "Auth" -Value $_.authenticationpackage
			$loggedonuser | Add-Member -MemberType NoteProperty -Name "StartTime" -Value $starttime
			
			$loggedonuser
		}
	}
	end
	{
		
	}
	
}

#.EXTERNALHELP PSCommonCore.psm1-Help.xml
function Test-Powershell
{
	$obj = New-Object System.Management.Automation.PSObject
	$obj | Add-Member -MemberType NoteProperty -Name Version -Value "$($PSVersionTable.psversion.major).$($PSVersionTable.psversion.minor)"
	$obj | Add-Member -MemberType NoteProperty -Name Major -Value $PSVersionTable.psversion.major
	$obj | Add-Member -MemberType NoteProperty -Name Minor -Value $PSVersionTable.psversion.minor
	$obj | Add-Member -MemberType NoteProperty -Name Build -Value $PSVersionTable.psversion.build
	if ($PSVersionTable.psedition -eq $null)
	{
		$obj | Add-Member -MemberType NoteProperty -Name Edition -Value "Desktop"
	}
	else
	{
		$obj | Add-Member -MemberType NoteProperty -Name Edition -Value $PSVersionTable.psedition
	}
	
	if ($obj.edition -eq 'Core')
	{
		$obj | Add-Member -MemberType NoteProperty -Name Platform -Value $PSVersionTable.platform
	}
	else
	{
		$obj | Add-Member -MemberType NoteProperty -Name Platform -Value "Win32NT"
	}
	$obj | Add-Member -MemberType NoteProperty -Name Elevated -Value ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
	if ($obj.Elevated -eq $true)
	{
		$obj | Add-Member -MemberType NoteProperty -Name RemotingEnabled -Value (Test-PsRemoting)
	}
	else
	{
		$obj | Add-Member -MemberType NoteProperty -Name RemotingEnabled -Value "N/A"
	}
	
	
	return $obj
}

#.EXTERNALHELP PSCommonCore.psm1-Help.xml
function Test-PsRemoting
{
	param (
		[Parameter(Mandatory = $false)]
		$computername = "Localhost",
		[Parameter(Mandatory = $false)]
		[switch]$Auth = $false,
		[Parameter(Mandatory = $false)]
		[System.Management.Automation.PSCredential]$credential
	)
	
	try
	{
		$errorActionPreference = "Stop"
		if ($Auth)
		{
			if ($credential -eq $null)
			{
				$credential = Get-Credential
			}
			$result = Invoke-Command -ComputerName $computername -Credential $credential -scriptblock { 1 }
		}
		else
		{
			$result = Invoke-Command -ComputerName $computername -scriptblock { 1 }
		}
		
	}
	catch
	{
		Write-Verbose $_
		return $false
	}
	
	## I've never seen this happen, but if you want to be
	## thorough....
	if ($result -ne 1)
	{
		Write-Verbose "Remoting to $computerName returned an unexpected result."
		return $false
	}
	
	$true
}

#.EXTERNALHELP PSCommonCore.psm1-Help.xml
function get-pagefilelocation
{
	param (
		[Parameter(Mandatory = $false)]
		$computername,
		[Parameter(Mandatory = $false)]
		[System.Management.Automation.PSCredential]$credential
	)
	
	if ($computername -eq $null)
	{
		$computername = $env:computername
	}
	
	if ($credential -eq $null)
	{
		
		$loc = (Get-WmiObject -ComputerName $computername -Class Win32_PageFileUsage | select name).name
		return $loc
	}
	else
	{
		$loc = (Get-WmiObject -ComputerName $computername -Class Win32_PageFileUsage -Credential $credential | select name).name
		return $loc
	}
	
	
}

#.EXTERNALHELP PSCommonCore.psm1-Help.xml
function convertto-e164
{
	param
	(
		[parameter(Mandatory = $true)]
		[string]$Number		
	)
	
	[CmdletBinding]
	
	#note this is somewhat difficult as phone numbers are complex this was designed for converting mostly AD phone fields for use with systems that require E164 standard
	#remove all spaces, (, ), -
	$TNumber = $Number -replace "\s|\(|\)|-", ""
	#check if there is an x in the number and replace with ;Ext=
	$TNumber = $TNumber -replace "x", ";Ext="
	#check if the first character is a +
	if ($TNumber -notmatch "^\+")
		{
			#add a plus sign
			$TNumber = "+" + $TNumber
		}
	#now the hard part dealing with country codes
	
}

#.EXTERNALHELP PSCommonCore.psm1-Help.xml
function create-key
{
	param (
		[Parameter(Mandatory = $True)]
		[string]$KeyString
		
	)
	#Write-Host "Processing keystring $($KeyString)"
	#Write-Host "KeyString length $($KeyString.Length)"
	
	if ($KeyString.Length -lt 16 -or $KeyString.Length -gt 32)
	{
		Write-Host "Key is not between 16 and 32 characters must be 128-256 bits returning -1"
		return -1
		
	}
	else
	{
		$Ret = New-Object Byte[] 32
		for ($i = 0; $i -lt $KeyString.Length; $i += 1)
		{
			#Write-Host "Saving current value $($KeyString.substring($i,1)) in position $($i)"
			$ret[$i] = [int][char]($KeyString.Substring($i,1))
		}
		$fill = $KeyString.Length
		for ($i = $fill; $i -lt 32; $i += 1)
		{
			$ret[$i] = [int]0
			
		}
		return $ret
	}
	
}

#.EXTERNALHELP PSCommonCore.psm1-Help.xml
function encrypt-data
{
	
	param (
		[Parameter(Mandatory = $True)]
		[string]$Data,
		[Parameter(Mandatory = $True)]
		[byte[]]$Key
	)
	
	$value = (ConvertTo-SecureString -String $Data -AsPlainText -Force) | ConvertFrom-SecureString -Key $key
	return $value
		
}

#.EXTERNALHELP PSCommonCore.psm1-Help.xml
function decrypt-data
{
	param (
		[Parameter(Mandatory = $True)]
		[string]$Data,
		[Parameter(Mandatory = $True)]
		[byte[]]$Key
	)
	$tvalue = $data | ConvertTo-SecureString -Key $Key
	$value = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($tvalue))
	return $value
}

#.EXTERNALHELP PSCommonCore.psm1-Help.xml
function convertto-nicexml
{
	[CmdletBinding()]
	
	param (
		[Parameter(valuefrompipeline)]
		$InputObject,
		[Parameter (Mandatory = $false)]
		$RootNodeName = "Objects"
	)
	
	
	BEGIN 
	{
		[xml]$Doc = New-Object System.Xml.XmlDocument
		$null = $Doc.appendchild($Doc.CreateXmlDeclaration("1.0", "UTF-8", $null))
		$root = $Doc.AppendChild($Doc.CreateElement($RootNodeName))
			
	}
	PROCESS
	{
		$ChildObject = $Doc.CreateElement($InputObject.gettype().name)
		foreach ($propitem in $InputObject.psobject.properties)
		{
			$PropNode = $Doc.CreateElement($propitem.name)
			$PropNode.InnerText = $propitem.value
			$null = $ChildObject.AppendChild($PropNode)
		}
		$null = $root.AppendChild($ChildObject)
	}
	END
	{
		return $Doc.OuterXml
	}
}

#.EXTERNALHELP PSCommonCore.psm1-Help.xml
function generate-machinekey
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory=$true)]
		$EmailAddress,
		[Parameter(Mandatory = $false)]
		$FileName = $null
	)
	
	
	BEGIN
	{
		#gather information
		$HostName = [environment]::machinename
		
		$IP = (Get-netadapter | sort -Property ifindex | ?{ $_.status -eq 'Up' } | select -First 1 | Get-NetIPAddress -AddressFamily ipv4).ipaddress
		$MACAddress = (Get-netadapter | sort -Property ifindex | ?{ $_.status -eq 'Up' } | select Name, MACAddress -first 1).macaddress
		$OSSerial = (Get-WmiObject -clas win32_operatingsystem | select serialnumber).serialnumber
		$date = Get-Date -Format MMddyyyy
		$key = "ThisIsThePublicQuestKey11!!"
		$EKey = create-key -KeyString $key
	}
	PROCESS
	{
		#create combined string
		$MString = "$($EmailAddress)|$($date)|$($HostName)|$($OSSerial)|$($IP)|$($MACAddress)"
		$EString = encrypt-data -Data $MString -Key $EKey
	}
	END
	{
		if ($FileName -ne $null)
		{
			$EString | Out-File -FilePath $FileName
		}
		else
		{
			return $EString
		}
		
	}
}

#.EXTERNALHELP PSCommonCore.psm1-Help.xml
function get-serviceinfo ()
{
	[cmdletbinding()]
	[OutputType("ServiceInfo", "String")]
	param
	(
		[Parameter(Mandatory = $false)]
		[string]$Computername = "Localhost",
		[Parameter(Mandatory = $false)]
		[validateset('All','Single')]
		[string]$Scope = "All",
		[Parameter(Mandatory = $false)]
		[string]$servicename,
		[Parameter(Mandatory = $false)]
		[System.Management.Automation.PSCredential]$credential
		
	)
	$return = @()
	#test is computer is available
	if ((Test-NetConnection -ComputerName $Computername -port 135).TCPTESTSUCCEEDED -eq $TRUE)
	{
		if ($Scope -ne "All")
		{
			if ($credential -eq $null)
			{
				$services = gwmi Win32_Service -ComputerName $Computername | ?{ $_.name -eq $servicename }
			}
			else
			{
				$services = gwmi Win32_Service -ComputerName $Computername -Credential $credential| ?{ $_.name -eq $servicename }
			}
			
		}
		else
		{
			if ($credential -eq $null)
			{
				$services = gwmi Win32_Service -ComputerName $computername
			}
			else
			{
				$services = gwmi Win32_Service -ComputerName $computername -Credential $credential
			}
			
		}
		foreach ($service in $services)
		{
			$return += [pscustomObject]@{
				PSTypename  = "ServiceInfo"
				Systemname  = $Computername
				ServiceName = $service.name
				Displayname = $service.displayname
				State	    = $service.state
				Pathname    = $service.pathname
				StartMode   = $service.startmode
				DelayedAutoStart = $service.delayedautostart
				Username    = $service.startname
				
			}
			 
		}
		
	}
	$return
}

#.EXTERNALHELP PSCommonCore.psm1-Help.xml
function Sign-Script ()
{
	param
	(
		[parameter(Mandatory = $false, Position = 1, ValueFromPipelineByPropertyName = $true)]
		[string]$script = $null,
		[parameter(Mandatory = $false, Position = 2, ValueFromPipelineByPropertyName = $true)]
		[string]$certificatethumbprint = $null
	)
	if ($script.Length -eq 0)
	{
		$script = ""
	}
	
	if ($certificatethumbprint = "")
	{
		$certificate = $null
	}
	else
	{
		$certificate = Get-ChildItem -Path cert: -Recurse | ?{ $_.thumbprint -eq $certificatethumbprint }
		if ($certificate -eq $null)
		{
			Write-Color -Text "Cert not found" -Color red
			$certificate = $null
		}
		
	}
	
	#if certificate is not provided provide list of certificates and allow user to select
	if ($certificate -eq $null)
	{
		$number = 0
		$choice = -1
		
		do
		{
			cls
			Write-Color -Text "Select to correct code signing certificate.`r`n`r`n"
			#get all code signing certificates in user certificate store
			$certs = @(Get-ChildItem cert:\CurrentUser\My -codesigning)
			$certs = $certs + @(Get-ChildItem cert:\Localmachine\My -codesigning)
			$location = ($cert.psparentpath -split "::")[1]
			#display list of certs for user selection
			$number = 0
			foreach ($cert in $certs)
			{
				$commonname = $cert.subject -split "," | select -First 1
				
				Write-Color -Text "[", "$($number)", "]", " CertCommonName: ", "$($commonname)", ", CertIssuer: ", "$($cert.issuer)", " Location: ", "$($location)" -Color white, red, white, yellow, green, yellow, green, yellow, green
				$number += 1
			}
			
			#get user choice
			$choice = Read-Host "Select the certificate to use"
		}
		while ([int]$choice -lt 0 -or [int]$choice -gt $number)
		
		$certificate = $certs[$choice]
	}
	
	if ($script -eq "")
	{
		$script = Read-Host "Enter Complete path to script to sign"
	}
	try
	{
		Set-AuthenticodeSignature $script $certificate -ea Stop
	}
	catch
	{
		Write-Color -Text "There was an erro processing script ","$($script)",", The error message is ","$($_.ErrorDetails.Message)" -color red,yellow,red,yellow
	}
}

<# Function removed due to MS creating good functions in existing cloud modules

#.EXTERNALHELP PSCommonCore.psm1-Help.xml
Function Set-Owner
{
    <#
        .SYNOPSIS
            Changes owner of a file or folder to another user or group.

        .DESCRIPTION
            Changes owner of a file or folder to another user or group.

		.PARAMETER Path
            The folder or file that will have the owner changed.

        .PARAMETER Account
            Optional parameter to change owner of a file or folder to specified account.

            Default value is 'Builtin\Administrators'

        .PARAMETER Recurse
            Recursively set ownership on subfolders and files beneath given folder.

        .NOTES
            Name: Set-Owner
            Author: Boe Prox
            Version History:
                 1.0 - Boe Prox
                    - Initial Version
				 2.0 - Gary Cook
					- Added PShellLogging to capture Success and Error Information
					- Fixed issues processing owner on specific server OS's
	
			Requires the PSHellLogging Modules Available on PowerShell gallery
			https://www.powershellgallery.com/packages/PShellLogging/1.1.13
			Use "Install-Module -Name PShellLogging" on powershell 5.0 and higher to install

        .EXAMPLE
            Set-Owner -Path C:\temp\test.txt -Log c:\logs\ownerlog.csv -logtype CSV

            Description
            -----------
            Changes the owner of test.txt to Builtin\Administrators, Logs output to c:\logs\ownerlog.csv in comma Seperated Format

        .EXAMPLE
            Set-Owner -Path C:\temp\test.txt -Account 'Domain\bprox -Log c:\logs\ownerlog.TXT -logtype TXT

            Description
            -----------
            Changes the owner of test.txt to Domain\bprox, Logs output to c:\logs\ownerlog.csv in text format

        .EXAMPLE
            Set-Owner -Path C:\temp -Recurse -Log c:\logs\ownerlog.csv -logtype CSV

            Description
            -----------
            Changes the owner of all files and folders under C:\Temp to Builtin\Administrators, Logs output to c:\logs\ownerlog.csv in comma Seperated Format

        .EXAMPLE
            Get-ChildItem C:\Temp | Set-Owner -Recurse -Account 'Domain\bprox' -Log c:\logs\ownerlog.csv -logtype CSV

            Description
            -----------
            Changes the owner of all files and folders under C:\Temp to Domain\bprox, Logs output to c:\logs\ownerlog.csv in comma Seperated Format
    #>
<#
[cmdletbinding(
			   SupportsShouldProcess = $True
			   )]
Param (
	[parameter(mandatory = $true, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
	[Alias('FullName')]
	[string[]]$Path,
	[parameter()]
	[string]$Account = 'Builtin\Administrators',
	[parameter()]
	[switch]$Recurse,
	[parameter()]
	[string]$Log,
	[parameter()]
	[ValidateSet("TXT", "CSV", "JSON")]
	[string]$logtype = "CSV"
)
Begin
{
	#Create Log if necessary
	if ($Log -ne $null)
	{
		$MyLog = Start-Log -Log $Log -Type $logtype
		$logging = $true
	}
	#Prevent Confirmation on each Write-Debug command when using -Debug
	If ($PSBoundParameters['Debug'])
	{
		$DebugPreference = 'Continue'
	}
	Try
	{
		[void][TokenAdjuster]
	}
	Catch
	{
		$AdjustTokenPrivileges = @"
            using System;
            using System.Runtime.InteropServices;

             public class TokenAdjuster
             {
              [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
              internal static extern bool AdjustTokenPrivileges(IntPtr htok, bool disall,
              ref TokPriv1Luid newst, int len, IntPtr prev, IntPtr relen);
              [DllImport("kernel32.dll", ExactSpelling = true)]
              internal static extern IntPtr GetCurrentProcess();
              [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
              internal static extern bool OpenProcessToken(IntPtr h, int acc, ref IntPtr
              phtok);
              [DllImport("advapi32.dll", SetLastError = true)]
              internal static extern bool LookupPrivilegeValue(string host, string name,
              ref long pluid);
              [StructLayout(LayoutKind.Sequential, Pack = 1)]
              internal struct TokPriv1Luid
              {
               public int Count;
               public long Luid;
               public int Attr;
              }
              internal const int SE_PRIVILEGE_DISABLED = 0x00000000;
              internal const int SE_PRIVILEGE_ENABLED = 0x00000002;
              internal const int TOKEN_QUERY = 0x00000008;
              internal const int TOKEN_ADJUST_PRIVILEGES = 0x00000020;
              public static bool AddPrivilege(string privilege)
              {
               try
               {
                bool retVal;
                TokPriv1Luid tp;
                IntPtr hproc = GetCurrentProcess();
                IntPtr htok = IntPtr.Zero;
                retVal = OpenProcessToken(hproc, TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, ref htok);
                tp.Count = 1;
                tp.Luid = 0;
                tp.Attr = SE_PRIVILEGE_ENABLED;
                retVal = LookupPrivilegeValue(null, privilege, ref tp.Luid);
                retVal = AdjustTokenPrivileges(htok, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
                return retVal;
               }
               catch (Exception ex)
               {
                throw ex;
               }
              }
              public static bool RemovePrivilege(string privilege)
              {
               try
               {
                bool retVal;
                TokPriv1Luid tp;
                IntPtr hproc = GetCurrentProcess();
                IntPtr htok = IntPtr.Zero;
                retVal = OpenProcessToken(hproc, TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, ref htok);
                tp.Count = 1;
                tp.Luid = 0;
                tp.Attr = SE_PRIVILEGE_DISABLED;
                retVal = LookupPrivilegeValue(null, privilege, ref tp.Luid);
                retVal = AdjustTokenPrivileges(htok, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
                return retVal;
               }
               catch (Exception ex)
               {
                throw ex;
               }
              }
             }
"@
		Add-Type $AdjustTokenPrivileges
	}
	
	#Activate necessary admin privileges to make changes without NTFS perms
	[void][TokenAdjuster]::AddPrivilege("SeRestorePrivilege") #Necessary to set Owner Permissions
	[void][TokenAdjuster]::AddPrivilege("SeBackupPrivilege") #Necessary to bypass Traverse Checking
	[void][TokenAdjuster]::AddPrivilege("SeTakeOwnershipPrivilege") #Necessary to override FilePermissions
}
Process
{
	ForEach ($Item in $Path)
	{
		Write-Verbose "FullName: $Item"
		if ($logging)
		{
			$MyLog | Write-Log -Line "Proicessing item $($Item)" -Level Info
		}
		#The ACL objects do not like being used more than once, so re-create them on the Process block
		$DirOwner = New-Object System.Security.AccessControl.DirectorySecurity
		$DirOwner.SetOwner([System.Security.Principal.NTAccount]$Account)
		$FileOwner = New-Object System.Security.AccessControl.FileSecurity
		$FileOwner.SetOwner([System.Security.Principal.NTAccount]$Account)
		$DirAdminAcl = New-Object System.Security.AccessControl.DirectorySecurity
		$FileAdminAcl = New-Object System.Security.AccessControl.DirectorySecurity
		$AdminACL = New-Object System.Security.AccessControl.FileSystemAccessRule('Builtin\Administrators', 'FullControl', 'ContainerInherit,ObjectInherit', 'InheritOnly', 'Allow')
		$FileAdminAcl.AddAccessRule($AdminACL)
		$DirAdminAcl.AddAccessRule($AdminACL)
		Try
		{
			$Item = Get-Item -LiteralPath $Item -Force -ErrorAction Stop
			If (-NOT $Item.PSIsContainer)
			{
				If ($PSCmdlet.ShouldProcess($Item, 'Set File Owner'))
				{
					Try
					{
						$Item.SetAccessControl($FileOwner)
					}
					Catch
					{
						Write-Warning "Couldn't take ownership of $($Item.FullName)! Taking FullControl of $($Item.Directory.FullName)"
						$Item.Directory.SetAccessControl($FileAdminAcl)
						$Item.SetAccessControl($FileOwner)
					}
				}
			}
			Else
			{
				If ($PSCmdlet.ShouldProcess($Item, 'Set Directory Owner'))
				{
					Try
					{
						$Item.SetAccessControl($DirOwner)
					}
					Catch
					{
						Write-Warning "Couldn't take ownership of $($Item.FullName)! Taking FullControl of $($Item.Parent.FullName)"
						$Item.Parent.SetAccessControl($DirAdminAcl)
						$Item.SetAccessControl($DirOwner)
					}
				}
				If ($Recurse)
				{
					[void]$PSBoundParameters.Remove('Path')
					Get-ChildItem $Item -Force | Set-Owner @PSBoundParameters
				}
			}
			if ($logging)
			{
				$MyLog | Write-Log -Line "Item $($Item) was successfully processed" -Level Info
			}
		}
		Catch
		{
			Write-Warning "$($Item): $($_.Exception.Message)"
			if ($logging)
			{
				$MyLog | Write-Log -Line "Item $($Item) processing failed error message $($_.Exception.Message)" -Level Error
			}
		}
	}
}
End
{
	#Remove priviledges that had been granted
	[void][TokenAdjuster]::RemovePrivilege("SeRestorePrivilege")
	[void][TokenAdjuster]::RemovePrivilege("SeBackupPrivilege")
	[void][TokenAdjuster]::RemovePrivilege("SeTakeOwnershipPrivilege")
	if ($logging)
	{
		$MyLog | Close-Log
	}
}
}


#.EXTERNALHELP PSCommonCore.psm1-Help.xml
function Test-ADUserExists
{
[CmdletBinding()]
[OutputType([object])]
param
(
	[Parameter(Mandatory = $false,
			   ValueFromPipeline = $true,
			   ValueFromPipelineByPropertyName = $true)]
	[Alias('DC')]
	[String]$TargetDC,
	[Parameter(Mandatory = $false,
			   ValueFromPipeline = $true,
			   ValueFromPipelineByPropertyName = $true)]
	[string]$ADProperty = 'SAMAccountName',
	[Parameter(Mandatory = $true,
			   ValueFromPipeline = $true,
			   ValueFromPipelineByPropertyName = $true)]
	[String]$Value
)

Begin
{
	if ($TargetDC -eq "")
	{
		$TargetDC = (Get-ADDomainController -Discover).hostname
	}
}
Process
{
	try
	{
		$obj = New-Object System.Management.Automation.PSObject
		$propfilter = "SID", "$($ADProperty)"
		$users = Get-ADUser -Filter * -Properties $propfilter -Server $TargetDC | ?{ $_.$ADProperty -eq $Value } -ea Stop
		$count = ($users | measure).count
		if ($count -ne 1)
		{
			if ($count -ne 0)
			{
				$obj | Add-Member -MemberType NoteProperty -Name SID -Value $users[0].sid.value
				$obj | Add-Member -MemberType NoteProperty -Name Count -Value $count
				$obj | Add-Member -MemberType NoteProperty -Name Error -Value "None"
			}
			else
			{
				$obj | Add-Member -MemberType NoteProperty -Name SID -Value ""
				$obj | Add-Member -MemberType NoteProperty -Name Count -Value $count
				$obj | Add-Member -MemberType NoteProperty -Name Error -Value "None"
			}
			
		}
		else
		{
			$obj | Add-Member -MemberType NoteProperty -Name SID -Value $users.sid.value
			$obj | Add-Member -MemberType NoteProperty -Name Count -Value $count
			$obj | Add-Member -MemberType NoteProperty -Name Error -Value "None"
		}
	}
	catch
	{
		$obj | Add-Member -MemberType NoteProperty -Name SID -Value ""
		$obj | Add-Member -MemberType NoteProperty -Name Count -Value -1
		$obj | Add-Member -MemberType NoteProperty -Name Error -Value $_.ErrorDetails.Message
	}
	
	
}
End
{
	return $obj
}
}

#.EXTERNALHELP PSCommonCore.psm1-Help.xml
function export-AutodiscoverSPC
{
	param
	(
		[parameter(Mandatory = $false)]
		[string]$DomainDN
	)
	$obj = @()
	
	if ($DomainDN -eq "")
	{
		$ADDomain = Get-ADDomain | Select DistinguishedName
		$DomainDN = $ADDomain.distinguishedname
	}
	
	
	$DSSearch = New-Object System.DirectoryServices.DirectorySearcher
	$DSSearch.Filter = '(&(objectClass=serviceConnectionPoint)(|(keywords=67661d7F-8FC4-4fa7-BFAC-E1D7794C1F68)(keywords=77378F46-2C66-4aa9-A6A6-3E7A48B19596)))'
	$DSSearch.SearchRoot = 'LDAP://CN=Configuration,' + $DomainDN
	$DSSearch.FindAll() | %{
		
		$ADSI = [ADSI]$_.Path
		$autodiscover = New-Object psobject -Property @{
			Server = [string]$ADSI.cn
			Site   = $adsi.keywords[0]
			DateCreated = $adsi.WhenCreated.ToShortDateString()
			AutoDiscoverInternalURI = [string]$adsi.ServiceBindingInformation
		}
		$obj += $autodiscover
		
	}
	
	return $obj
}

#.EXTERNALHELP PSCommonCore.psm1-Help.xml
function Connect-AD
{
	param
	(
		[parameter(Mandatory = $False,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true, Position = 1)]
		[Alias ('DC')]
		[string]$TargetDC = "",
		[parameter(Mandatory = $False,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true, Position = 2)]
		[System.Management.Automation.PSCredential]$Credential
		
	)
	BEGIN
	{
		if ($TargetDC -eq "")
		{
			$SystemDomainInfo = gwmi -Class win32_ntdomain
			if ($SystemDomainInfo.domaincontrollername -ne $null)
			{
				$TargetDC = $SystemDomainInof.domaincontrollername -replace "\\", ""
				$TargetDC = "$($TargetDC).$($SystemDomainInfo.dnsforestname)"
				
			}
			else
			{
				Write-Error "This Computer is not domain joined and no TargetDC (DC) was supplied.  Error - Cannot connect to domain controller"
				end
				
			}
			
		}
		if ($Credential -eq $null)
		{
			$Credential = Get-Credential -Message "Please Enter you admin password for domain controller $($TargetDC)"
			
		}
	}
	PROCESS
	{
		try
		{
			$ErrorActionPreference = 'Stop'
			$session = New-PSSession -ComputerName $targetDC -Credential $Credential
			Invoke-Command $session -Scriptblock { Import-Module ActiveDirectory }
			#Import-PSSession -Session $session -AllowClobber
			$good = $true
		}
		catch
		{
			Write-Error "Failed to connect to domain controller $($TargetDC)"
			Write-Error "$($_.ErrorDetails.Message)"
			$good = $false
		}
		$session = Get-PSSession | ?{ $_.ComputerName -eq $TargetDC }
		
		
		#if ($good)
		#{
		#	Import-PSSession -Session $session -AllowClobber -module activedirectory
		#}
		return $session
	}
	END
	{
		
	}
}

#.EXTERNALHELP PSCommonCore.psm1-Help.xml
function Connect-Exchange
{
	param
	(
		[parameter(Mandatory = $false)]
		[System.Management.Automation.PSCredential]$Credential,
		[parameter(Mandatory = $false)]
		[ValidateSet ("EONPREM", "EOL")]
		[string]$Type = "EONPREM"
	)
	Begin
	{
		if ($Credential -eq $null)
		{
			$Credential = Get-Credential -Message "Please Enter your Exchange Admin Credential."
		}
		
	}
	Process
	{
		
		try
		{
			$ErrorActionPreference = 'Stop'
			if ($Type -eq "EOL")
			{
				$Session = New-PSSession -ConfigurationName "Microsoft.Exchange" -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $Credential -Authentication Basic -AllowRedirection
				
			}
			else
			{
				$server = Read-Host "Please enter On Prem Exchange Server Name"
				$Session = New-PSSession -ConfigurationName "Microsoft.Exchange" -ConnectionUri "https://$($server)/powershell/" -Credential $Credential -AllowRedirection
				
			}
			#Invoke-Command $session -Scriptblock {"1"}
			$good = $true
		}
		catch
		{
			Write-Error "Failed to connect to Exchange"
			Write-Error "$($_.ErrorDetails.Message)"
			$good = $false
		}
		#[int]$rtvalue = $Session.Id
		#Write-Host "The Session id is $($rtvalue)"
		#$null = Import-PSSession $Session -DisableNameChecking -AllowClobber
		#$session = Get-PSSession | ?{ $_.ConfigurationName -eq "Microsoft.Exchange" } | select -First 1
		
		
		#if ($good)
		#{
		#	
		#	Import-PSSession -Session $session -AllowClobber -DisableNameChecking
		#}
		
		return $session
	}
	END
	{
		
	}
}

#.EXTERNALHELP PSCommonCore.psm1-Help.xml
function Disconnect-Exchange
{
	param
	(
		[parameter(Mandatory = $true, ValueFromPipeline = $true)]
		[int]$SessionId
	)
	$session = Get-PSSession | ?{ $_.Id -eq $SessionId }
	Remove-PSSession -Session $Session
}

#>