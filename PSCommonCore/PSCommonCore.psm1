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




Function Validate-Module
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


