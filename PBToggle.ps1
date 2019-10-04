###################################################################################################
# 
# Program:     PBToggle.ps1
# Description: Powershell script that handles automated toggle between VisionStar
#              Application Versions 8.x to 2019.x. Vistar.ini config files are
#              Automated, Oracle folders are re-named for toggle to be sucessful
# 
# Version:     1.0 Initial Version
# 
###################################################################################################

$VAR_VistarIniFile = 'C:\Windows\vistar.ini'
$VAR_Version8 = 'Version 8'
$VAR_Version2019 = 'Version 2019.1'
$VAR_DBMS8 = 'O84 ORACLE 8.0.4'
$VAR_DBMS2019 = '"O10 Oracle10g (10.1.0)"'
$VAR_Database8 = 'vistar.world'
$VAR_Database2019 = 'nemo.world/nemo'
$VAR_CommitConfigChanges = $false
$VAR_FolderORANT = 'C:\orant'
$VAR_FolderORANT8 = 'C:\orant8'
$VAR_FolderORANT12 = 'C:\orant12'

Write-Host "`nSTARTING VisionStar script: PBToggle.ps1..."
Write-Host "Prompting User to select a VisionStar target version..."

$AppVersion = $VAR_Version8,$VAR_Version2019 | Out-GridView -OutputMode Single -Title "Please select the desired App version"

if ([string]::IsNullOrWhitespace($AppVersion))
{
	Write-Host "You cancelled the App version selection, ABORTING script operatons"
}
else
{
	Write-Host "You selected to configure App with: $AppVersion"

	Write-Host "Testing for existence of config file: $VAR_VistarIniFile..."
	$TestIniFile = Test-Path $VAR_VistarIniFile -PathType Leaf
	if (!$TestIniFile)
	{
		Write-Host "Could not find config file: $VAR_VistarIniFile, ABORTING script operatons"
	}
	else
	{
		Write-Host "`nReading content from config file: $VAR_VistarIniFile..."
		$VistarIni = Get-Content -Path $VAR_VistarIniFile

		Write-Host "Reading DBMS key=value in config file..."
		$VistarDBMS = $VistarIni | Where-Object { $_ -match 'DBMS=' -and $_ -notlike "*;*" }
		if ([string]::IsNullOrWhitespace($VistarDBMS))
		{
			Write-Host "Could not find DBMS key=value in config file: $VAR_VistarIniFile, ABORTING script operatons"
		}
		else
		{
			Write-Host "Current key/value of DBMS key is: $VistarDBMS"

			Write-Host "Reading DBConnectionLastUsed key=value in config file..."
			$VistarDatabase = $VistarIni | Where-Object { $_ -match 'DBConnectionLastUsed=' -and $_ -notlike "*;*" }
			if ([string]::IsNullOrWhitespace($VistarDatabase))
			{
				Write-Host "Could not find DBConnectionLastUsed key=value in config file: $VAR_VistarIniFile, ABORTING script operatons"
			}
			else
			{
				Write-Host "Current key/value of DBConnectionLastUsed is: $VistarDatabase"

				if ($AppVersion -eq $VAR_Version8)
				{
					Write-Host "`nSetting new key/value of DBMS to: DBMS=$VAR_DBMS8"
					$VistarIni = $VistarIni.Replace($VistarDBMS,"DBMS=$VAR_DBMS8")

					Write-Host "Setting new key/value of DBConnectionLastUsed to: DBConnectionLastUsed=$VAR_Database8"
					$VistarIni = $VistarIni.Replace($VistarDatabase,"DBConnectionLastUsed=$VAR_Database8")

					$VAR_CommitConfigChanges = $true
				}
				else
				{
					if ($AppVersion -eq $VAR_Version2019)
					{
						Write-Host "`nSetting new key/value of DBMS to: DBMS=$VAR_DBMS2019"
						$VistarIni = $VistarIni.Replace($VistarDBMS,"DBMS=$VAR_DBMS2019")

						Write-Host "Setting new key/value of DBConnectionLastUsed to: DBConnectionLastUsed=$VAR_Database2019"
						$VistarIni = $VistarIni.Replace($VistarDatabase,"DBConnectionLastUsed=$VAR_Database2019")

						$VAR_CommitConfigChanges = $true
					}
				}

				if (!$VAR_CommitConfigChanges)
				{
					Write-Host "VAR_CommitConfigChanges variable was not set to True, ABORTING script operatons"
				}
				else
				{
					Write-Host "`nSaving changes to config file..."
					Set-Content -Path 'C:\Windows\vistar.ini' -Value $VistarIni

					Write-Host "Saving changes to config file...COMPLETE"

					################################
					# Handling Oracle folder rename
					################################
					Write-Host "`nTesting for existence of Oracle Folder: $VAR_FolderORANT..."
					$TestORANT = Test-Path -Path $VAR_FolderORANT
					if (!$TestORANT)
					{
						Write-Host "Could not find expected folder: $VAR_FolderORANT, ABORTING script operatons"
					}
					else
					{
						Write-Host "Found expected folder: $VAR_FolderORANT"

						if ($AppVersion -eq $VAR_Version2019)
						{
							Write-Host "`nTesting for existence of Oracle 12 library file: C:\orant\oci.dll..."
							$TestOracle12File = Test-Path 'C:\orant\oci.dll' -PathType Leaf
							if (!$TestOracle12File)
							{
								Write-Host "Testing for existence of Oracle Folder: $VAR_FolderORANT12..."
								$TestORANT12 = Test-Path -Path $VAR_FolderORANT12
								if ($TestORANT12)
								{
									Write-Host "Renaming folder: C:\orant to C:\orant8..."
									Rename-Item 'C:\orant' 'C:\orant8'

									Write-Host "Renaming folder: C:\orant12 to C:\orant..."
									Rename-Item 'C:\orant12' 'C:\orant'

									Write-Host "Oracle 12 folder setup...COMPLETE"
								}
							}
							else
							{
								Write-Host "Oracle 12 folder is already in place, nothing needed"
							}
						}
						else
						{
							if ($AppVersion -eq $VAR_Version8)
							{
								Write-Host "`nTesting for existence of Oracle 8 library file: C:\orant\jdk.exe..."
								$TestOracle8File = Test-Path 'C:\orant\jdk.exe' -PathType Leaf
								if (!$TestOracle8File)
								{
									Write-Host "Testing for existence of Oracle Folder: $VAR_FolderORANT8..."
									$TestORANT8 = Test-Path -Path $VAR_FolderORANT8
									if ($TestORANT8)
									{
										Write-Host "Renaming folder: C:\orant to C:\orant12..."
										Rename-Item 'C:\orant' 'C:\orant12'

										Write-Host "Renaming folder: C:\orant8 to C:\orant..."
										Rename-Item 'C:\orant8' 'C:\orant'

										Write-Host "Oracle 8 folder setup...COMPLETE"
									}
								}
								else
								{
									Write-Host "Oracle 8 folder is already in place, nothing needed"
								}
							}
						}
					}
					####################################
					# END Handling Oracle folder rename
					####################################
				}

			}
		}
	}
}

Write-Host "`nEND VisionStar script: PBToggle.ps1`n"
# END program: PBToggle.ps1

