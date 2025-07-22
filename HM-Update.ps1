<#
.DESCRIPTION
PowerShell script to update Harbour Masters ports.
Gets update data and updates the port in the scripts current working directory.

.NOTES
Exit Codes:
0 - Success. Successful update, already up to date or decided not to update.
1 - Failed before update. Files are not altered.
2 - Failed during update. Some files might be altered and need to be rolled back manually from the "UpBackup" folder.
3 - Failed to pack backup. Update was successful, but packed backup is incomplete. Manually save files from the "UpBackup" folder before next update.


Harbour Masters Update - PowerShell Script v25.07.23
    
    MIT License

    Copyright (C) 2025 Spodi

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

.EXAMPLE
.\HM-Update.ps1 -AutoUpdate -RegenAssetArchive -DeleteRando
Do an unattended update that deletes asset archives and Randomizer saves (saves will be backed up).
.EXAMPLE
.\HM-Update.ps1 -AutoUpdate -RegenAssetArchive:$false -DeleteRando:$false
Do an unattended update that keeps asset archives Randomizer saves (saves will still be backed up).
.Example
.\HM-Update.ps1 -Version "Nightly"
Update to the latest Nightly.
.EXAMPLE
.\HM-Update.ps1 -ForceGame 'Ship of Harkinian'
Download and unpack the latest release of Ship of Harkinian. Best to only use this in an empty folder.
#>

[CmdletBinding(DefaultParameterSetName = 'Default')]
param (
    <#
    Update the port found in this path.
    Uses working directory when not given.
    #>
    [Parameter(ParameterSetName = 'Default')] [string[]]$Path,

    # Automatically starts the game after this script finishes. Except on errors.
    [Parameter(ParameterSetName = 'Default')] [switch]$Autostart,

    <#
    Force downloading the release with the given GitHub tag. Usually like "9.0.2" or "v1.0.0". Use "Nightly" for the latest Nightly build.
    Implies "-ForceUpdate" and "-SkipCheckTimer".
    #>
    [Parameter(ParameterSetName = 'Default')] [string]$Version,

    <#
    Assume "yes" for updating without asking.
    Also set -RegenAssetArchive or -RegenAssetArchive:$false and -DeleteRando or -DeleteRando:$false for an unattended update.
    #>
    [Parameter(ParameterSetName = 'Default')] [switch]$AutoUpdate,

    <#
    Assume "yes" for deleting asset archives without asking. Use -RegenAssetArchive:$false to assume "no".
    Also set -AutoUpdate and -DeleteRando or -DeleteRando:$false for an unattended update.
    #>
    [Parameter(ParameterSetName = 'Default')] [switch]$RegenAssetArchive,

    <#
    Assume "yes" for deleting Randomizer saves without asking. Use -DeleteRando:$false to assume "no".
    Also set -AutoUpdate and -RegenAssetArchive or -RegenAssetArchive:$false for an unattended update.
    #>
    [Parameter(ParameterSetName = 'Default')] [switch]$DeleteRando,

    # Check for updates more often than once per hour.
    [Parameter(ParameterSetName = 'Default')] [switch]$SkipCheckTimer,

    <#
    Force an update, regardless of your current version.
    Implies "-SkipCheckTimer".
    #>
    [Parameter(ParameterSetName = 'Default')] [switch]$ForceUpdate,

    <#
    Use update data for the selected game and force an update, no matter what is actually detected. Can be used to simply download and extract any of the ports.
    Implies "-ForceUpdate" and "-SkipCheckTimer".
    > [!CAUTION]
    > -ForceGame can corrupt your game if you select a different game than the one already installed!
    #>
    [Parameter(ParameterSetName = 'Default')] [string]$ForceGame
)

#region Functions
function Get-FileSystemEntries {
    <#
    .SYNOPSIS
    Basically "Get-ChildItem", but slightly faster. Gets paths only.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)] [string] $Path,
        [Parameter()] [switch] $File,
        [Parameter()] [switch] $Directory,
        [Parameter()] [switch] $Recurse
    )
    begin {
        $prevDir = [System.IO.Directory]::GetCurrentDirectory()
        [System.IO.Directory]::SetCurrentDirectory((Get-Location))
        $Queue = [System.Collections.Queue]@()
    }
    process {
        $Queue.Enqueue($Path)
        while ($Queue.count -gt 0) {
            try {
                $Current = $Queue.Dequeue()
                [System.IO.Directory]::EnumerateDirectories($Current) | & { process {
                        if (!$File) { Write-Output ($_ + [System.IO.Path]::DirectorySeparatorChar ) }
                        if ($Recurse) { $Queue.Enqueue($_) }
                    } }
                if (!$Directory) { Write-Output ([System.IO.Directory]::EnumerateFiles($Current)) }
            }
            catch [System.Management.Automation.RuntimeException] {
                $catchedError = $_
                switch ($catchedError.Exception.InnerException.GetType().FullName) {
                    'System.UnauthorizedAccessException' { Write-Warning $catchedError.Exception.InnerException.Message }
                    'System.Security.SecurityException' { Write-Warning $catchedError.Exception.InnerException.Message }
                    default {
                        Throw $catchedError
                    }
                }  
            }
        
        }
    }
    end {
        [System.IO.Directory]::SetCurrentDirectory($prevDir)
    }
}
function CopyAndBackup {
    [CmdletBinding()]
    Param ($path, $destination, $root, $backupPath)
	
    $items = Get-FileSystemEntries $path
    $items | ForEach-Object {
        $Name = Split-Path -Leaf $_
        $newPath =	Join-Path $path			$Name
        $newDestination	=	Join-Path $destination	$Name
        $newBackupPath	=	Join-Path $backupPath	$Name
		
        if ((Test-Path $_ -PathType Container)) {

            if (-not (Test-Path -LiteralPath $newDestination -PathType Container)) { [void](New-Item $newDestination -Type Directory) }
            elseif (-not (Test-Path -LiteralPath $newBackupPath -PathType Container)) { [void](New-Item $newBackupPath -Type Directory) }
			
            CopyAndBackup $newPath $newDestination $root $newBackupPath
			
			
        }
        else {
            $destItem	=	Join-Path $destination	$Name
            if ((Test-Path -LiteralPath $destItem -PathType Leaf)) {
                Copy-Item -LiteralPath $destItem $backupPath -Force
                Move-Item -LiteralPath $_ $destItem -Force
            }
            else {
                Copy-Item -LiteralPath $_ $destItem -Force
                $workingDir = Get-Location
                Set-Location $root
                #(Get-Item -LiteralPath $destItem | Resolve-Path -Relative)
                Set-Location $workingDir
            }
			

        }
		
    }
}
function Remove-EmptyFoldersSubroutine {
    <#
    .SYNOPSIS
    Recursely removes all empty folders in a given path.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path
    )
    Get-FileSystemEntries $Path -Directory | & { Process {
            Remove-EmptyFoldersSubroutine -Path $_
        } }   
    if ($null -eq (Get-FileSystemEntries $Path | Select-Object -First 1)) {
        Write-Host "Removing empty folder at path `"$Path`"."
        Remove-Item -Force -LiteralPath $Path
    }
}
function Remove-EmptyFolders {
    <#
    .SYNOPSIS
    Recursely removes all empty folders in a given path.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)] [string] $Path
    )
    Process {
        Get-FileSystemEntries $Path -Directory | & { Process {
                Remove-EmptyFoldersSubroutine -Path $_
            } }
    }
}

function Get-7zip {
    if ((Test-Path (Join-Path $PSScriptRoot '7z.exe') -PathType Leaf) -and (Test-Path (Join-Path $PSScriptRoot '7z.dll') -PathType Leaf)) {
        return (Join-Path $PSScriptRoot '7z.exe')
    }
    if ((Test-Path (Join-Path $PSScriptRoot '7za.exe') -PathType Leaf)) {
        return (Join-Path $PSScriptRoot '7za.exe')
    }
    if (Test-Path 'Registry::HKEY_CURRENT_USER\Software\7-Zip') {
        $path = (Get-ItemProperty 'Registry::HKEY_CURRENT_USER\Software\7-Zip').Path64
        if (!$path) {
            $path = (Get-ItemProperty 'Registry::HKEY_CURRENT_USER\Software\7-Zip').Path 
        }
        if ((Test-Path (Join-Path $path '7z.exe') -PathType Leaf) -and (Test-Path (Join-Path $path '7z.dll') -PathType Leaf)) {
            return (Join-Path $path '7z.exe')
        }
    }  
    if (Test-Path 'Registry::HKEY_LOCAL_MACHINE\Software\7-Zip') {
        $path = (Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\Software\7-Zip').Path64
        if (!$path) {
            $path = (Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\Software\7-Zip').Path 
        }
        if ((Test-Path (Join-Path $path '7z.exe') -PathType Leaf) -and (Test-Path (Join-Path $path '7z.dll') -PathType Leaf)) {
            return (Join-Path $path '7z.exe')
        }
    }
}
function Expand-7z {
    [CmdletBinding(PositionalBinding = $false)]
    param (
        [Parameter(Mandatory, ParameterSetName = 'simple', Position = 1)]   [string] $DestinationPath,
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'simple', Position = 0)] [Alias('ArchivePath')]  [string[]] $Path
    )
    if (![System.IO.Path]::IsPathRooted($DestinationPath)) {
        $DestinationPath = Join-Path (Get-Location) $DestinationPath
    }
    $7zip = Get-7zip
    if (!$7zip) {
        Write-Error 'No 7zip found!'
        break
    }
    if (!$root) {
        $root = '.'
    }
    $Process = Start-Process -PassThru -Wait -WorkingDirectory $root -FilePath $7zip -ArgumentList @('x', "`"$Path`"", '-y', "-o`"$DestinationPath`"")
            
    if ($Process.ExitCode -ne 0) {
        Write-Progress @ProgressParameters -Completed
        Write-Error "7zip aborted with an Exit-Code of $($Process.ExitCode)."
        return
    }


}
function Compress-7z {
    [CmdletBinding(PositionalBinding = $false, DefaultParameterSetName = 'simple')]
    param (
        [Parameter(Mandatory, ParameterSetName = 'simple', Position = 0)][Parameter(Mandatory, ParameterSetName = 'advanced', Position = 0)] [Alias('ArchivePath')]   [string] $DestinationPath,
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'simple', Position = 1)][Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'advanced', Position = 1)]   [string[]] $Path,
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'advanced')][AllowEmptyString()]   [string] $Type,
        [Parameter(ParameterSetName = 'simple')][Parameter(ParameterSetName = 'advanced')]   [string] $Root,
        [Parameter(ParameterSetName = 'simple')][Parameter(ParameterSetName = 'advanced')]   [switch] $NonSolid,
        [Parameter(ParameterSetName = 'simple')][Parameter(ParameterSetName = 'advanced')]   [string] $ProgressAction
    ) 

    begin {
        if ($ProgressAction) {
            $prevProgress = $ProgressPreference
            $ProgressPreference = $ProgressAction 
        }
        if (![System.IO.Path]::IsPathRooted($DestinationPath)) {
            $DestinationPath = Join-Path (Get-Location) $DestinationPath
        }
        $7zip = Get-7zip
        if (!$7zip) {
            Write-Error 'No 7zip found!'
            break
        }
        if (!$root) {
            $root = '.'
        }
        if ($nonSolid) {
            $solid = 'off'
        }
        else {
            $solid = 'on'
        }
        $list = [System.Collections.ArrayList]::new()
    }
    process {
        [void]$List.add((
                [PSCustomObject]@{
                    Path = $path
                    Type = $type
                }
            ))
    }
    end {
        $list = $list | Group-Object Type
        $i = 0
        $ProgressParameters = @{
            Activity        = 'Compressing'
            Status          = "$i / $($List.count)"
            PercentComplete = ($i * 100 / $List.count)
        }
        Write-Progress @ProgressParameters

        foreach ($fileType in $List) {
            switch ($filetype.name) {
                'CD-Audio' { $options = '-mf=Delta:4 -m0=LZMA:x9:mt2:d1g:lc1:lp2'; break } 
                'Text' { $options = '-m0=PPmD:x9:o32:mem1g'; break }
                'Binary' { $options = '-m0=LZMA:mt2:x9:d1g'; break }
                'Fast' { $options = '-m0=LZMA:mt2:x3'; break }
                Default { $options = '-m0=LZMA:mt2:x9:d1g'; break }  
            }
            $files = "`"$($fileType.Group.path -join '" "')`""
            $Process = Start-Process -PassThru -Wait -WorkingDirectory $root -FilePath $7zip -ArgumentList @('u', '-r0', '-mqs', "-ms=$solid", $options, "`"$DestinationPath`"", $files)
            
            if ($Process.ExitCode -ne 0) {
                Write-Progress @ProgressParameters -Completed
                Write-Error "7zip aborted with an Exit-Code of $($Process.ExitCode)."
                return
            }

            $i++
            $ProgressParameters = @{
                Activity        = 'Compressing'
                Status          = "$i / $($List.count)"
                PercentComplete = ($i * 100 / $List.count)
            }
            Write-Progress @ProgressParameters
        }
        Write-Progress @ProgressParameters -Completed
        if ($prevProgress) {
            $prevProgress = $ProgressPreference
        }
    }
}

function Invoke-GitHubRestRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)] [string] $Uri,
        [Parameter()] [GitHub_RateLimits] $rates
    )
    class GitHub_RateLimits {
        [int]$remaining = 60
        [DateTimeOffset]$reset = [DateTimeOffset]::('now')
        [DateTimeOffset]$retry = [DateTimeOffset]::('now')
        [DateTimeOffset]$next = [DateTimeOffset]::('now')
    }
    Update-TypeData -Force -TypeName 'GitHub_RateLimits' -MemberName 'next' -MemberType ScriptProperty -Value {
        if ($this.remaining -gt 0) {
            $next = $this.retry
        }
        else {
            $next = (@($this.reset, $this.retry) | Measure-Object -Maximum).Maximum
        }
        return [DateTimeOffset]$next
    } -SecondValue {
        # Allow trying to set "next", but ignore it.
    }

    $prevProg = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'
    
    if (!$rates) {
        $rates = [GitHub_RateLimits]::new()
    }

    $now = [DateTimeOffset]::('now')
    if ($rates.remaining -le 0 -and $rates.next -gt $now) {
        $nextTimespan = '{0:dd}d:{0:hh}h:{0:mm}m:{0:ss}s' -f ($rates.next - $now)
        $ErrorString = "GitHub API rate limit exceeded. Try again in $nextTimespan ($($rates.next.ToLocalTime().toString()))."
        Write-Error $ErrorString
        $ProgressPreference = $prevProg
        return
    }
    elseif ($rates.next -gt $now) {
        $nextTimespan = '{0:dd}d:{0:hh}h:{0:mm}m:{0:ss}s' -f ($rates.next - $now)
        $ErrorString = "GitHub API requested a wait. Try again in $nextTimespan ($($rates.next.ToLocalTime().toString()))."
        Write-Error $ErrorString
        $ProgressPreference = $prevProg
        return
    }

    try {
        $request = Invoke-WebRequest -Uri $Uri
        if ($request.Headers.ContainsKey('x-ratelimit-remaining')) {
            $rates.remaining = $request.Headers['x-ratelimit-remaining']
        }
        if ($request.Headers.ContainsKey('x-ratelimit-reset')) {
            $rates.reset = ([datetimeoffset] '1970-01-01Z').AddSeconds($request.Headers['x-ratelimit-reset'])
        }
        if ($request.Headers.ContainsKey('retry-after')) {
            $rates.retry = $now.AddSeconds($request.Headers['retry-after'])
        }
        else {
            $rates.retry = ($now)
        }
        $RequestError = 0
    }
    catch [System.Net.WebException] {
        if ($_.Exception.Response.Headers['x-ratelimit-remaining']) {
            $rates.remaining = $_.Exception.Response.Headers['x-ratelimit-remaining']
        }
        if ($_.Exception.Response.Headers['x-ratelimit-reset']) {
            $rates.reset = ([datetimeoffset] '1970-01-01Z').AddSeconds($_.Exception.Response.Headers['x-ratelimit-reset'])
        }
        if ($_.Exception.Response.Headers['retry-after']) {
            $rates.retry = $now.AddSeconds($_.Exception.Response.Headers['retry-after'])
        }
        else {
            $rates.retry = $now.AddSeconds(60)
        }
        if ($_.Exception.Response.StatusCode.value__) {
            $RequestError = $_.Exception.Response.StatusCode.value__
            if (($RequestError -eq 403 -or $RequestError -eq 429) -and $rates.remaining -le 0) {
                $rates.remaining--
            }
        }
        else { $RequestError = $_ }
    }

    if ($RequestError) {
        if ($rates.remaining -lt 0) {
            $nextTimespan = '{0:dd}d:{0:hh}h:{0:mm}m:{0:ss}s' -f ($rates.next - [DateTimeOffset]::('now'))
            $ErrorString = "GitHub API rate limit exceeded. Try again in $nextTimespan ($($rates.next.ToLocalTime().toString()))."
            Write-Error $ErrorString
            $ProgressPreference = $prevProg
            return
        }
        else {
            Write-Error $RequestError
            $ProgressPreference = $prevProg
            return
        }
    }
    $ProgressPreference = $prevProg
    return [PSCustomObject]@{
        Response = $request.Content | ConvertFrom-Json
        GH_Rates = $rates
    }
}
#endregion Functions

#region Script
class UpdateConfig {
    [DateTimeOffset]$LastCheck = [DateTimeOffset]::('now').AddHours(-1)
}
class GameEntry {
    [String]$Name
    [System.IO.FileInfo]$Exe
    [String]$Owner
    [String]$Repo
    [String]$Nightly
    [System.IO.FileInfo]$Config
    [System.IO.FileInfo[]]$Save
    [System.IO.FileInfo[]]$AssetArchives
}
class UpdateDatabase {
    [GameEntry[]]$Games
    [PSCustomObject]$GH_Rates
}

$exitCode = 0
$UpdateConfig = [UpdateConfig]::new()

$DatabasePath = Join-Path $PSScriptRoot '.\UpdateDatabase.json'

if (Test-Path -LiteralPath '.\HM-Update\config.json' -PathType Leaf) {
    [UpdateConfig]$UpdateConfig = Get-Content '.\HM-Update\config.json' | ConvertFrom-Json
}
if (!(Test-Path 'HM-Update' -PathType Container)) {
    [void](New-Item -ItemType Directory 'HM-Update')
}
$UpdateDatabase = [UpdateDatabase]::new()
if (Test-Path -LiteralPath $DataBasePath -PathType Leaf) {
    [UpdateDatabase]$UpdateDatabase = Get-Content $DataBasePath | ConvertFrom-Json
}

if ($forceGame) {
    if ($UpdateDatabase.Games.Name -contains $forceGame) {
        $GameInfo = [PSCustomObject]@{
            ProductName       = $forceGame
            ProductVersion    = '0.0.0'
            ProductVersionRaw = [version]::new()
        }
    }
    else {
        $text = "`"$forceGame`" is not supported. Supported games: '$($UpdateDatabase.Games.Name -join "', '")'"
        Throw $text
    }
   
}
else {
    foreach ($exe in $UpdateDatabase.Games.Exe) {
        if ($exe.Exists) {
            $GameInfo = $exe.VersionInfo
            break;
        }  
    }
    if (!$GameInfo) {
        Write-Error "Could not find a compatible game! Try -forceGame <'$($UpdateDatabase.Games.Name -join "' | '")'>"
        Exit 1
    } 
}

if (!$GameInfo.ProductName -or !$GameInfo.ProductVersionRaw ) {
    Write-Error 'Could not read version data from game. Try -forceUpdate'
    Write-Host
    Exit 1
}

if (Test-Path -LiteralPath 'HM-Update\config.json' -PathType Leaf) {
    $UpdateConfig = Get-Content 'HM-Update\config.json' | ConvertFrom-Json
}
else {
    $UpdateConfig | ConvertTo-Json | Out-File 'HM-Update\config.json'
}

$UpdateContext = $UpdateDatabase.Games | Where-Object 'Name' -EQ $GameInfo.ProductName

Write-Host "Game:           $($GameInfo.ProductName)"
Write-Host "Local Version:  $($GameInfo.ProductVersion)"

$RemoteVersion = $null
$now = [DateTimeOffset]::('now')

if ($Version) {
    if ($Version -eq 'Nightly') {
        $RemoteVersion = 'Nightly'
        $Download = [PSCustomObject]@{
            URL = $UpdateContext.Nightly
        }
    }
    else {
        $Request = Invoke-GitHubRestRequest -Uri ('https://api.github.com/repos/' + $UpdateContext.Owner + '/' + $UpdateContext.Repo + '/releases/tags/' + $Version) -Rates $UpdateDatabase.GH_Rates
    }
}
elseif (($now -gt $UpdateConfig.LastCheck.AddHours(1)) -or $forceUpdate -or $forceGame -or $SkipCheckTimer) {
    $Request = Invoke-GitHubRestRequest -Uri ('https://api.github.com/repos/' + $UpdateContext.Owner + '/' + $UpdateContext.Repo + '/releases/latest') -Rates $UpdateDatabase.GH_Rates
}

$ReleaseInfo = $Request.Response
$UpdateDatabase.GH_Rates = $Request.GH_Rates
$UpdateConfig.LastCheck = $now
$UpdateConfig | ConvertTo-Json | Out-File 'HM-Update\config.json'
$UpdateDatabase | ConvertTo-Json | Out-File $DatabasePath

if (!$ReleaseInfo -and (($now -gt $UpdateConfig.LastCheck.AddHours(1)) -or $forceUpdate -or $forceGame -or $SkipCheckTimer)) {
    Write-Error 'Could not get required information from GitHub. (Wrong release tag?)'
    Exit 1
}
    
# Try to convert into a Version number
if ($RemoteVersion -ne 'Nightly') {
    $RemoteVersion = $ReleaseInfo.tag_name
    $Download = $ReleaseInfo.Assets | & { Process {
            if ($_.name -match 'Win64' -or $_.name -match 'Windows') {
                [PSCustomObject]@{
                    URL = $_.browser_download_url
                }
            }
    
        } }
    
    $RemoteVersion = [version]($RemoteVersion -replace '^v ?', '')
}

if ($RemoteVersion) {
    Write-Host "Remote Version: $($RemoteVersion)"
}
else {
    $RemoteVersion = $GameInfo.ProductVersionRaw
    Write-Host 'Update check skipped. Last check was less than an hour ago. Use -SkipCheckTimer to try an update anyway.' -ForegroundColor 'yellow'
}
Write-Host ''

#region Rando saves
$save = $null
if ($RemoteVersion -gt $GameInfo.ProductVersionRaw -or $forceUpdate -or $forceGame -or $Version) {
    Write-Host 'There is a newer version available!' -ForegroundColor 'Blue'
    Write-Host "Download URL: $($Download.URL)"

    $IsRando = & { for ($i = 1; $i -le 3; $i++) {
            # SoH
            if (Test-Path -Path ".\Save\file$i.sav" -PathType 'Leaf') { 
                $save = Get-Content ".\Save\file$i.sav" | ConvertFrom-Json
                if ($save.sections.base.data.n64ddFlag) {
                    if ($save.sections.sohStats.data.gameComplete) {
                        Write-Output ([PSCustomObject]@{
                                Name      = ".\Save\file$i.sav"
                                Completed = $true
                            })
                    }
                    else {
                        Write-Output ([PSCustomObject]@{
                                Name      = ".\Save\file$i.sav"
                                Completed = $false
                            })
                    }
                
                }
            }
            # 2S2H
            elseif (Test-Path -Path ".\Save\file$i.json" -PathType 'Leaf') {
                $save = Get-Content ".\Save\file$i.json" | ConvertFrom-Json
                if ($save.owlSave) {
                    $saveType = 'owlSave'
                }
                else {
                    $saveType = 'newCycleSave'
                }
                if ((($save.$saveType.save.shipSaveInfo.saveType -eq 1))) {
                    if ($save.$saveType.save.shipSaveInfo.fileCompletedAt -eq 0) {
                        Write-Output ([PSCustomObject]@{
                                Name      = ".\Save\file$i.json"
                                Completed = $true
                            })
                    }
                    else {
                        Write-Output ([PSCustomObject]@{
                                Name      = ".\Save\file$i.json"
                                Completed = $false
                            })
                    }
                }
            }
        }
    }
    #endregion Rando saves
    Write-Host 'Updating overwrites your current version, but there will be a backup (including saves and settings).' -ForegroundColor 'DarkYellow'
    Write-Host 'Depending on the update you might need to regenerate your OTR/O2R!' -ForegroundColor 'DarkYellow'
    if ($IsRando) {
        Write-Host 'WARNING: Randomizer save detected:' -ForegroundColor 'yellow'
        Write-Host (($IsRando | Format-Table | Out-String).Trim())
        Write-Host 'Updating will most likely break your Randomizer save!' -ForegroundColor 'yellow'
        Write-Host 'In some cases the game will crash at startup until those saves are deleted!' -ForegroundColor 'red'
    }
    Write-Host 'Do you want to download and install? (y/n)' -NoNewline
    do {
        $answer = Read-Host ' '
    } until ($answer -eq 'y' -or $answer -eq 'n')

    if ($answer -eq 'y') {
        if ($PSBoundParameters.Key -notcontains 'RegenAssetArchive') {
            Write-Host 'Delete asset archive (otr/o2r) after the patch to regenerate? (y/n)' -NoNewline
            do {
                $answer = Read-Host ' '
            } until ($answer -eq 'y' -or $answer -eq 'n')
            if ($answer -eq 'y') {
                $RegenAssetArchive = $true
            }
            else {
                $RegenAssetArchive = $false
            }
        }
        if ($IsRando -and $PSBoundParameters.Key -notcontains 'DeleteRando') {
            Write-Host 'Delete all Randomizer saves after updating? (y/n)' -NoNewline
            do {
                $answer = Read-Host ' '
            } until ($answer -eq 'y' -or $answer -eq 'n')
            if ($answer -eq 'y') {
                $DeleteRando = $true
            }
            else {
                $DeleteRando = $false
            }
        }

        Write-Host "Downloading `"$($Download.URL)`"... " -NoNewline
        $prevProg = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        try {
            Invoke-WebRequest $Download.URL -OutFile 'HM-Update\update.zip'
            Write-Host 'Done!' -ForegroundColor 'Green'
        }
        catch {
            Write-Host 'Error!' -ForegroundColor 'Red'
            Write-Error "Could not download $($Download.URL)"
            Exit 1
        }


        if (Test-Path 'HM-Update\update.zip' -PathType Leaf) {
            if (Test-Path 'HM-Update\update' -PathType Container) {
                Write-Host 'Removing old update directory.'
                Remove-Item -Recurse -Force 'HM-Update\update\'
            }
            Write-Host "Unpacking `".\HM-Update\update.zip`"... " -NoNewline
            try {
                if (Get-7zip) {
                    Expand-7z -ArchivePath 'HM-Update\update.zip' -DestinationPath 'HM-Update\update\' -ErrorAction Stop
                }
                else {
                    Expand-Archive -LiteralPath 'HM-Update\update.zip' -DestinationPath 'HM-Update\update\' -ErrorAction Stop
                }
                
                Write-Host 'Done!' -ForegroundColor 'Green' 
            }
            catch {
                if (Test-Path 'HM-Update\update' -PathType Container) {
                    Remove-Item -Recurse -Force 'HM-Update\update\'
                }
                Write-Host 'Error!' -ForegroundColor 'Red'
                Write-Host $_ -ForegroundColor 'Red'
                Exit 1
            }
            if (Test-Path 'HM-Update\UpBackup' -PathType Container) {
                Write-Host 'Removing old UpBackup directory.'
                Remove-Item -Recurse -Force 'HM-Update\UpBackup\'
            }
            Write-Host 'Updating. Do not abort! ... ' -NoNewline
            try {
                [void](New-Item -ItemType Directory 'HM-Update\UpBackup')
                foreach ($save in $UpdateContext.Save) {
                    if ($UpdateContext.Save.Exists) {
                        Copy-Item -LiteralPath $save -Destination 'HM-Update\UpBackup\' -Recurse -ErrorAction Stop
                    }
                }
                if ($UpdateContext.Config.exists) {
                    Copy-Item -LiteralPath $UpdateContext.Config -Destination 'HM-Update\UpBackup\' -ErrorAction Stop
                }
                if (Test-Path -LiteralPath 'Assets' -PathType 'Container') {
                    Move-Item -LiteralPath 'Assets\' -Destination 'HM-Update\UpBackup\' -ErrorAction Stop
                }
                CopyAndBackup -path 'HM-Update\Update\' -destination '.' -root (Get-Location) -backupPath 'HM-Update\UpBackup\' -ErrorAction Stop
                Remove-EmptyFolders 'HM-Update\UpBackup\'
                Write-Host 'Done!' -ForegroundColor 'Green'
            }
            catch {
                Remove-EmptyFolders 'HM-Update\UpBackup\'
                if (Test-Path 'HM-Update\update' -PathType Container) {
                    Remove-Item -Recurse -Force 'HM-Update\update\'
                }
                Write-Host 'Error!' -ForegroundColor 'Red' 
                Write-Error $_
                Write-Host "Someting went wrong. But you can manually restore previous files from `"HM-Update\UpBackup`" until you start another update."
                Exit 2
            }
            $nowString = [DateTimeOffset]::('now').ToLocalTime().toString('yyyyMMdd_HHmmss')
            
            try {
                if (Get-7zip) {
                    $ext = '7z'
                    Write-Host "Compressing `".\HM-Update\Backup_$nowString.$ext`" ... " -NoNewline
                    @(
                        [PSCustomObject]@{
                            # Use fast compression for .pdb or this takes ages.
                            Path = '.\HM-Update\UpBackup\*.pdb'
                            Type = 'Fast'
                        },
                        # Use default (binary) for everything else.
                        '.\HM-Update\UpBackup\*'
                    ) | Compress-7z -ArchivePath "HM-Update\Backup_$nowString.$ext" -ProgressAction 'Continue' -ErrorAction Stop
                    
                }
                else {
                    $ext = 'zip'
                    Write-Host "Compressing `".\HM-Update\Backup_$nowString.$ext`" ... " -NoNewline
                    Compress-Archive '.\HM-Update\UpBackup\*' "HM-Update\Backup_$nowString.$ext" -ErrorAction Stop
                }
                Remove-Item -Recurse -Force 'HM-Update\UpBackup\'
                Write-Host 'Done!' -ForegroundColor 'Green'
            }
            catch {
                Write-Host 'Error!' -ForegroundColor 'Red' 
                Write-Error $_
                Write-Host 'Someting went wrong while packing a backup. But you can still restore previous files from `".\HM-Update\UpBackup`" until you start another update.'
                $exitCode = 3
            }
            if ($IsRando -and $DeleteRando) {
                Write-Host 'Deleting Randomizer saves ... ' -NoNewline
                $Script:DelError = $false
                $IsRando | & { Process {
                        try {
                            Remove-Item -LiteralPath $_.Name -Force -ErrorAction Stop
                        }
                        catch {
                            $Script:DelError = $true
                        }  
                    } }
                Write-Host 'Done!' -ForegroundColor 'Green'
                if ($DelError) {
                    Write-Host 'Error!' -ForegroundColor 'Red'
                    Write-Host 'Someting went wrong while deleting Randomizer saves.'
                }
            }
            if ($RegenAssetArchive) {
                Write-Host 'Deleting asset archives ... ' -NoNewline
                $Script:DelError = $false
                $UpdateContext.AssetArchives | & { Process {
                        try {
                            Remove-Item -LiteralPath $_.Name -Force -ErrorAction Stop
                        }
                        catch {
                            $Script:DelError = $true
                        }  
                    } }
                Write-Host 'Done!' -ForegroundColor 'Green'
                if ($DelError) {
                    Write-Host 'Error!' -ForegroundColor 'Red'
                    Write-Host 'Someting went wrong while deleting asset archives.'
                }
            }

            if (Test-Path 'HM-Update\update' -PathType Container) {
                Remove-Item -Recurse -Force 'HM-Update\update\'
            }
            if (Test-Path 'HM-Update\update.zip' -PathType Leaf) {
                Remove-Item -Force '.\HM-Update\update.zip'
            }
            Write-Host "Update finnished. If something doesn't work like expected, you can manually restore previous files from `".\HM-Update\Backup_$nowString.$ext`"."
        }
        $ProgressPreference = $prevProg
        if ($exitCode -ne 0) {
            Exit $exitCode
        }
    }
}
else {
    Write-Host 'Your game looks up to date.' -ForegroundColor 'green'
}
if ($autostart -and $GameInfo) {
    Start-Process $GameInfo.FileName
}
#endregion Script
