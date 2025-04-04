[CmdletBinding(DefaultParameterSetName = 'none')]
param (
    # This will force the update, regardless of your current version.
    [Parameter(ParameterSetName = 'force')][switch]$forceUpdate,
    # Use update data for the selected game, no matter what is actually detected. Implies "-forceUpdate". This can corrupt your game!
    [Parameter(ParameterSetName = 'force')][ValidateSet('Ship of Harkinian', '2 Ship 2 Harkinian', 'Starship')][string]$forceGame
)

function CopyAndBackup {
    [CmdletBinding()]
    Param ($path, $destination, $root, $backupPath)
	
    $items = Get-ChildItem $path
    $items | ForEach-Object {
        $newPath =	Join-Path $path			$_.Name
        $newDestination	=	Join-Path $destination	$_.Name
        $newBackupPath	=	Join-Path $backupPath	$_.Name
		
        if ((Test-Path $_.FullName -PathType Container)) {

            if (-not (Test-Path -LiteralPath $newDestination -PathType Container)) { [void](New-Item $newDestination -Type Directory) }
            elseif (-not (Test-Path -LiteralPath $newBackupPath -PathType Container)) { [void](New-Item $newBackupPath -Type Directory) }
			
            CustomCopy $newPath $newDestination $root $newBackupPath
			
			
        }
        else {
            $destItem	=	Join-Path $destination	$_.Name
            if ((Test-Path -LiteralPath $destItem -PathType Leaf)) {
                Copy-Item -LiteralPath $destItem $backupPath -Force
                Copy-Item -LiteralPath $_.FullName $destItem -Force
            }
            else {
                Copy-Item -LiteralPath $_.FullName $destItem -Force
                $workingDir = Get-Location
                Set-Location $root
				(Get-Item -LiteralPath $destItem | Resolve-Path -Relative)
                Set-Location $workingDir
            }
			

        }
		
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
    begin {
        $prevDir = [System.IO.Directory]::GetCurrentDirectory()
        [System.IO.Directory]::SetCurrentDirectory((Get-Location))
    }
    Process {
        foreach ($childDirectory in [System.IO.Directory]::EnumerateDirectories($path)) {
            Remove-EmptyFolders -Path $childDirectory
        }
        $currentChildren = Write-Output ([System.IO.Directory]::EnumerateFileSystemEntries($path))
        if ($null -eq $currentChildren) {
            Write-Host "Removing empty folder at path '${Path}'."
            Remove-Item -Force -LiteralPath $Path
        }
    }
    end {
        [System.IO.Directory]::SetCurrentDirectory($prevDir)
    }
}

function Invoke-GitHubRestRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)] [string] $Uri
    )
    class GitHub_RateLimits {
        [int]$remaining = 60
        [DateTimeOffset]$reset
        [DateTimeOffset]$retry
        [DateTimeOffset]$next
    }
    Update-TypeData -TypeName 'GitHub_RateLimits' -MemberName 'next' -MemberType ScriptProperty -Value {
        if ($this.remaining -gt 0) {
            $next = $this.retry
        }
        else {
            $next = (@($this.reset, $this.retry) | Measure-Object -Maximum).Maximum
        }
        return [DateTimeOffset]$next
    } -Force

    $prevProg = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'
    $rates = [GitHub_RateLimits]::new()

    if (Test-Path -LiteralPath 'GitHub_RateLimits.json' -PathType Leaf) {
        [GitHub_RateLimits]$rates = Get-Content 'GitHub_RateLimits.json' | ConvertFrom-Json
    }
    else {
        $rates | Select-Object * -ExcludeProperty 'next' | ConvertTo-Json | Out-File 'GitHub_RateLimits.json'
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

    $rates | Select-Object * -ExcludeProperty 'next' | ConvertTo-Json | Out-File 'GitHub_RateLimits.json'

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
    return $request.Content | ConvertFrom-Json | Add-Member -NotePropertyName 'GhRemainingRate' -NotePropertyValue $rates.remaining -PassThru
}

if ($forceGame) {
    $GameInfo = [PSCustomObject]@{
        ProductName       = $forceGame
        ProductVersion    = '0.0.0'
        ProductVersionRaw = [version]::new()
    }
   
}
else {
    if (Test-Path -LiteralPath 'soh.exe' -PathType Leaf) {
        $GameInfo = (Get-Item 'soh.exe').VersionInfo
    }
    elseif (Test-Path -LiteralPath '2ship.exe' -PathType Leaf) {
        $GameInfo = (Get-Item '2ship.exe').VersionInfo
    }
    elseif (Test-Path -LiteralPath 'Starship.exe' -PathType Leaf) {
        $GameInfo = (Get-Item 'Starship.exe').VersionInfo
    }
    else {
        Write-Error "Could not find a compatible game! Try -forceGame <'Ship of Harkinian' | '2 Ship 2 Harkinian' | 'Starship'>"
        Exit 1
    }
}
if (!$GameInfo.ProductName -or !$GameInfo.ProductVersionRaw ) {
    Write-Error 'Could not read version data from game. Try -forceUpdate'
    Write-Host
    Exit 1
}


Write-Host "Game:           $($GameInfo.ProductName)"
Write-Host "Local Version:  $($GameInfo.ProductVersion)"

$lookup = @{
    'Ship of Harkinian'  = @{
        Owner = 'HarbourMasters'
        Repo  = 'Shipwright'
    }
    '2 Ship 2 Harkinian' = @{
        Owner = 'HarbourMasters'
        Repo  = '2ship2harkinian'
    }
    'Starship'           = @{
        Owner = 'HarbourMasters'
        Repo  = 'Starship'
    } 
}

$ReleaseInfo = Invoke-GitHubRestRequest -Uri ('https://api.github.com/repos/' + $lookup[$GameInfo.ProductName].Owner + '/' + $lookup[$GameInfo.ProductName].Repo + '/releases/latest')

if (!$ReleaseInfo) {
    Write-Error 'Could not get required information from GitHub.'
    Exit 1
}

# Try to convert into a Version number

$RemoteVersion = $ReleaseInfo.tag_name
$Download = $ReleaseInfo.Assets | & { Process {
        if ($_.name -match 'Win64' -or $_.name -match 'Windows') {
            [PSCustomObject]@{
                Name = $_.name
                URL  = $_.browser_download_url
            }           
        }

    } }



$RemoteVersion = [version]$RemoteVersion

Write-Host "Remote Version: $($RemoteVersion)"
Write-Host ''
$OngoingRando = $null
$save = $null

if ($RemoteVersion -gt $GameInfo.ProductVersionRaw -or $forceUpdate -or $forceGame) {
    Write-Host 'There is a newer version available!' -ForegroundColor 'Blue'
    Write-Host "Download URL: $($Download.URL)"
    for ($i = 1; $i -le 3; $i++) {
        $save = Get-Content "Save\file$i.sav" | ConvertFrom-Json
        if ($save.sections.base.data.n64ddFlag -and !$save.sections.sohStats.data.gameComplete) {
            $OngoingRando = $true
            break;
        }
    }

    Write-Host 'Updating overwrites your current version!' -ForegroundColor 'yellow'
    if ($OngoingRando) {
        Write-Host 'WARNING: Ongoing Randomizer save detected!' -ForegroundColor 'red'
        Write-Host 'Updating will most likely break your Randomizer save!' -ForegroundColor 'red'
    }
    Write-Host 'Do you want to download and install? (y/n)' -NoNewline
    do {
        $answer = Read-Host ' '
    } until ($answer -eq 'y' -or $answer -eq 'n')

    if ($answer -eq 'y') {
        Write-Host "Downloading `"$($Download.URL)`"... " -NoNewline
        $prevProg = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        try {
            Invoke-WebRequest $Download.URL -OutFile 'update.zip'
            Write-Host 'Done!' -ForegroundColor 'Green'
        }
        catch {
            Write-Host 'Error!' -ForegroundColor 'Red'
            Write-Error "Could not download $($Download.URL)"
            Exit 1
        }


        if (Test-Path 'update.zip' -PathType Leaf) {
            if (Test-Path 'update' -PathType Container) {
                Write-Host 'Removing old update directory.'
                Remove-Item -Recurse -Force 'update\'
            }
            Write-Host "Unpacking `"update.zip`"... " -NoNewline
            try {
                Expand-Archive -LiteralPath 'update.zip' -DestinationPath 'update\' -ErrorAction Stop
                Write-Host 'Done!' -ForegroundColor 'Green' 
            }
            catch {
                if (Test-Path 'update' -PathType Container) {
                    Remove-Item -Recurse -Force 'update\'
                }
                Write-Host 'Error!' -ForegroundColor 'Red'
                Write-Host $_ -ForegroundColor 'Red'
                Exit 1
            }
            if (Test-Path 'UpBackup' -PathType Container) {
                Write-Host 'Removing old UpBackup directory.'
                Remove-Item -Recurse -Force 'UpBackup\'
            }
            Write-Host 'Updating. Do not abort! ... ' -NoNewline
            try {
                CopyAndBackup -path 'Update\' -destination '.' -root (Get-Location) -backupPath 'UpBackup\' -ErrorAction Stop
                Remove-EmptyFolders 'UpBackup\'
                Write-Host 'Done!' -ForegroundColor 'Green'
            }
            catch {
                Remove-EmptyFolders 'UpBackup\'
                if (Test-Path 'update' -PathType Container) {
                    Remove-Item -Recurse -Force 'update\'
                }
                Write-Host 'Error!' -ForegroundColor 'Red' 
                Write-Error $_
                Write-Host 'Someting went wrong. But you can manually restore previous files from "UpBackup" directory until you start another update.'
                Exit 2
            }
            if (Test-Path 'update' -PathType Container) {
                Remove-Item -Recurse -Force 'update\'
            }
            Write-Host "Update finnished. If something doesn't work like expected, you can manually restore previous files from the `"UpBackup`" directory until you start another update."
        }
        $ProgressPreference = $prevProg

    }
}
else {
    Write-Host 'Your game looks up to date.' -ForegroundColor 'green'
}
