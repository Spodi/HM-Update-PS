# HM-Update-PS
## DESCRIPTION
PowerShell script to update Harbour Masters ports.
Gets update data and updates the port in the scripts current working directory.

## PARAMETER
### -Path <String\>
Update the port found in this path.
Uses working directory when not given.

### -Autostart
Automatically starts the game after this script finishes. Except on errors.

### -Version <String\>
Force downloading the release with the given GitHub tag. Usually like "9.0.2" or "v1.0.0". Use "Nightly" for the latest Nightly build.
Implies "-ForceUpdate" and "-SkipCheckTimer".

### -AutoUpdate
Assume "yes" for updating without asking.
Also set -RegenAssetArchive or -RegenAssetArchive:$false and -DeleteRando or -DeleteRando:$false for an unattended update.

### -RegenAssetArchive
Assume "yes" for deleting asset archives without asking. Use -RegenAssetArchive:$false to assume "no".
Also set -AutoUpdate and -DeleteRando or -DeleteRando:$false for an unattended update.

### -DeleteRando
Assume "yes" for deleting Randomizer saves without asking. Use -DeleteRando:$false to assume "no".
Also set -AutoUpdate and -RegenAssetArchive or -RegenAssetArchive:$false for an unattended update.
    
### -SkipCheckTimer
Check for updates more often than once per hour.

### -ForceUpdate
Force an update, regardless of your current version.
Implies "-SkipCheckTimer".

### -ForceGame <String\>
Use update data for the selected game and force an update, no matter what is actually detected. Can be used to simply download and extract any of the ports.
Implies "-ForceUpdate" and "-SkipCheckTimer".
> [!CAUTION]
> -forceGame can corrupt your game if you select a different game than the one already installed!
