# HM-Update-PS
## DESCRIPTION
PowerShell script to update Harbour Masters ports.
Gets update data and updates the port in the scripts current working directory.

## PARAMETER
### -Autostart
Automatically starts the game after this script finishes. Except on errors.

### -Nightly
Force an update to the latest nightly version.
Implies "-ForceUpdate" and "-SkipCheckTimer".

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
