# HM-Update-PS
## DESCRIPTION
PowerShell script to update HaborMasters ports.
Gets update data and updtates the port in the scripts current working directory.

## PARAMETER
### -forceUpdate
This will force an update, regardless of your current version.

### -nightly
This will update to the latest nightly version.\
Implies "-forceUpdate".

### -forceGame <"Ship of Harkinian", "2 Ship 2 Harkinian", "Starship"\>
Use update data for the selected game, no matter what is actually detected. Can be used to simply download and extract any of the ports.\
Implies "-forceUpdate".
> [!CAUTION]
> -forceGame can corrupt your game if you select a different game than the one already installed!
