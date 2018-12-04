# PowerShell-SolarWinds_N-Central
Scripts to help work with https://www.solarwindsmsp.com/products/n-central

If the N-Able agent has issues it can go into a disconnected state (Grey Hourglass Icon), if this persists following a computer reboot then an uninstallation and reinstallation usually gets things going again. I have created two functions to do this from PowerShell as detailed below.
Note: NcentralAssetTool.exe is used to clear unique identifiers on the system (https://support.solarwindsmsp.com/kb/solarwinds_n-central/Agent-install-error-3042-Device-Already-Exists) and PsExec from Sysinternal https://docs.microsoft.com/en-us/sysinternals/downloads/psexec is used to enable PowerShell remoting. 

## Examples

 ```powershell
Uninstall-NAbleAgent -computerName "W10L-12345"
```
![Uninstall](https://github.com/jfrmilner/PowerShell-SolarWinds_N-Central/blob/master/Images/Uninstall-NableAgent.JPG)

 ```powershell
Install-NAbleAgent -computerName "W10L-12345" -customerID 321
```
![Install](https://github.com/jfrmilner/PowerShell-SolarWinds_N-Central/blob/master/Images/Install-NableAgent.JPG)
