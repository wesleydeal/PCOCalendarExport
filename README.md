# PCOCalendarExport.ps1
PowerShell Cmdlets to assist in migration from [Planning Center Online](https://www.planningcenter.com/) Calendar to another service (such as FMX) which insists on a CSV input format.

## Usage
1. Provide your App ID and Secret obtained from [here](https://api.planningcenteronline.com/oauth/applications) near the top of the script in the ```*** CONFIGURATION ***``` section.
2. Load the Cmdlets in Powershell by dot-sourcing: `. .\PCOCalendarExport.ps1`
3. Run `Export-AllPCOData` in the folder of your choice. A timestamped folder tree will be created with Attachments, JSON, and CSV subfolders. The script will first download attachments, then JSON data, then will process the JSON data offline to flatten it into CSV format.

## Acknowledgements
Thanks to [Ronald Bode](https://github.com/iRon7) for the `Flatten-Object` Cmdlet which facilitated the data conversion from JSON to CSV.
