Disclaimer:
This script is provided free of charge and on an “as is,” “as available” basis, without any warranties or representations of any kind—express or implied—including (but not limited to) merchantability, fitness for a particular purpose, reliability, security, accuracy, or completeness.
By downloading, using, or modifying this script, you acknowledge and agree that:

You run this script entirely at your own risk.
You are solely responsible for reviewing, validating, and testing the script in your environment before deployment.
The author and the author’s employer assume no responsibility or liability for any damages or losses, including but not limited to:

- data loss or corruption
- system damage or malfunction
- security incidents
- downtime or service interruption
- or any direct, indirect, incidental, consequential, special, exemplary, or punitive damages


No support, guarantee, or ongoing maintenance is provided.

By using this script, you agree that neither the author nor the author’s employer shall be held liable under any circumstances. Use at your own risk.

Usage:

# Save as Get-ServerInventory.ps1
# Run PowerShell as Administrator (recommended)

# Current credentials
.\Get-ServerInventory.ps1 -ComputerName @('srv01','srv02','srv03') -ExcelPath 'C:\Temp\Inv\ServerInventory.xlsx'

# Alternate credentials (used for CIM, admin shares, and short-sample remoting)
$cred = Get-Credential
.\Get-ServerInventory.ps1 -ComputerName (Get-Content .\servers.txt) -Credential $cred -OutputFolder 'C:\Temp\Inv'

# Show Excel while writing (debug/visual)
.\Get-ServerInventory.ps1 -ComputerName (Get-Content .\servers.txt) -ExcelVisible