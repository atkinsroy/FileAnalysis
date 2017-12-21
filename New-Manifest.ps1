# 
# Create a new Manifest for FileAnalysis
#

New-ModuleManifest -Path '.\FileAnalysis.psd1'  `
                   -Author 'Roy Atkins'                                                   `
                   -CompanyName 'Hewlett Packard Enterprise'                              `
                   -ModuleVersion 0.1.0                                                   `
                   -RootModule '.\FileAnalysis.psm1'                                      `
                   -FunctionsToExport Get-FileListing, Get-FileAnalysis                   `
                   -Description 'Active Directory user maintenance module for McDonalds'  `
                   -Copyright '(c) 2017 Hewlett Packard Enterprise. All rights reserved.' `
                   -PowerShellVersion 3.0