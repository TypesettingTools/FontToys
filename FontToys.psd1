@{
RootModule = 'FontToys.psm1'
ModuleVersion = '0.0.1'
GUID = '0544c2df-1a5e-4da7-95c8-1583af4b5b9b'
Author = 'line0'
Description = 'A small set of PowerShell modules that enable batch processing of Matroska files with mkvtoolnix.'

PowerShellVersion = '3.0'

NestedModules = @('Import-Font.psm1', 'Export-Font.psm1', 'Format-Font.psm1')
FunctionsToExport = @('Import-Font', 'Format-Font', 'Export-Font')
CmdletsToExport = ''
VariablesToExport = '*'
AliasesToExport = '*'
RequiredModules = @('MkvTools')


# HelpInfo URI of this module
# HelpInfoURI = ''
}