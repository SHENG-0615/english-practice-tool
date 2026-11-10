param([switch]$NoBrowser)
$ErrorActionPreference = 'Stop'
try {
    $syncSource = Join-Path $PSScriptRoot 'local-sync.cs'
    $syncCode = [System.IO.File]::ReadAllText($syncSource, [System.Text.Encoding]::UTF8)
    Add-Type -TypeDefinition $syncCode -ReferencedAssemblies 'System.Web.Extensions' -Language CSharp
    [IeltsLocalSync]::Run($PSScriptRoot, -not $NoBrowser)
} catch {
    $syncError = $_.Exception.ToString()
    [System.IO.File]::WriteAllText((Join-Path $PSScriptRoot 'launch-error.txt'), $syncError, [System.Text.Encoding]::UTF8)
    if ($NoBrowser) { throw }
    $popup = New-Object -ComObject WScript.Shell
    $null = $popup.Popup('The local practice could not start. See launch-error.txt in this folder.', 0, 'English practice', 16)
}
