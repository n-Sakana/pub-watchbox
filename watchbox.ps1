# watchbox.ps1 - Unified monitoring and manifest generation tool
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

$combined = (Get-ChildItem "$PSScriptRoot\src\*.cs" | Sort-Object Name |
    ForEach-Object { Get-Content $_ -Raw }) -join "`n"

$pat = '(?m)^\s*using\s+[\w][\w.]*\s*;'
$usings = [regex]::Matches($combined, $pat) |
    ForEach-Object { $_.Value.Trim() } | Sort-Object -Unique
$body = $combined -replace $pat, ''
$source = ($usings -join "`n") + "`n`n" + $body

Add-Type -AssemblyName 'System.Xaml'

$refs = @(
    [System.Windows.Window].Assembly.Location
    [System.Windows.UIElement].Assembly.Location
    [System.Windows.DependencyObject].Assembly.Location
    [System.Xaml.XamlReader].Assembly.Location
    'Microsoft.CSharp'
)

Add-Type -TypeDefinition $source -ReferencedAssemblies $refs
[WatchBox.App]::Run($PSScriptRoot)
