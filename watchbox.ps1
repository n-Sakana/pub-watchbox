# watchbox.ps1 - Unified monitoring and manifest generation tool
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# Load WebView2 assemblies from lib/
$libDir = "$PSScriptRoot\lib"
if (Test-Path "$libDir\Microsoft.Web.WebView2.Core.dll") {
    [System.Reflection.Assembly]::LoadFrom("$libDir\Microsoft.Web.WebView2.Core.dll") | Out-Null
    [System.Reflection.Assembly]::LoadFrom("$libDir\Microsoft.Web.WebView2.Wpf.dll") | Out-Null
}

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
    'System.Drawing'
    'System.IO.Compression'
    'System.IO.Compression.FileSystem'
)

# Add WebView2 references if available
if (Test-Path "$libDir\Microsoft.Web.WebView2.Wpf.dll") {
    $refs += "$libDir\Microsoft.Web.WebView2.Core.dll"
    $refs += "$libDir\Microsoft.Web.WebView2.Wpf.dll"
}

Add-Type -TypeDefinition $source -ReferencedAssemblies $refs
$viewerOnly = $args -contains '--viewer'
[WatchBox.App]::Run($PSScriptRoot, $viewerOnly)
