param([switch]$CheckOnly)
$ErrorActionPreference = 'Stop'

# WPF needs an STA thread to load the XAML, so restart in Windows PowerShell with -STA if needed
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    & powershell -NoProfile -STA -ExecutionPolicy Bypass -File $PSCommandPath @PSBoundParameters
    exit $LASTEXITCODE
}

$root = $PSScriptRoot
$src = Join-Path $root 'src\m365-global-auto-reply.ps1'
$dist = Join-Path $root 'dist'

# Check before building: the script parses and every window's XAML loads with all named controls
$text = [IO.File]::ReadAllText($src)
$errors = $null
[void][Management.Automation.Language.Parser]::ParseInput($text, [ref]$null, [ref]$errors)
if ($errors) { $errors | ForEach-Object { Write-Host "Line $($_.Extent.StartLineNumber): $($_.Message)" -ForegroundColor Red }; throw 'Script has syntax errors' }

Add-Type -AssemblyName PresentationFramework
function Get-Block($name) { [regex]::Match($text, "(?s)\`$$name = @'\r?\n(.*?)\r?\n'@").Groups[1].Value }
$styles = Get-Block 'styles'
foreach ($name in 'mainXaml', 'replyXaml', 'confirmXaml', 'viewXaml') {
    $xaml = (Get-Block $name).Replace('<!--STYLES-->', $styles) -replace '(?<![\w.])(Text|Content|Title|ToolTip)="[^"{][^"]*"', '$1="x"'
    try { $window = [Windows.Markup.XamlReader]::Parse($xaml) }
    catch { throw "$name does not load: $($_.Exception.Message)" }
    if ($name -eq 'mainXaml') {
        $list = [regex]::Match($text, "(?s)foreach \(\`$n in (.*?)\) \{\s*Set-Variable").Groups[1].Value
        $missing = [regex]::Matches($list, "'(\w+)'") | ForEach-Object { $_.Groups[1].Value } | Where-Object { -not $window.FindName($_) }
        if ($missing) { throw "Controls not found in mainXaml: $($missing -join ', ')" }
    }
}
Write-Host 'Check passed' -ForegroundColor Green
if ($CheckOnly) { return }

if (-not (Get-Module -ListAvailable ps2exe)) { Install-Module ps2exe -Scope CurrentUser -Force }
Import-Module ps2exe

# Bake the icon and the font into a copy of the script, so the exe is a single file
$assets = [ordered]@{}
foreach ($name in 'app.ico', 'fa-solid-900.ttf') {
    $assets[$name] = [Convert]::ToBase64String([IO.File]::ReadAllBytes((Join-Path $root "assets\$name")))
}
$block = '$embeddedAssets = @{' + (($assets.Keys | ForEach-Object { " '$_' = '$($assets[$_])'" }) -join ';') + ' }'
$text = [IO.File]::ReadAllText($src).Replace('# EMBEDDED-ASSETS', $block)
$temp = Join-Path ([IO.Path]::GetTempPath()) 'm365-global-auto-reply.ps1'
[IO.File]::WriteAllText($temp, $text, (New-Object Text.UTF8Encoding $true))

# PS2EXE breaks on paths with [ ] (they are read as wildcards), so build in the temp folder and copy afterwards
$tempExe = Join-Path ([IO.Path]::GetTempPath()) 'm365-global-auto-reply.exe'
$tempIco = Join-Path ([IO.Path]::GetTempPath()) 'm365-global-auto-reply.ico'
[IO.File]::Copy((Join-Path $root 'assets\app.ico'), $tempIco, $true)
Invoke-ps2exe -inputFile $temp -outputFile $tempExe `
    -noConsole -STA -x64 -DPIAware -iconFile $tempIco `
    -title 'm365-global-auto-reply' -product 'm365-global-auto-reply' -version '0.1.0.0'
[void][IO.Directory]::CreateDirectory($dist)
[IO.File]::Copy($tempExe, (Join-Path $dist 'm365-global-auto-reply.exe'), $true)
foreach ($f in $temp, $tempExe, $tempIco) { [IO.File]::Delete($f) }