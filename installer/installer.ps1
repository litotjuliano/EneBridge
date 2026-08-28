$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

$issPath = Join-Path $repoRoot "installer\eneBridge.iss"
$csprojPath = Join-Path $repoRoot "src\eneBridge.Wpf\eneBridge.Wpf.csproj"
$content = Get-Content $issPath -Raw

$match = [regex]::Match($content, '#define MyAppVersion "([^"]+)"')
if (-not $match.Success) {
    Write-Error "Could not find MyAppVersion in $issPath"
    Read-Host "Press Enter to exit"
    exit 1
}
$currentVersion = $match.Groups[1].Value

Write-Host "Current installer version: $currentVersion"
$newVersion = Read-Host "Enter new version (press Enter to keep $currentVersion)"

if ([string]::IsNullOrWhiteSpace($newVersion)) {
    $newVersion = $currentVersion
    Write-Host "Keeping version $currentVersion."
} else {
    Write-Host "Updating version to $newVersion..."
    $updatedContent = $content -replace '#define MyAppVersion "[^"]+"', "#define MyAppVersion `"$newVersion`""
    Set-Content -Path $issPath -Value $updatedContent -NoNewline

    # Keep the app's own <Version> (eneBridge.Wpf.csproj) in sync -- this is what the running
    # app reads at startup for its window title, so it must match the installer version.
    $csprojContent = Get-Content $csprojPath -Raw
    $updatedCsproj = $csprojContent -replace '<Version>[^<]+</Version>', "<Version>$newVersion</Version>"
    Set-Content -Path $csprojPath -Value $updatedCsproj -NoNewline

    # Log what changed, in the same step as the bump, so CHANGELOG.md can't drift out of sync
    # with the actual installer version.
    Write-Host ""
    Write-Host "What changed in version $newVersion? (press Enter on an empty line to finish)"
    $changes = @()
    while ($true) {
        $line = Read-Host "  - "
        if ([string]::IsNullOrWhiteSpace($line)) { break }
        $changes += $line
    }
    if ($changes.Count -eq 0) {
        $changes = @("(no description provided)")
    }

    $changelogPath = Join-Path $repoRoot "CHANGELOG.md"
    $changelogContent = Get-Content $changelogPath -Raw
    $entryLines = $changes | ForEach-Object { "- $_" }
    $newEntry = "## [$newVersion] - $(Get-Date -Format 'yyyy-MM-dd')`n" + ($entryLines -join "`n") + "`n`n"
    $updatedChangelog = $changelogContent -replace '(?<=# Changelog\r?\n\r?\n)', $newEntry
    Set-Content -Path $changelogPath -Value $updatedChangelog -NoNewline
}

Write-Host ""
Write-Host "Publishing app (Release)..."
dotnet publish "src\eneBridge.Wpf\eneBridge.Wpf.csproj" -c Release -o "installer\publish"
if ($LASTEXITCODE -ne 0) {
    Write-Error "dotnet publish failed - see errors above."
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Write-Host "Compiling installer with Inno Setup..."

$isccCmd = Get-Command iscc -ErrorAction SilentlyContinue
if ($isccCmd) {
    $isccPath = $isccCmd.Source
} else {
    $isccPath = Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe"
}

if (-not (Test-Path $isccPath)) {
    Write-Error "Could not find ISCC.exe (Inno Setup compiler). Checked PATH and '$isccPath'."
    Read-Host "Press Enter to exit"
    exit 1
}

& $isccPath "installer\eneBridge.iss"
if ($LASTEXITCODE -ne 0) {
    Write-Error "Installer compilation failed - see errors above."
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Write-Host "Done. Installer created at installer\Output\eneBridge-Setup.exe (version $newVersion)"
Read-Host "Press Enter to exit"
