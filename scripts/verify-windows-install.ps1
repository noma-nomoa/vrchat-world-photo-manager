param(
    [Parameter(Mandatory)][ValidateSet('fresh', 'setup-overwrite', 'squirrel-update')][string]$Scenario,
    [Parameter(Mandatory)][string]$CandidateDirectory,
    [Parameter(Mandatory)][string]$BaselineDirectory,
    [Parameter(Mandatory)][string]$OutputDirectory
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
# These operations intentionally exercise Windows registration and shortcuts.
# Never run them on a developer's ordinary Windows account.
if ($env:GITHUB_ACTIONS -ne 'true' -or $env:RUNNER_ENVIRONMENT -ne 'github-hosted' -or $env:OS -ne 'Windows_NT') {
    throw 'Installation verification runs only in a disposable GitHub-hosted Windows job.'
}
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
if ($repoRoot -ne [IO.Path]::GetFullPath($env:GITHUB_WORKSPACE)) { throw 'Unexpected checkout location.' }
foreach ($variable in @('WORLDSHOT_HOME_ROOT', 'WORLDSHOT_APPDATA_ROOT', 'WORLDSHOT_LOCALAPPDATA_ROOT')) {
    if ([Environment]::GetEnvironmentVariable($variable)) { throw "Unexpected profile override: $variable" }
}
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
if (-not $OutputDirectory.StartsWith((Join-Path $repoRoot 'out') + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Output must be inside this checkout/out.'
}
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$installRoot = Join-Path $env:LOCALAPPDATA 'worldshot_log'
$registryPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\worldshot_log'
$report = [ordered]@{ ok = $false; scenario = $Scenario; startedAt = [DateTime]::UtcNow.ToString('o'); installRoot = $installRoot; steps = @(); limits = @('Windows hosted runner only; no claim of all Windows/GPU compatibility.', 'Squirrel update applies the real package; the UI quitAndInstall/restart interaction and public update feed are separate checks.') }

function Save-Report {
    $report | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'result.json') -Encoding utf8
}
function Invoke-CheckedProcess([string]$Name, [string]$Executable, [string[]]$Arguments, [int]$TimeoutSeconds = 180) {
    $stdout = Join-Path $OutputDirectory "$Name.stdout.log"
    $stderr = Join-Path $OutputDirectory "$Name.stderr.log"
    $process = Start-Process -FilePath $Executable -ArgumentList $Arguments -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr
    if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
        & taskkill.exe /PID $process.Id /T /F | Out-Null
        throw "$Name timed out; see $stdout and $stderr"
    }
    $process.WaitForExit()
    $report.steps += @{ name = $Name; exitCode = $process.ExitCode; stdout = $stdout; stderr = $stderr }
    Save-Report
    if ($process.ExitCode -ne 0) { throw "$Name exited $($process.ExitCode); see $stdout and $stderr" }
}
function Verify-Feed([string]$Directory, [string]$Version) {
    $lines = @((Get-Content -LiteralPath (Join-Path $Directory 'RELEASES')) | Where-Object { $_.Trim() })
    if ($lines.Count -ne 1 -or $lines[0] -notmatch '^([0-9a-fA-F]{40})\s+(\S+)\s+(\d+)$') { throw "Malformed RELEASES: $Directory" }
    $sha1 = $Matches[1]; $name = $Matches[2]; $bytes = [long]$Matches[3]
    if ($name -ne "worldshot_log-$Version-full.nupkg") { throw "Unexpected package: $name" }
    $package = Join-Path $Directory $name
    if ((Get-Item -LiteralPath $package).Length -ne $bytes -or (Get-FileHash -LiteralPath $package -Algorithm SHA1).Hash -ne $sha1) { throw "RELEASES integrity mismatch: $package" }
    return $package
}
function Verify-Installation([string]$Version, [string]$Package) {
    $appRoot = Join-Path $installRoot "app-$Version"
    $application = Join-Path $appRoot 'WorldShotLog.exe'
    if (-not (Test-Path -LiteralPath $application)) { throw "Installed EXE missing: $application" }
    $exeVersion = (Get-Item -LiteralPath $application).VersionInfo.ProductVersion
    if ($exeVersion -ne $Version) { throw "Installed version $exeVersion differs from $Version" }
    $registration = Get-ItemProperty -LiteralPath $registryPath
    if ($registration.DisplayVersion -ne $Version -or $registration.UninstallString -notlike "*$installRoot*Update.exe*--uninstall*") { throw 'Uninstall registration is missing or stale.' }
    $shell = New-Object -ComObject WScript.Shell
    $shortcuts = @()
    foreach ($folder in @([Environment]::GetFolderPath('DesktopDirectory'), [Environment]::GetFolderPath('StartMenu'))) {
        foreach ($file in @(Get-ChildItem -LiteralPath $folder -Recurse -File -Filter '*WorldShot*.lnk' -ErrorAction SilentlyContinue)) {
            $shortcut = $shell.CreateShortcut($file.FullName)
            # Squirrel 2 uses the stable root execution stub, which selects the
            # installed version; it does not pass --processStart to Update.exe.
            if ($shortcut.TargetPath -ne (Join-Path $installRoot 'WorldShotLog.exe') -or $shortcut.Arguments) { throw "Unexpected shortcut target: $($file.FullName)" }
            $shortcuts += @{ path = $file.FullName; target = $shortcut.TargetPath; arguments = $shortcut.Arguments }
        }
    }
    if ($shortcuts.Count -lt 1) { throw 'No installed WorldShot shortcut was created.' }
    $archive = [IO.Compression.ZipFile]::OpenRead($Package)
    $verified = @()
    try {
        foreach ($entry in $archive.Entries) {
            if (-not $entry.FullName.StartsWith('lib/net45/') -or $entry.FullName.EndsWith('/')) { continue }
            $relative = $entry.FullName.Substring('lib/net45/'.Length)
            if ($relative.Contains('..') -or $relative.Contains(':') -or $relative.StartsWith('/')) { throw 'Unsafe NUPKG path.' }
            $target = if ($relative.EndsWith('_ExecutionStub.exe')) { Join-Path $installRoot ($relative.Replace('_ExecutionStub.exe', '.exe')) } else { Join-Path $appRoot $relative }
            $stream = $entry.Open()
            $hash = [Security.Cryptography.SHA256]::Create()
            try { $expected = [Convert]::ToHexString($hash.ComputeHash($stream)) } finally { $stream.Dispose(); $hash.Dispose() }
            if ((Get-Item -LiteralPath $target).Length -ne $entry.Length -or (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash -ne $expected) { throw "Installed payload mismatch: $relative" }
            $verified += @{ file = $relative; sha256 = $expected.ToLowerInvariant(); bytes = $entry.Length }
        }
    } finally { $archive.Dispose() }
    if ($verified.Count -lt 20) { throw 'NUPKG payload verification was incomplete.' }
    $updaterHash = (Get-FileHash -LiteralPath (Join-Path $installRoot 'Update.exe') -Algorithm SHA256).Hash
    $payloadUpdaterHash = (Get-FileHash -LiteralPath (Join-Path $appRoot 'squirrel.exe') -Algorithm SHA256).Hash
    if ($updaterHash -ne $payloadUpdaterHash) { throw 'Installed updater differs from packaged Squirrel.' }
    return @{ version = $Version; productVersion = $exeVersion; application = $application; registration = @{ DisplayVersion = $registration.DisplayVersion; UninstallString = $registration.UninstallString }; shortcuts = $shortcuts; files = $verified }
}
function Invoke-Runtime([string]$Name, [string]$Version, [string]$Mode) {
    $helper = Join-Path $PSScriptRoot 'lib/verify-installed-runtime.js'
    $exe = Join-Path $installRoot "app-$Version/WorldShotLog.exe"
    Invoke-CheckedProcess $Name (Get-Command node.exe).Source @(('"' + $helper + '"'), ('"' + $exe + '"'), $Version, $Mode, ('"' + $OutputDirectory + '"')) 180
}

try {
    if ((Test-Path -LiteralPath $installRoot) -or (Test-Path -LiteralPath $registryPath)) { throw 'Runner already contains WorldShot; refusing to overwrite it.' }
    $candidatePackage = Verify-Feed $CandidateDirectory '2.5.4'
    if ($Scenario -ne 'fresh') {
        $baselinePackage = Verify-Feed $BaselineDirectory '2.5.3'
        Invoke-CheckedProcess 'install-baseline' (Join-Path $BaselineDirectory 'WorldShotLogSetup.exe') @('--silent')
        $report.baseline = Verify-Installation '2.5.3' $baselinePackage
        Invoke-Runtime 'seed-baseline-data' '2.5.3' 'seed'
    }
    if ($Scenario -eq 'squirrel-update') {
        Invoke-CheckedProcess 'apply-squirrel-update' (Join-Path $installRoot 'Update.exe') @('--update', ('"' + [IO.Path]::GetFullPath($CandidateDirectory) + '"'))
    } else {
        Invoke-CheckedProcess 'install-candidate' (Join-Path $CandidateDirectory 'WorldShotLogSetup.exe') @('--silent')
    }
    # Squirrel's self-update helper completes after its parent exits.
    $deadline = [DateTime]::UtcNow.AddSeconds(20)
    do {
        $ready = (Get-FileHash -LiteralPath (Join-Path $installRoot 'Update.exe') -Algorithm SHA256).Hash -eq (Get-FileHash -LiteralPath (Join-Path $installRoot 'app-2.5.4/squirrel.exe') -Algorithm SHA256).Hash
        if (-not $ready) { Start-Sleep -Milliseconds 200 }
    } until ($ready -or [DateTime]::UtcNow -gt $deadline)
    $report.candidate = Verify-Installation '2.5.4' $candidatePackage
    Invoke-Runtime 'candidate-first-launch' '2.5.4' $(if ($Scenario -eq 'fresh') { 'seed' } else { 'verify' })
    Invoke-Runtime 'candidate-restart' '2.5.4' 'verify'
    Invoke-CheckedProcess 'installed-packaged-smoke' (Get-Command node.exe).Source @(('"' + (Join-Path $PSScriptRoot 'smoke-packaged-app.js') + '"'), ('"--exe=' + (Join-Path $installRoot 'app-2.5.4/WorldShotLog.exe') + '"')) 240
    $report.ok = $true
} catch {
    $report.error = $_.ToString() + "`n" + $_.ScriptStackTrace
    Write-Error -Message $report.error -ErrorAction Continue
} finally {
    $report.finishedAt = [DateTime]::UtcNow.ToString('o')
    Save-Report
    foreach ($source in @($installRoot, (Join-Path $env:LOCALAPPDATA 'SquirrelTemp'))) {
        if (Test-Path -LiteralPath $source) {
            foreach ($log in @(Get-ChildItem -LiteralPath $source -File -Filter '*.log' -ErrorAction SilentlyContinue)) {
                Copy-Item -LiteralPath $log.FullName -Destination (Join-Path $OutputDirectory ($log.Directory.Name + '-' + $log.Name)) -Force
            }
        }
    }
}
if (-not $report.ok) { throw "Windows installation gate failed: $OutputDirectory/result.json" }
