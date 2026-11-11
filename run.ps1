# Planning Progress - run from a terminal.
#   .\run.ps1                 open the app (syncs with the real Excel file)
#   .\run.ps1 -NoSync         open the app, read-only towards Excel
#   .\run.ps1 sync status     any other uat_tracker command, e.g. summary, review, sync once, -h
param(
    [switch]$NoSync,
    [Parameter(ValueFromRemainingArguments = $true)][string[]]$Rest
)

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$py = Join-Path $env:USERPROFILE ".venvs\uat-tracker\Scripts\python.exe"
if (-not (Test-Path $py)) {
    Write-Host "Python environment not found: $py" -ForegroundColor Red
    Write-Host "Create it once: python -m venv `$env:USERPROFILE\.venvs\uat-tracker ; & $py -m pip install -e `"$here[test]`""
    exit 1
}
$config = Join-Path $here "config.toml"

if ($Rest -and $Rest.Count -gt 0) {
    $cmd = $Rest
}
elseif ($NoSync) {
    $cmd = @("ui", "--no-sync")
}
else {
    $cmd = @("ui")
}

Push-Location $here
try {
    & $py -m uat_tracker --config $config @cmd
    exit $LASTEXITCODE
}
finally {
    Pop-Location
}
