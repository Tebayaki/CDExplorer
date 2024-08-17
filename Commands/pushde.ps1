$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$targetDir = & "$scriptDir\ExplorerDir.exe"
if ($LASTEXITCODE -eq 0) {
    Push-Location $targetDir
}
else {
    exit $LASTEXITCODE
}