param([string]$Python = 'python')

$ErrorActionPreference = 'Stop'
$testsDir = Join-Path (Split-Path -Parent $PSScriptRoot) 'tests'
foreach ($testFile in @(
    'test_dat_reader.py'
    'test_note_parser.py'
    'test_chart_config.py'
    'test_data2excel_static.py'
    'test_doppler.py'
)) {
    if (-not (Test-Path -LiteralPath (Join-Path $testsDir $testFile))) {
        throw "Missing test file: $testFile"
    }
    & $Python -B -m unittest discover -s $testsDir -p $testFile -v
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}
exit 0
