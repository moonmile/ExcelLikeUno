param(
    [string]$scriptfile = '.'
)

$env:PYTHONPATH='..\src\'
& 'C:\Program Files\LibreOffice\program\python' $scriptfile
