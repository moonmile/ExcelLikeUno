param(
    [string]$TestFile = '.'
)

$env:PYTHONPATH='H:\LibreOffice-ExcelLike\src\'
& 'C:\Program Files\LibreOffice\program\python' -m pytest $TestFile
