param(
    [string]$TestFile = '.'
)
$env:PYTHONPATH='H:\LibreOffice-ExcelLike\src;H:\LibreOffice-ExcelLike\src\stubs'
& 'C:\Program Files\LibreOffice\program\python' -m pytest $TestFile
