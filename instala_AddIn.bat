@ECHO Off
SET p=%AppData%\Microsoft\Excel
SET a="XL"
for /D %%x in (%a%*) do if not defined f set "f=%%x"
SET pa=%p%\%f%

CD %pa%

bitsadmin.exe /transfer "ExcelGuru" https://github.com/excelguru/controle-domestica/raw/master/escreve_por_extenso.xlam %pa%\escreve_por_extenso.xlam
bitsadmin.exe /transfer "ExcelGuru" https://github.com/excelguru/controle-domestica/raw/master/escreve_por_extenso.xlam %cd%\escreve_por_extenso.xlam