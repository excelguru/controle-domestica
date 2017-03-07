@ECHO Off

:: Desenvolvido por ExcelGuru
:: http://excelguru.com.br
:: 
:: Script desenvolvido apenas para instalar a extesão 
:: do Excel automaticamente e carregar junto com o Excel na abertura.
:: 
:: Esta extensão é bem leve e necessária para utilizar a planilha de emissão
:: de recibo de vale transporte.
:: 
:: ########################################################################################

:: SETLOCAL - ensures that I don’t clobber any existing variables after my script exits
:: ENABLEEXTENSIONS - turns on a very helpful feature called command processor extensions
:: as of: http://steve-jansen.github.io/guides/windows-batch-scripting/index.html
SETLOCAL ENABLEEXTENSIONS
SET me=%~n0
SET parent=%~dp0

SET interactive=0
ECHO %CMDCMDLINE% | FIND /I "/c" >NUL 2>&1
IF %ERRORLEVEL% == 0 SET interactive=1

:: Link de onde vai baixar a extensão
SET LINK=https://github.com/excelguru/recibo-vale-transporte/raw/master/escreve_por_extenso.xlam

:: Nome final do arquivo
SET FILE_NAME=escreve_por_extenso.xlam

:: Pasta Geral do Excel do usuário
SET EXCEL_PATH=%AppData%\Microsoft\Excel

:: Primeiras letras a procurar
SET SEARCH=XL

:: Acessa pasta do Excel
CD %excel_path%

:: Pesquisa se existe uma pasta que começa com as letras definidas em SEARCH
for /D %%x in (%search%*) do if not defined FOLDER set "FOLDER=%%x"

:: Configura pasta de extensões
SET EXCEL_START_UP=%excel_path%\%folder%

:: Baixa o a extensão
bitsadmin.exe /transfer "ExcelGuru" %link% %excel_start_up%\%file_name%

:: Verifica se houve algum erro
IF %ERRORLEVEL% NEQ 0 (
    ECHO ------------------------------------------------------- & echo. & echo.
    ECHO Erro ao instalar o AddIn. & echo. & echo.
    ECHO Favor acesse o site abaixo para configurar manualmente o AddIn 'Escreve-por-Extenso' & echo. & echo.
    ECHO Link: http://excelguru.com.br & echo. & echo.
    ECHO ------------------------------------------------------- & echo. & echo.
) ELSE (
    ECHO ------------------------------------------------------- & echo. & echo.
    ECHO AddIn instalado com sucesso! & echo. & echo.
    ECHO Favor fechar e abrir novamente o Excel & echo. & echo.
    ECHO ------------------------------------------------------- & echo. & echo.
)

IF "%interactive%"=="1" PAUSE

ENDLOCAL
ECHO ON
@EXIT /B 0