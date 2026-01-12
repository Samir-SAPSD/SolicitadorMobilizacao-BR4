@echo off

:: Caminho do diretório onde os arquivos devem ser atualizados
set DESTINO=C:\PyProjects\SolicitadorMobilizacao-BR4

:: Caminho temporário para o clone
set TEMP=%TEMP%\repo_temp

:: URL do repositório remoto
set REPO=https://github.com/Samir-SAPSD/SolicitadorMobilizacao-BR4.git

:: Apaga pasta temporária se já existir
rmdir /s /q "%TEMP%"

:: Clona o repositório na pasta temporária
git clone %REPO% "%TEMP%"

:: Copia os arquivos do repositório clonado para o destino, sobrescrevendo tudo
xcopy "%TEMP%\*" "%DESTINO%\" /E /H /Y /C

:: Limpa a pasta temporária
rmdir /s /q "%TEMP%"

echo Atualização concluída com sucesso!
pause
