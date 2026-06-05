@echo off 
echo Instalando MEPL... 
xcopy /Y "scripts_vbs\*.vbs" "C:\Users\usuario\AppData\Roaming\SAP\SAP GUI\Scripts\" 
pip install requests google-auth google-auth-oauthlib google-api-python-client --quiet 
echo. 
echo Instalacion completada. 
pause 
