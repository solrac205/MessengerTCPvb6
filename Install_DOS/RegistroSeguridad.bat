echo off
cls
echo *****************************************************************************
echo Este proceso debe ejecutarse unicamente si en el directorio donde se localiza
echo esta localizado tambien el archivo CRAEncryptTool.dll y el archivo 
echo TCPMessenger.txt, si este archivo no esta localizado en el mismo directorio
echo favor cancele el proceso...
echo *****************************************************************************
cd
echo .
echo Si decea cancelar el proceso precione [Ctrl + C]........
pause
echo *****************************************************************************
echo Registrando DLL de Seguridad TCPMessenger V1.0
echo *****************************************************************************
echo .
REGSVR32 CRAEncryptTool.dll
RENAME TCPMessenger.txt TCPMessenger.ini
echo .
echo *****************************************************************************
echo .
echo DLL ha sido registrado.............................
echo .
pause
echo on