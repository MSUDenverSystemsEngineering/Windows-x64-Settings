@TITLE "Metropolitan State University of Denver Data Transfer"
@COLOR 2

FOR /F "tokens=*" %%s IN ('dir "%SystemDrive%\Data\%USERNAME%\Desktop" /S/B') DO MOVE "%%s" "%USERPROFILE%\Desktop"
FOR /F "tokens=*" %%s IN ('dir "%SystemDrive%\Data\%USERNAME%\Favorites" /S/B') DO MOVE "%%s" "%USERPROFILE%\Favorites"
FOR /F "tokens=*" %%s IN ('dir "%SystemDrive%\Data\%USERNAME%\Music" /S/B') DO MOVE "%%s" "%USERPROFILE%\Music"
FOR /F "tokens=*" %%s IN ('dir "%SystemDrive%\Data\%USERNAME%\Pictures" /S/B') DO MOVE "%%s" "%USERPROFILE%\Pictures"
FOR /F "tokens=*" %%s IN ('dir "%SystemDrive%\Data\%USERNAME%\Videos" /S/B') DO MOVE "%%s" "%USERPROFILE%\Videos"
FOR /F "tokens=*" %%s IN ('dir "%SystemDrive%\Data\%USERNAME%\Downloads" /S/B') DO MOVE "%%s" "%USERPROFILE%\Downloads"
FOR /F "tokens=*" %%s IN ('dir "%SystemDrive%\Data\%USERNAME%\Documents" /S/B') DO MOVE "%%s" "%USERPROFILE%\Documents"

EXIT /B 0
