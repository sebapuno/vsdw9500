@Echo Off
    Echo .  Chequeando archivo de Log...
    If Exist C:\Data\Log\Log.dat Goto respalda
    goto fincheq
:Respalda
    Echo .  Transmitiendo Log.Dat a %SRM_NS%...
    Copy C:\Data\Log\Log.dat C:\Data\Log\%SRM_NO%.vsd
    If Exist C:\Data\Log\%SRM_NO%.vsd Call C:\Bin\S\SendLog.Bat C:\Data\Log\%SRM_NO%.vsd %SRM_NO%.vsd
    If Exist C:\Data\Meson\Ftp_err.dat Goto fin
    erase c:\data\log\%SRM_NO%.vsd > Nul
    cls
    Echo .  Transmisi¢n exitosa de Log...
    goto salida
:fin
    erase c:\data\log\%SRM_NO%.vsd > Nul
    echo .  Fall¢ transmision de Log...
    goto salida
:finresp
    erase c:\data\log\log.a* > Nul
    echo  .  Fall¢ respaldo de Log...
    goto salida
:fincheq
    echo .  No Existe Archivo de Log
:salida

