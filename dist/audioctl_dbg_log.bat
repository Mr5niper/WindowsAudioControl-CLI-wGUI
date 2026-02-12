@echo off
:: ==========================================================================
:: Debug Logging while running the GUI (works with UAC elevation)
:: ==========================================================================
:: This batch file launches audioctl.exe with AUDIOCTL_DEBUG=1 so the GUI writes
:: verbose [DBG ...] lines to audioctl_gui.log.
::
:: NOTE ABOUT ADMIN MODE:
:: If audioctl.exe is launched elevated (UAC / Run as Administrator), the process
:: may not inherit environment variables from the non-elevated shell. To make
:: AUDIOCTL_DEBUG=1 reliable, this script starts an elevated cmd.exe that sets
:: the variable and then launches audioctl.exe.
:: ==========================================================================

setlocal

echo Starting audioctl.exe GUI as Administrator with verbose logging...

set "APP=%~dp0audioctl.exe"
set "WD=%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass ^
  -Command "Start-Process -Verb RunAs -FilePath cmd.exe -WorkingDirectory '%WD%' -ArgumentList '/c', 'set AUDIOCTL_DEBUG=1&&\"%APP%\"'"

endlocal

