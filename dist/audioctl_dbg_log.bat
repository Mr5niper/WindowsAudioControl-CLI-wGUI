@echo off
:: ==========================================================================
:: Debug Logging while running the GUI
:: ==========================================================================
:: This bat file sets AUDIOCTL_DEBUG=1
:: allowing all debug messages to be logged in audioctl_gui.log
:: ==========================================================================
echo Starting audioctl.exe GUI with verbose logging
set AUDIOCTL_DEBUG=1
audioctl.exe
