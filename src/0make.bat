@echo off
echo ----------------------------------------------------
echo TiMidity++ generation %1
echo Copyright(C) 2014 NEET co.ltd.
echo ----------------------------------------------------

echo Include 1config.bat
call 1config.bat
echo TiMidity++ Binary : %TIMIDITYBIN%
echo TiMidity++ Patch  : %TIMIDITYPAT%
echo OUTPUT WAVE       : %OUTPUT%

echo ----------------------------------------------------
echo Include project.bat
call project.bat
echo Title             : %TITLE%
echo TiMidity++ Config : %CONFIG%
echo TiMidity++ Option : %OPTION%
echo ----------------------------------------------------

SET MIDI=%MIDI1%
SET MIDIPLAY=%MIDI1PLAY%
call :build

SET MIDI=%MIDI2%
SET MIDIPLAY=%MIDI2PLAY%
call :build

SET MIDI=%MIDI3%
SET MIDIPLAY=%MIDI3PLAY%
call :build

SET MIDI=%MIDI4%
SET MIDIPLAY=%MIDI4PLAY%
call :build

SET MIDI=%MIDI5%
SET MIDIPLAY=%MIDI5PLAY%
call :build

SET MIDI=%MIDI6%
SET MIDIPLAY=%MIDI6PLAY%
call :build

SET MIDI=%MIDI7%
SET MIDIPLAY=%MIDI7PLAY%
call :build

SET MIDI=%MIDI8%
SET MIDIPLAY=%MIDI8PLAY%
call :build

SET MIDI=%MIDI9%
SET MIDIPLAY=%MIDI9PLAY%
call :build

goto exit

:build

IF "%MIDI%"=="" goto :exitb

pushd "%TIMIDITYBIN%"
copy "%INPUT%\%CONFIG%" "%TIMIDITYCFG%"
IF "%MIDIPLAY%"=="1" call :preview

echo %TIMIDITYBIN%\timidity.exe -c %TIMIDITYRCFGPATH%/%CONFIG% -Ow %OPTION% %MIDI%
%TIMIDITYBIN%\timidity.exe -c %TIMIDITYRCFGPATH%/%CONFIG% -Ow %OPTION% %INPUT%\%MIDI%

:exitb
popd
exit /b

:preview
start %TIMIDITYBIN%\timidity.exe -c %TIMIDITYRCFGPATH%/%CONFIG% %OPTION% %INPUT%\%MIDI%
exit /b

:exit
del "%TIMIDITYCFG%\%CONFIG%"
