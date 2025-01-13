goto CL
:MinGWx64
REM リソース
windres WheelSheetChanger.rc -o WheelSheetChanger.o
REM DLL
g++ -shared -o WheelSheetChanger.dll WheelSheetChangerDll.cpp -luser32 -lgdi32 -Wl,--out-implib,libWheelSheetChanger.a
REM exe
g++ -o WheelSheetChanger.exe WheelSheetChanger.cpp WheelSheetChanger.o -mwindows -municode -luser32 -lshell32 -lWheelSheetChanger -Wl,-rpath=. -L.

goto end

:CL

REM 環境設定
call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat"
REM リソース
rc WheelSheetChanger.rc
REM DLL
cl /LD /FeWheelSheetChanger.dll WheelSheetChangerDll.cpp User32.lib Gdi32.lib
REM exe
cl /FeWheelSheetChanger.exe WheelSheetChanger.cpp WheelSheetChanger.lib User32.lib Shell32.lib WheelSheetChanger.res

goto end

:end