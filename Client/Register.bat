@ECHO OFF
ECHO Installs runtimes
ECHO.
PAUSE
COPY comdlg32.ocx C:\Windows\System32 /Y
REGSVR32 C:\Windows\System32\comdlg32.ocx /S
ECHO comdlg32.ocx registered.
COPY mscomctl.ocx C:\Windows\System32 /Y
REGSVR32 C:\Windows\System32\mscomctl.ocx /S
ECHO mscomctl.ocx registered.
COPY mswinsck.ocx C:\Windows\System32 /Y
REGSVR32 C:\Windows\System32\mswinsck.ocx /S
ECHO mswinsck.ocx registered.
COPY richtx32.ocx C:\Windows\System32 /Y
REGSVR32 C:\Windows\System32\richtx32.ocx /S
ECHO richtx32.ocx registered.
COPY tabctl32.ocx C:\Windows\System32 /Y
REGSVR32 C:\Windows\System32\tabctl32.ocx /S
ECHO tabctl32.ocx registered.
COPY dx8vb.dll C:\Windows\SysWOW /Y
ECHO.
ECHO Done
ECHO.
PAUSE