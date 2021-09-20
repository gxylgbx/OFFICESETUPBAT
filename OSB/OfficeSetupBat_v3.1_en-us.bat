@echo off
setlocal enabledelayedexpansion
%1 mshta vbscript:CreateObject("Shell.Application").ShellExecute("cmd.exe","/c %~s0 ::","","runas",1)(window.close)&&exit
cd /d %~dp0
set aboutver=v3.1
set "singleline=-------------------------------------------------------------"
set "doubleline=============================================================="
for /f "tokens=3" %%i in ('echo list volume ^| diskpart ^| findstr "16.0"') do (if exist "%%i:\Office" set batofficeltr=%%i&goto checklanguage)
goto batofficeno

:checklanguage
for /f "delims=. tokens=6" %%i in ('dir /b /s %batofficeltr%:\Office\Data ^| findstr "proof"') do (set batofficelanguage=%%i&goto batofficemain)

:batofficemain
cls
title  Office Setup Bat %aboutver%
mode con cols=98 lines=30
echo=
echo=
echo                   ______________________________________________________________
echo                  ^|                                                              ^| 
echo                  ^|                                                              ^|
echo                  ^|     [0] Exit                                                 ^|
echo                  ^|                                                              ^|
echo                  ^|     [1] Install Office 2016                                  ^|
echo                  ^|                                                              ^|
echo                  ^|     [2] Install Office 2019                                  ^|
echo                  ^|                                                              ^|
echo                  ^|     [3] Install Office 2021                                  ^|
echo                  ^|                                                              ^|
echo                  ^|     [4] Install Office 365                                   ^|
echo                  ^|                                                              ^|
echo                  ^|______________________________________________________________^|
echo=
choice /C:012349 /N /M ">                       Enter Your Choice in the Keyboard [0,1,2,3,4] : "
if errorlevel  6 (set batofficever=mondo) & set verblank=& goto selectapp
if errorlevel  5 (set batofficever=365) & set verblank=  & goto selectapp
if errorlevel  4 (set batofficever=2021) & set verblank= & goto selectapp
if errorlevel  3 (set batofficever=2019) & set verblank= & goto selectapp
if errorlevel  2 (set batofficever=2016) & set verblank= & goto selectapp
if errorlevel  1 goto end

:selectapp
cls
title  Office Setup Bat %aboutver%
mode con cols=98 lines=30
set AIDAccess=1
set AIDExcel=1
set AIDGroove=0
set AIDLync=0
set AIDOneDrive=1
set AIDOneNote=1
set AIDOutlook=1
set AIDPowerPoint=1
set AIDPublisher=1
set AIDWord=1
set AIDTeams=1
set AIDBing=1
set AIDProject=0
set AIDVisio=0
set AIDGrooveV=0
set AIDOneDriveV=1
echo=
echo=
echo                   ______________________________________________________________
echo                  ^|                                                              ^| 
echo                  ^|                                                              ^|
echo                  ^|     [0] Return To Version Choice                             ^|
echo                  ^|                                                              ^|
echo                  ^|     [1] Install Office %batofficever% mini(word excle powerpoint)%verblank%     ^|
echo                  ^|                                                              ^|
echo                  ^|     [2] Install Office %batofficever% normal%verblank%                          ^|
echo                  ^|                                                              ^|
echo                  ^|     [3] Install Project %batofficever% %verblank%                               ^|
echo                  ^|                                                              ^|
echo                  ^|     [4] Install Visio %batofficever% %verblank%                                 ^|
echo                  ^|                                                              ^|
echo                  ^|     [5] Custom Office APP                                    ^|
echo                  ^|                                                              ^|
echo                  ^|______________________________________________________________^|
echo=
choice /C:012345 /N /M ">                       Enter Your Choice in the Keyboard [0,1,2,3,4,5] : "
if errorlevel  6 (set batofficetype=Office %batofficever% custom) & set typeblank=& goto customsetup
if errorlevel  5 (set batofficetype=Visio %batofficever%) & set typeblank=        & goto Visiosetup
if errorlevel  4 (set batofficetype=Project %batofficever%) & set typeblank=      & goto Projectsetup
if errorlevel  3 (set batofficetype=Office %batofficever% normal) & set typeblank=& goto normalsetup
if errorlevel  2 (set batofficetype=Office %batofficever% mini) & set typeblank=  & goto minisetup
if errorlevel  1 goto batofficemain

:minisetup
cls
set appnumber=
set AIDAccess=0
set AIDExcel=1
set AIDGroove=0
set AIDLync=0
set AIDOneDrive=0
set AIDOneNote=0
set AIDOutlook=0
set AIDPowerPoint=1
set AIDPublisher=0
set AIDWord=1
set AIDTeams=0
set AIDBing=0
set AIDProject=0
set AIDVisio=0
set AIDGrooveV=0
set AIDOneDriveV=0
echo   %batofficetype%
echo %doubleline%
echo      - Word        :  *
echo      - Excel       :  *
echo      - PowerPoint  :  *
echo=%singleline%
echo   0  - Last
echo   99 - Next
echo=%doubleline%
set /p appnumber= ^> Enter your option and press "Enter": 
if "%appnumber%" == "0" goto selectapp
if "%appnumber%" == "99" goto selectchannel
goto minisetup

:normalsetup
cls
set appnumber=
set AIDAccess=1
set AIDExcel=1
set AIDGroove=0
set AIDLync=0
set AIDOneDrive=0
set AIDOneNote=1
set AIDOutlook=1
set AIDPowerPoint=1
set AIDPublisher=1
set AIDWord=1
set AIDTeams=0
set AIDBing=0
if "%batofficever%" == "mondo" (
    set AIDVisio=1
    set AIDProject=1
) else (
    set AIDVisio=0
    set AIDProject=0
)
set AIDGrooveV=0
set AIDOneDriveV=0
echo   %batofficetype%
echo %doubleline%
echo      - Access      :  *
echo      - Excel       :  *
echo      - OneNote     :  *
echo      - Outlook     :  *
echo      - PowerPoint  :  *
echo      - Publisher   :  *
echo      - Word        :  *
if "%batofficever%" == "mondo" (echo      - Project     :  *)
if "%batofficever%" == "mondo" (echo      - Visio       :  *)
echo %singleline%
echo   0  - Last
echo   99 - Next
echo %doubleline%
set /p appnumber= ^> Enter your option and press "Enter": 
if "%appnumber%" == "0" goto selectapp
if "%appnumber%" == "99" goto selectchannel
goto normalsetup

:Projectsetup
cls
set appnumber=
set AIDAccess=0
set AIDExcel=0
set AIDGroove=0
set AIDLync=0
set AIDOneDrive=0
set AIDOneNote=0
set AIDOutlook=0
set AIDPowerPoint=0
set AIDPublisher=0
set AIDWord=0
set AIDTeams=0
set AIDBing=0
set AIDProject=1
set AIDVisio=0
set AIDGrooveV=0
set AIDOneDriveV=0
echo   %batofficetype%
echo %doubleline%
echo      - Project     :  *
echo %singleline%
echo   0  - Last
echo   99 - Next
echo %doubleline%
set /p appnumber= ^> Enter your option and press "Enter": 
if "%appnumber%" == "0" goto selectapp
if "%appnumber%" == "99" goto selectchannel
goto Projectsetup

:Visiosetup
cls
set appnumber=
set AIDAccess=0
set AIDExcel=0
set AIDGroove=0
set AIDLync=0
set AIDOneDrive=0
set AIDOneNote=0
set AIDOutlook=0
set AIDPowerPoint=0
set AIDPublisher=0
set AIDWord=0
set AIDTeams=0
set AIDBing=0
set AIDProject=0
set AIDVisio=1
echo   %batofficetype%
echo %doubleline%
echo      - Visio       :  *
echo %singleline%
if %AIDVisio% equ 1 (
    if %AIDGrooveV% equ 1 (echo   13 - Groove^(V^)   :  * ) else (echo   13 - Groove^(V^)   :    )
    if %AIDOneDriveV% equ 1 (echo   14 - OneDrive^(V^) :  * ) else (echo   14 - OneDrive^(V^) :    )
)
echo=
echo   0  - Last
echo   99 - Next
echo %doubleline%
set /p appnumber= ^> Enter your option and press "Enter": 
if "%appnumber%" == "0" goto selectapp
if "%appnumber%" == "99" goto selectchannel
if "%appnumber%" == "13" (if %AIDGrooveV% equ 0 (set AIDGrooveV=1) else (set AIDGrooveV=0)) & goto Visiosetup
if "%appnumber%" == "14" (if %AIDOneDriveV% equ 0 (set AIDOneDriveV=1) else (set AIDOneDriveV=0)) & goto Visiosetup
goto Visiosetup

:customsetup
cls
set appnumber=
if %AIDVisio% equ 0 (
    set AIDGrooveV=0
    set AIDOneDriveV=1
)
echo   %batofficetype%
echo %doubleline%
if %AIDAccess% equ 1 (echo   1  - Access      :  * ) else (echo   1  - Access      :    )
if %AIDExcel% equ 1 (echo   2  - Excel       :  * ) else (echo   2  - Excel       :    )
if %AIDGroove% equ 1 (echo   3  - Groove      :  * ) else (echo   3  - Groove      :    )
if %AIDLync% equ 1 (echo   4  - Lync        :  * ) else (echo   4  - Lync        :    )
if %AIDOneDrive% equ 1 (echo   5  - OneDrive    :  * ) else (echo   5  - OneDrive    :    )
if %AIDOneNote% equ 1 (echo   6  - OneNote     :  * ) else (echo   6  - OneNote     :    )
if %AIDOutlook% equ 1 (echo   7  - Outlook     :  * ) else (echo   7  - Outlook     :    )
if %AIDPowerPoint% equ 1 (echo   8  - PowerPoint  :  * ) else (echo   8  - PowerPoint  :    )
if %AIDPublisher% equ 1 (echo   9  - Publisher   :  * ) else (echo   9  - Publisher   :    )
if %AIDWord% equ 1 (echo   10 - Word        :  * ) else (echo   10 - Word        :    )
if %AIDProject% equ 1 (echo   11 - Project     :  * ) else (echo   11 - Project     :    )
echo -------------------------
if %AIDVisio% equ 1 (echo   12 - Visio       :  *  ) else (echo   12 - Visio       :     )
if %AIDVisio% equ 1 (
    if %AIDGrooveV% equ 1 (echo   13 - Groove^(V^)   :  *  ) else (echo   13 - Groove^(V^)   :     )
    if %AIDOneDriveV% equ 1 (echo   14 - OneDrive^(V^) :  *  ) else (echo   14 - OneDrive^(V^) :     )
) else (
    echo   13 - Groove^(V^)   :  -  
    echo   14 - OneDrive^(V^) :  -  
)
echo -------------------------
if "%batofficever%" == "2021" (
    if %AIDTeams% equ 1 (echo   15 - Teams       :  * ) else (echo   15 - Teams       :    )
)
if "%batofficever%" == "365" (
    if %AIDTeams% equ 1 (echo   15 - Teams       :  * ) else (echo   15 - Teams       :    )
)
if "%batofficever%" == "365" (
    if %AIDBing% equ 1 (echo   16 - Bing        :  * ) else (echo   16 - Bing        :    )
)
echo=
echo   0  - Last
echo   99 - Next
echo %doubleline%
set /p appnumber= ^> Enter your option and press "Enter": 
if "%appnumber%" == "0" goto selectapp
if "%appnumber%" == "99" goto selectchannel
if "%appnumber%" == "1" (if %AIDAccess% equ 0 (set AIDAccess=1) else (set AIDAccess=0)) & goto customsetup
if "%appnumber%" == "2" (if %AIDExcel% equ 0 (set AIDExcel=1) else (set AIDExcel=0)) & goto customsetup
if "%appnumber%" == "3" (if %AIDGroove% equ 0 (set AIDGroove=1) else (set AIDGroove=0)) & goto customsetup
if "%appnumber%" == "4" (if %AIDLync% equ 0 (set AIDLync=1) else (set AIDLync=0)) & goto customsetup
if "%appnumber%" == "5" (if %AIDOneDrive% equ 0 (set AIDOneDrive=1) else (set AIDOneDrive=0)) & goto customsetup
if "%appnumber%" == "6" (if %AIDOneNote% equ 0 (set AIDOneNote=1) else (set AIDOneNote=0)) & goto customsetup
if "%appnumber%" == "7" (if %AIDOutlook% equ 0 (set AIDOutlook=1) else (set AIDOutlook=0)) & goto customsetup
if "%appnumber%" == "8" (if %AIDPowerPoint% equ 0 (set AIDPowerPoint=1) else (set AIDPowerPoint=0)) & goto customsetup
if "%appnumber%" == "9" (if %AIDPublisher% equ 0 (set AIDPublisher=1) else (set AIDPublisher=0)) & goto customsetup
if "%appnumber%" == "10" (if %AIDWord% equ 0 (set AIDWord=1) else (set AIDWord=0)) & goto customsetup
if "%appnumber%" == "11" (if %AIDProject% equ 0 (set AIDProject=1) else (set AIDProject=0)) & goto customsetup
if "%appnumber%" == "12" (if %AIDVisio% equ 0 (set AIDVisio=1) else (set AIDVisio=0)) & goto customsetup
if "%appnumber%" == "13" (
    if %AIDVisio% equ 1 (
        if %AIDGrooveV% equ 0 (set AIDGrooveV=1) else (set AIDGrooveV=0)
    ) else (set AIDGrooveV=0)
) & goto customsetup
if "%appnumber%" == "14" (
    if %AIDVisio% equ 1 (
        if %AIDOneDriveV% equ 0 (set AIDOneDriveV=1) else (set AIDOneDriveV=0)
    ) else (set AIDOneDriveV=1)
) & goto customsetup
if "%appnumber%" == "15" (
    if "%batofficever%" == "2021" (if %AIDTeams% equ 0 (set AIDTeams=1) else (set AIDTeams=0))
    if "%batofficever%" == "365" (if %AIDTeams% equ 0 (set AIDTeams=1) else (set AIDTeams=0))
) & goto customsetup
if "%appnumber%" == "16" (
    if "%batofficever%" == "365" (if %AIDBing% equ 0 (set AIDBing=1) else (set AIDBing=0))
) & goto customsetup
goto customsetup

:selectchannel
cls
title  Office Setup Bat %aboutver%
mode con cols=98 lines=30
echo=
echo=
echo                   ______________________________________________________________
echo                  ^|                                                              ^| 
echo                  ^|                                                              ^|
echo                  ^|     [0] Return To APP Choice                                 ^|
echo                  ^|                                                              ^|
echo                  ^|     [1] select %batofficetype% Channel: Current%verblank%%typeblank%          ^|
echo                  ^|                                                              ^|
echo                  ^|     [2] select %batofficetype% Channel: Monthly Ent%verblank%%typeblank%      ^|
echo                  ^|                                                              ^|
echo                  ^|     [3] select %batofficetype% Channel: Semi-Annual Ent%verblank%%typeblank%  ^|
echo                  ^|                                                              ^|
echo                  ^|______________________________________________________________^|
echo=
choice /C:01234 /N /M ">                       Enter Your Choice in the Keyboard [0,1,2,3] : "
if errorlevel  4 (set batofficechannel=SemiAnnual) & set channelblank=       & goto checkedition
if errorlevel  3 (set batofficechannel=MonthlyEnterprise) & set channelblank=& goto checkedition
if errorlevel  2 (set batofficechannel=Current) & set channelblank=          & goto checkedition
if errorlevel  1 goto selectapp

:checkedition
if "%PROCESSOR_ARCHITEW6432%"=="AMD64" (
    goto WIN64
) else if "%PROCESSOR_ARCHITECTURE%"=="AMD64" (
    goto WIN64
) else if "%PROCESSOR_ARCHITECTURE%"=="x86" (
    goto WIN32
) else (
    echo Not support this operating system
    goto end
)

:WIN64
cls
title  Office Setup Bat %aboutver%
mode con cols=98 lines=30
echo=
echo=
echo                   ______________________________________________________________
echo                  ^|                                                              ^| 
echo                  ^|                                                              ^|
echo                  ^|     [0] Return To Channel Choice                             ^|
echo                  ^|                                                              ^|
echo                  ^|     [1] Install %batofficetype% 32Bit(%batofficechannel%)%verblank%%typeblank%%channelblank% ^|
echo                  ^|                                                              ^|
echo                  ^|     [2] Install %batofficetype% 64Bit(%batofficechannel%)%verblank%%typeblank%%channelblank% ^|
echo                  ^|                                                              ^|
echo                  ^|______________________________________________________________^|
echo=          
choice /C:012 /N /M ">                       Enter Your Choice in the Keyboard [0,1,2] : "
if errorlevel  3 (set batofficeedition=64) & goto batofficexml
if errorlevel  2 (set batofficeedition=32) & goto batofficexml
if errorlevel  1 goto selectchannel

:WIN32
cls
title  Office Setup Bat %aboutver%
mode con cols=98 lines=30
echo=
echo=
echo                   ______________________________________________________________
echo                  ^|                                                              ^| 
echo                  ^|                                                              ^|
echo                  ^|     [0] Return To Channel Choice                             ^|
echo                  ^|                                                              ^|
echo                  ^|     [1] Install %batofficetype% 32Bit(%batofficechannel%)%typeblank%%verblank%%channelblank% ^|
echo                  ^|                                                              ^|
echo                  ^|______________________________________________________________^|
echo=
choice /C:012 /N /M ">                       Enter Your Choice in the Keyboard [0,1,2] : "
if errorlevel  2 (set batofficeedition=32) & goto batofficexml
if errorlevel  1 goto cn

:batofficexml
if "%batofficever%" == "2016" (
    set /a AIDoffice=%AIDAccess%+%AIDExcel%+%AIDGroove%+%AIDLync%+%AIDOneDrive%+%AIDOneNote%+%AIDOutlook%+%AIDPowerPoint%+%AIDPublisher%+%AIDWord%
    if !AIDoffice! neq 0 (set PIDoffice=ProPlusRetail)
    if %AIDProject% neq 0 (set PIDProject=ProjectProRetail)
    if %AIDVisio% neq 0 (set PIDVisio=VisioProRetail)
)else if "%batofficever%" == "2019" (
    set /a AIDoffice=%AIDAccess%+%AIDExcel%+%AIDGroove%+%AIDLync%+%AIDOneDrive%+%AIDOneNote%+%AIDOutlook%+%AIDPowerPoint%+%AIDPublisher%+%AIDWord%
    if !AIDoffice! neq 0 (set PIDoffice=ProPlus2019Retail)
    if %AIDProject% neq 0 (set PIDProject=ProjectPro2019Retail)
    if %AIDVisio% neq 0 (set PIDVisio=VisioPro2019Retail)
)else if "%batofficever%" == "2021" (
    set /a AIDoffice=%AIDAccess%+%AIDExcel%+%AIDGroove%+%AIDLync%+%AIDOneDrive%+%AIDOneNote%+%AIDOutlook%+%AIDPowerPoint%+%AIDPublisher%+%AIDWord%+%AIDTeams%
    if !AIDoffice! neq 0 (set PIDoffice=ProPlus2021Retail)
    if %AIDProject% neq 0 (set PIDProject=ProjectPro2021Retail)
    if %AIDVisio% neq 0 (set PIDVisio=VisioPro2021Retail)
)else if "%batofficever%" == "365" (
    set /a AIDoffice=%AIDAccess%+%AIDExcel%+%AIDGroove%+%AIDLync%+%AIDOneDrive%+%AIDOneNote%+%AIDOutlook%+%AIDPowerPoint%+%AIDPublisher%+%AIDWord%+%AIDTeams%+%AIDBing%
    if !AIDoffice! neq 0 (set PIDoffice=O365ProPlusRetail)
    if %AIDProject% neq 0 (set PIDProject=ProjectProRetail)
    if %AIDVisio% neq 0 (set PIDVisio=VisioProRetail)
)else (
    set /a AIDoffice=%AIDAccess%+%AIDExcel%+%AIDGroove%+%AIDLync%+%AIDOneDrive%+%AIDOneNote%+%AIDOutlook%+%AIDPowerPoint%+%AIDPublisher%+%AIDWord%+%AIDProject%+%AIDVisio%
    if !AIDoffice! neq 0 (set PIDoffice=MondoRetail)
)
set /a AIDPID=%AIDoffice%+%AIDProject%+%AIDVisio%
if %AIDPID% neq 0 (
    echo ^<Configuration^>>batofficetemp.xml
    echo   ^<Add OfficeClientEdition="%batofficeedition%" Channel="%batofficechannel%" SourcePath="%batofficeltr%:"^>>>batofficetemp.xml
    if %AIDoffice% neq 0 (
        echo     ^<Product ID="%PIDoffice%"^>>>batofficetemp.xml
        echo       ^<Language ID="%batofficelanguage%" /^>>>batofficetemp.xml
        if %AIDAccess% equ 0 (echo       ^<ExcludeApp ID="Access" /^>>>batofficetemp.xml)
        if %AIDExcel% equ 0 (echo       ^<ExcludeApp ID="Excel" /^>>>batofficetemp.xml)
        if %AIDGroove% equ 0 (echo       ^<ExcludeApp ID="Groove" /^>>>batofficetemp.xml)
        if %AIDLync% equ 0 (echo       ^<ExcludeApp ID="Lync" /^>>>batofficetemp.xml)
        if %AIDOneDrive% equ 0 (echo       ^<ExcludeApp ID="OneDrive" /^>>>batofficetemp.xml)
        if %AIDOneNote% equ 0 (echo       ^<ExcludeApp ID="OneNote" /^>>>batofficetemp.xml)
        if %AIDOutlook% equ 0 (echo       ^<ExcludeApp ID="Outlook" /^>>>batofficetemp.xml)
        if %AIDPowerPoint% equ 0 (echo       ^<ExcludeApp ID="PowerPoint" /^>>>batofficetemp.xml)
        if %AIDPublisher% equ 0 (echo       ^<ExcludeApp ID="Publisher" /^>>>batofficetemp.xml)
        if %AIDWord% equ 0 (echo       ^<ExcludeApp ID="Word" /^>>>batofficetemp.xml)
        if "%batofficever%" == "2021" (
            if %AIDTeams% equ 0 (echo       ^<ExcludeApp ID="Teams" /^>>>batofficetemp.xml)
        )
        if "%batofficever%" == "365" (
            if %AIDTeams% equ 0 (echo       ^<ExcludeApp ID="Teams" /^>>>batofficetemp.xml)
        )
        if "%batofficever%" == "365" (
            if %AIDBing% equ 0 (echo       ^<ExcludeApp ID="Bing" /^>>>batofficetemp.xml)
        )
        if "%batofficever%" == "mondo" (
            if %AIDProject% equ 0 (echo       ^<ExcludeApp ID="Project" /^>>>batofficetemp.xml)
            if %AIDVisio% equ 0 (echo       ^<ExcludeApp ID="Visio" /^>>>batofficetemp.xml)
        )
        echo     ^</Product^>>>batofficetemp.xml
    )
    
    if not "%batofficever%" == "mondo" (
        if %AIDProject% neq 0 (
            echo     ^<Product ID="%PIDProject%"^>>>batofficetemp.xml
            echo       ^<Language ID="%batofficelanguage%" /^>>>batofficetemp.xml
            echo     ^</Product^>>>batofficetemp.xml
        )        
    )

    if not "%batofficever%" == "mondo" (
        if %AIDVisio% neq 0 (
            echo     ^<Product ID="%PIDVisio%"^>>>batofficetemp.xml
            echo       ^<Language ID="%batofficelanguage%" /^>>>batofficetemp.xml
            if %AIDGrooveV% equ 0 (echo       ^<ExcludeApp ID="Groove" /^>>>batofficetemp.xml)
            if %AIDOneDriveV% equ 0 (echo       ^<ExcludeApp ID="OneDrive" /^>>>batofficetemp.xml)
            echo     ^</Product^>>>batofficetemp.xml
        )

    )
    echo   ^</Add^>>>batofficetemp.xml
    echo ^</Configuration^>>>batofficetemp.xml
    goto batofficesetup
) else goto noselectapp

:batofficesetup
setup.exe /configure batofficetemp.xml
if %errorlevel%==0 (
    echo Installation completed.
    del /f /q batofficetemp.xml >nul 2>nul
    pause
    exit
) else (
    echo Installation failed,errorlevel:%errorlevel%
    del /f /q batofficetemp.xml >nul 2>nul
    goto end
)

:noselectapp
echo No select any installations.
goto end

:batofficeno
echo No find any Office images,please mount a Office image first.
goto end

:end
echo Installation has exited.
pause
