@echo off
setlocal enabledelayedexpansion
%1 mshta vbscript:CreateObject("Shell.Application").ShellExecute("cmd.exe","/c %~s0 ::","","runas",1)(window.close)&&exit
cd /d %~dp0
set "singleline=-------------------------------------------------------------"
set "doubleline=============================================================="
for /f "tokens=3" %%i in ('echo list volume ^| diskpart ^| findstr "16.0"') do (
    if exist "%%i:\Office" set officedir=%%i&goto officemain
    
)
goto officeno

:officemain
cls
title  office 2016/2019/2021/365安装向导
mode con cols=98 lines=30
echo=
echo=
echo                   ______________________________________________________________
echo                  ^|                                                              ^| 
echo                  ^|                                                              ^|
echo                  ^|     [0] 退出安装                                             ^|
echo                  ^|                                                              ^|
echo                  ^|     [1] 安装 office 2016                                     ^|
echo                  ^|                                                              ^|
echo                  ^|     [2] 安装 office 2019                                     ^|
echo                  ^|                                                              ^|
echo                  ^|     [3] 安装 office 2021                                     ^|
echo                  ^|                                                              ^|
echo                  ^|     [4] 安装 office 365                                      ^|
echo                  ^|                                                              ^|
echo                  ^|______________________________________________________________^|
echo=          
choice /C:01234 /N /M ">                       请在键盘中输入你的选择 [0、1、2、3、4] : "
if errorlevel  5 (set officever=365) & set verkg=  & goto capp
if errorlevel  4 (set officever=2021) & set verkg= & goto capp
if errorlevel  3 (set officever=2019) & set verkg= & goto capp
if errorlevel  2 (set officever=2016) & set verkg= & goto capp
if errorlevel  1 goto end

:capp
cls
title  office 2016/2019/2021/365安装向导
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
set AIDProject=0
set AIDVisio=0
set AIDTeams=1
set AIDGrooveV=0
set AIDOneDriveV=1
echo=
echo=
echo                   ______________________________________________________________
echo                  ^|                                                              ^| 
echo                  ^|                                                              ^|
echo                  ^|     [0] 返回选择版本                                         ^|
echo                  ^|                                                              ^|
echo                  ^|     [1] 安装 office %officever% mini(word excle powerpoint)%verkg%        ^|
echo                  ^|                                                              ^|
echo                  ^|     [2] 安装 office %officever% normal%verkg%                             ^|
echo                  ^|                                                              ^|
echo                  ^|     [3] 安装 Project %officever% %verkg%                                  ^|
echo                  ^|                                                              ^|
echo                  ^|     [4] 安装 Visio %officever% %verkg%                                    ^|
echo                  ^|                                                              ^|
echo                  ^|     [5] 自定义office组件                                     ^|
echo                  ^|                                                              ^|
echo                  ^|______________________________________________________________^|
echo=          
choice /C:012345 /N /M ">                       请在键盘中输入你的选择 [0、1、2、3、4、5] : "
if errorlevel  6 (set offictype=office %officever% custom) & set typekg= & goto Scustom
if errorlevel  5 (set offictype=Visio %officever%) & set typekg=         & goto SVisio
if errorlevel  4 (set offictype=Project %officever%) & set typekg=       & goto SProject
if errorlevel  3 (set offictype=office %officever% normal) & set typekg= & goto Snormal
if errorlevel  2 (set offictype=office %officever% mini) & set typekg=   & goto Smini
if errorlevel  1 goto officemain

:Smini
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
set AIDVisio=0
set AIDProject=0
set AIDGrooveV=0
set AIDOneDriveV=0
echo   %offictype%
echo %doubleline%
echo      - Word        :  *
echo      - Excel       :  *
echo      - PowerPoint  :  *
echo=%singleline%
echo   0  - 上一步
echo   99 - 下一步
echo=%doubleline%
set /p appnumber= ^> 请输入选项编号并按“Enter”键: 
if not defined appnumber goto Smini
set appnumber=%appnumber:~0,2%
if %appnumber% equ 0 goto capp
if %appnumber% equ 99 goto Cchannel
goto Smini

:Snormal
cls
set appnumber=
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
if %officever% equ 365 (
    set AIDVisio=0
    set AIDProject=0
) else (
    set AIDVisio=0
    set AIDProject=0
)
set AIDGrooveV=0
set AIDOneDriveV=1
echo   %offictype%
echo %doubleline%
echo      - Access      :  *
echo      - Excel       :  *
echo      - OneDrive    :  *
echo      - OneNote     :  *
echo      - Outlook     :  *
echo      - PowerPoint  :  *
echo      - Publisher   :  *
echo      - Word        :  *
if %officever% equ 2021 (
    echo      - Teams       :  *
)
echo %singleline%
echo   0  - 上一步
echo   99 - 下一步
echo %doubleline%
set /p appnumber= ^> 请输入选项编号并按“Enter”键: 
if not defined appnumber goto Snormal
set appnumber=%appnumber:~0,2%
if %appnumber% equ 0 goto capp
if %appnumber% equ 99 goto Cchannel
goto Snormal

:SProject
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
set AIDProject=1
set AIDVisio=0
set AIDTeams=0
set AIDGrooveV=0
set AIDOneDriveV=0
echo   %offictype%
echo %doubleline%
echo      - Project     :  *
echo %singleline%
echo   0  - 上一步
echo   99 - 下一步
echo %doubleline%
set /p appnumber= ^> 请输入选项编号并按“Enter”键: 
if not defined appnumber goto SProject
set appnumber=%appnumber:~0,2%
if %appnumber% equ 0 goto capp
if %appnumber% equ 99 goto Cchannel
goto SProject

:SVisio
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
set AIDProject=0
set AIDVisio=1
set AIDTeams=0
echo   %offictype%
echo %doubleline%
echo      - Visio       :  *
echo %singleline%
if %AIDVisio% equ 1 (
    if %AIDGrooveV% equ 1 (echo   13 - Groove^(V^)   :  * ) else (echo   13 - Groove^(V^)   :    )
    if %AIDOneDriveV% equ 1 (echo   14 - OneDrive^(V^) :  * ) else (echo   14 - OneDrive^(V^) :    )
)
echo=
echo   0  - 上一步
echo   99 - 下一步
echo %doubleline%
set /p appnumber= ^> 请输入选项编号并按“Enter”键: 
if not defined appnumber goto SVisio
set appnumber=%appnumber:~0,2%
if %appnumber% equ 0 goto capp
if %appnumber% equ 99 goto Cchannel
if %appnumber% equ 13 (if %AIDGrooveV% equ 0 (set AIDGrooveV=1) else (set AIDGrooveV=0)) & goto SVisio
if %appnumber% equ 14 (if %AIDOneDriveV% equ 0 (set AIDOneDriveV=1) else (set AIDOneDriveV=0)) & goto SVisio
goto SVisio

:Scustom
cls
set appnumber=
if %AIDVisio% equ 0 (
    set AIDGrooveV=0
    set AIDOneDriveV=1
)
echo   %offictype%
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
if %officever% equ 2021 (
    if %AIDTeams% equ 1 (echo   15 - Teams       :  * ) else (echo   15 - Teams       :    )
)
echo=
echo   0  - 上一步
echo   99 - 下一步
echo %doubleline%
set /p appnumber= ^> 请输入选项编号并按“Enter”键: 
if not defined appnumber goto Scustom
set appnumber=%appnumber:~0,2%
if %appnumber% equ 0 goto capp
if %appnumber% equ 99 goto Cchannel
if %appnumber% equ 1 (if %AIDAccess% equ 0 (set AIDAccess=1) else (set AIDAccess=0)) & goto Scustom
if %appnumber% equ 2 (if %AIDExcel% equ 0 (set AIDExcel=1) else (set AIDExcel=0)) & goto Scustom
if %appnumber% equ 3 (if %AIDGroove% equ 0 (set AIDGroove=1) else (set AIDGroove=0)) & goto Scustom
if %appnumber% equ 4 (if %AIDLync% equ 0 (set AIDLync=1) else (set AIDLync=0)) & goto Scustom
if %appnumber% equ 5 (if %AIDOneDrive% equ 0 (set AIDOneDrive=1) else (set AIDOneDrive=0)) & goto Scustom
if %appnumber% equ 6 (if %AIDOneNote% equ 0 (set AIDOneNote=1) else (set AIDOneNote=0)) & goto Scustom
if %appnumber% equ 7 (if %AIDOutlook% equ 0 (set AIDOutlook=1) else (set AIDOutlook=0)) & goto Scustom
if %appnumber% equ 8 (if %AIDPowerPoint% equ 0 (set AIDPowerPoint=1) else (set AIDPowerPoint=0)) & goto Scustom
if %appnumber% equ 9 (if %AIDPublisher% equ 0 (set AIDPublisher=1) else (set AIDPublisher=0)) & goto Scustom
if %appnumber% equ 10 (if %AIDWord% equ 0 (set AIDWord=1) else (set AIDWord=0)) & goto Scustom
if %appnumber% equ 11 (if %AIDProject% equ 0 (set AIDProject=1) else (set AIDProject=0)) & goto Scustom
if %appnumber% equ 12 (if %AIDVisio% equ 0 (set AIDVisio=1) else (set AIDVisio=0)) & goto Scustom
if %appnumber% equ 13 (
    if %AIDVisio% equ 1 (
        if %AIDGrooveV% equ 0 (set AIDGrooveV=1) else (set AIDGrooveV=0)
    ) else (set AIDGrooveV=0)
) & goto Scustom
if %appnumber% equ 14 (
    if %AIDVisio% equ 1 (
        if %AIDOneDriveV% equ 0 (set AIDOneDriveV=1) else (set AIDOneDriveV=0)
    ) else (set AIDOneDriveV=1)
) & goto Scustom
if %appnumber% equ 15 (
    if %officever% equ 2021 (if %AIDTeams% equ 0 (set AIDTeams=1) else (set AIDTeams=0))
) & goto Scustom
goto Scustom

:Cchannel
cls
title  office 2016/2019/2021/365安装向导
mode con cols=98 lines=30
echo=
echo=
echo                   ______________________________________________________________
echo                  ^|                                                              ^| 
echo                  ^|                                                              ^|
echo                  ^|     [0] 返回选择组件                                         ^|
echo                  ^|                                                              ^|
echo                  ^|     [1] 选择 %offictype% 更新通道: 当前通道%typekg%%verkg%         ^|
echo                  ^|                                                              ^|
echo                  ^|     [2] 选择 %offictype% 更新通道: 月度企业通道%typekg%%verkg%     ^|
echo                  ^|                                                              ^|
echo                  ^|     [3] 选择 %offictype% 更新通道: 半年度企业通道%typekg%%verkg%   ^|
echo                  ^|                                                              ^|
echo                  ^|                                                              ^|
echo                  ^|______________________________________________________________^|
echo=          
choice /C:01234 /N /M ">                       请在键盘中输入你的选择 [0、1、2、3] : "
if errorlevel  4 (set officecn=SemiAnnual) & set cnkg=        & goto ed
if errorlevel  3 (set officecn=MonthlyEnterprise) & set cnkg= & goto ed
if errorlevel  2 (set officecn=Current) & set cnkg=           & goto ed
if errorlevel  1 goto capp

:ed
if "%PROCESSOR_ARCHITEW6432%"=="AMD64" (
    goto ed64main
) else if "%PROCESSOR_ARCHITECTURE%"=="AMD64" (
    goto ed64main
) else if "%PROCESSOR_ARCHITECTURE%"=="x86" (
    goto ed32main
) else (
    echo 不支持当前的操作系统
    goto end
)

:ed64main
cls
title  office 2016/2019/2021/365安装向导
mode con cols=98 lines=30
echo=
echo=
echo                   ______________________________________________________________
echo                  ^|                                                              ^| 
echo                  ^|                                                              ^|
echo                  ^|     [0] 返回选择通道                                         ^|
echo                  ^|                                                              ^|
echo                  ^|     [1] 安装32位 %offictype% ( %officecn% )%typekg%%verkg%%cnkg% ^|
echo                  ^|                                                              ^|
echo                  ^|     [2] 安装64位 %offictype% ( %officecn% )%typekg%%verkg%%cnkg% ^|
echo                  ^|                                                              ^|
echo                  ^|______________________________________________________________^|
echo=          
choice /C:012 /N /M ">                       请在键盘中输入你的选择 [0、1、2] : "
if errorlevel  3 (set officeed=64) & goto officexml
if errorlevel  2 (set officeed=32) & goto officexml
if errorlevel  1 goto Cchannel

:ed32main
cls
title  office 2016/2019/2021/365安装向导
mode con cols=98 lines=30
echo=
echo=
echo                   ______________________________________________________________
echo                  ^|                                                              ^| 
echo                  ^|                                                              ^|
echo                  ^|     [0] 返回选择通道                                         ^|
echo                  ^|                                                              ^|
echo                  ^|     [1] 安装32位 %offictype% ( %officecn% )%typekg%%verkg%%cnkg% ^|
echo                  ^|                                                              ^|
echo                  ^|______________________________________________________________^|
echo=          
choice /C:012 /N /M ">                       请在键盘中输入你的选择 [0、1、2] : "
if errorlevel  2 (set officeed=32) & goto officexml
if errorlevel  1 goto cn

:officexml
if %officever% equ 2016 (
    set /a AIDoffice=%AIDAccess%+%AIDExcel%+%AIDGroove%+%AIDLync%+%AIDOneDrive%+%AIDOneNote%+%AIDOutlook%+%AIDPowerPoint%+%AIDPublisher%+%AIDWord%
    if !AIDoffice! neq 0 (set PIDoffice=ProPlusRetail)
    if %AIDProject% neq 0 (set PIDProject=ProjectProRetail)
    if %AIDVisio% neq 0 (set PIDVisio=VisioProRetail)
)else if %officever% equ 2019 (
    set /a AIDoffice=%AIDAccess%+%AIDExcel%+%AIDGroove%+%AIDLync%+%AIDOneDrive%+%AIDOneNote%+%AIDOutlook%+%AIDPowerPoint%+%AIDPublisher%+%AIDWord%
    if !AIDoffice! neq 0 (set PIDoffice=ProPlus2019Retail)
    if %AIDProject% neq 0 (set PIDProject=ProjectPro2019Retail)
    if %AIDVisio% neq 0 (set PIDVisio=VisioPro2019Retail)
)else if %officever% equ 2021 (
    set /a AIDoffice=%AIDAccess%+%AIDExcel%+%AIDGroove%+%AIDLync%+%AIDOneDrive%+%AIDOneNote%+%AIDOutlook%+%AIDPowerPoint%+%AIDPublisher%+%AIDWord%+%AIDTeams%
    if !AIDoffice! neq 0 (set PIDoffice=ProPlus2021Retail)
    if %AIDProject% neq 0 (set PIDProject=ProjectPro2021Retail)
    if %AIDVisio% neq 0 (set PIDVisio=VisioPro2021Retail)
)else (
    set /a AIDoffice=%AIDAccess%+%AIDExcel%+%AIDGroove%+%AIDLync%+%AIDOneDrive%+%AIDOneNote%+%AIDOutlook%+%AIDPowerPoint%+%AIDPublisher%+%AIDWord%+%AIDProject%+%AIDVisio%
    if !AIDoffice! neq 0 (set PIDoffice=MondoRetail)
)
set /a AIDPID=%AIDoffice%+%AIDProject%+%AIDVisio%
if %AIDPID% neq 0 (
    echo ^<Configuration^>>officetemp.xml
    echo   ^<Add OfficeClientEdition="%officeed%" Channel="%officecn%" SourcePath="%officedir%:"^>>>officetemp.xml
    if %AIDoffice% neq 0 (
        echo     ^<Product ID="%PIDoffice%"^>>>officetemp.xml
        echo       ^<Language ID="zh-cn" /^>>>officetemp.xml
        if %AIDAccess% equ 0 (echo       ^<ExcludeApp ID="Access" /^>>>officetemp.xml)
        if %AIDExcel% equ 0 (echo       ^<ExcludeApp ID="Excel" /^>>>officetemp.xml)
        if %AIDGroove% equ 0 (echo       ^<ExcludeApp ID="Groove" /^>>>officetemp.xml)
        if %AIDLync% equ 0 (echo       ^<ExcludeApp ID="Lync" /^>>>officetemp.xml)
        if %AIDOneDrive% equ 0 (echo       ^<ExcludeApp ID="OneDrive" /^>>>officetemp.xml)
        if %AIDOneNote% equ 0 (echo       ^<ExcludeApp ID="OneNote" /^>>>officetemp.xml)
        if %AIDOutlook% equ 0 (echo       ^<ExcludeApp ID="Outlook" /^>>>officetemp.xml)
        if %AIDPowerPoint% equ 0 (echo       ^<ExcludeApp ID="PowerPoint" /^>>>officetemp.xml)
        if %AIDPublisher% equ 0 (echo       ^<ExcludeApp ID="Publisher" /^>>>officetemp.xml)
        if %AIDWord% equ 0 (echo       ^<ExcludeApp ID="Word" /^>>>officetemp.xml)
        if %officever% equ 365 (
            if %AIDProject% equ 0 (echo       ^<ExcludeApp ID="Project" /^>>>officetemp.xml)
            if %AIDVisio% equ 0 (echo       ^<ExcludeApp ID="Visio" /^>>>officetemp.xml)
        )
        if %officever% equ 2021 (
            if %AIDTeams% equ 0 (echo       ^<ExcludeApp ID="Teams" /^>>>officetemp.xml)
        )
        echo     ^</Product^>>>officetemp.xml
    )
    
    if %officever% neq 365 (
        if %AIDProject% neq 0 (
            echo     ^<Product ID="%PIDProject%"^>>>officetemp.xml
            echo       ^<Language ID="zh-cn" /^>>>officetemp.xml
            echo     ^</Product^>>>officetemp.xml
        )        
    )

    if %officever% neq 365 (
        if %AIDVisio% neq 0 (
            echo     ^<Product ID="%PIDVisio%"^>>>officetemp.xml
            echo       ^<Language ID="zh-cn" /^>>>officetemp.xml
            if %AIDGrooveV% equ 0 (echo       ^<ExcludeApp ID="Groove" /^>>>officetemp.xml)
            if %AIDOneDriveV% equ 0 (echo       ^<ExcludeApp ID="OneDrive" /^>>>officetemp.xml)
            echo     ^</Product^>>>officetemp.xml
        )

    )
    echo   ^</Add^>>>officetemp.xml
    echo ^</Configuration^>>>officetemp.xml
    goto officesetup
) else goto officenosetup

:officesetup
setup.exe /configure officetemp.xml
if %errorlevel%==0 (
    echo 安装已完成
    del /f /q officetemp.xml >nul 2>nul
    pause
    exit
) else (
    echo 安装失败,errorlevel:%errorlevel%
    del /f /q officetemp.xml >nul 2>nul
    goto end
)

:officenosetup
echo 没有选择任何安装
goto end

:officeno
echo 未检测到office镜像，请先挂载office镜像。
goto end

:end
echo 安装已退出
pause
