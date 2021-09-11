@echo off

%1 mshta vbscript:CreateObject("Shell.Application").ShellExecute("cmd.exe","/c %~s0 ::","","runas",1)(window.close)&&exit


cd /d %~dp0


for /f "tokens=3" %%i in ('echo list volume ^| diskpart ^| findstr "16.0"') do (
    if exist "%%i:\Office" set officedir=%%i&goto officemain
    
)
goto officeno


:officemain
cls
title  office 2016/2019/365安装向导
mode con cols=98 lines=30


echo:
echo:
echo                   _____________________________________________________________
echo                  ^|                                                             ^| 
echo                  ^|                                                             ^|
echo                  ^|      [0] 退出                                               ^|
echo                  ^|                                                             ^|
echo                  ^|      [1] 安装 office2016（word、excle、powerpoint）         ^|
echo                  ^|                                                             ^|
echo                  ^|      [2] 安装 office2019（word、excle、powerpoint）         ^|
echo                  ^|                                                             ^|
echo                  ^|      [3] 安装 office365（word、excle、powerpoint）          ^|
echo                  ^|                                                             ^|
echo                  ^|      [4] 安装 office365（ALL）                              ^|
echo                  ^|                                                             ^|
echo                  ^|_____________________________________________________________^|
echo:          
choice /C:01234 /N /M ">                         请在键盘中输入你的选择 [0、1、2、3、4] ："
if errorlevel  5 (set officever=office365all) & set verkg=    & goto cn
if errorlevel  4 (set officever=office365mini) & set verkg=   & goto cn
if errorlevel  3 (set officever=office2019mini) & set verkg=  & goto cn
if errorlevel  2 (set officever=office2016mini) & set verkg=  & goto cn
if errorlevel  1 echo 已退出 & goto end


:cn
cls
title  office 2016/2019/365安装向导
mode con cols=98 lines=30


echo:
echo:
echo                   _____________________________________________________________
echo                  ^|                                                             ^| 
echo                  ^|                                                             ^|
echo                  ^|      [0] 返回                                               ^|
echo                  ^|                                                             ^|
echo                  ^|      [1] 选择 %officever% 更新通道：当前通道 %verkg%          ^|
echo                  ^|                                                             ^|
echo                  ^|      [2] 选择 %officever% 更新通道：月度企业通道 %verkg%      ^|
echo                  ^|                                                             ^|
echo                  ^|      [3] 选择 %officever% 更新通道：半年度企业通道 %verkg%    ^|
echo                  ^|                                                             ^|
echo                  ^|                                                             ^|
echo                  ^|_____________________________________________________________^|
echo:          
choice /C:01234 /N /M ">                         请在键盘中输入你的选择 [0、1、2、3、4] ："
if errorlevel  4 (set officecn=SemiAnnual) & set cnkg=         & goto ed
if errorlevel  3 (set officecn=MonthlyEnterprise) & set cnkg=  & goto ed
if errorlevel  2 (set officecn=Current) & set cnkg=            & goto ed
if errorlevel  1 echo 已退出 & goto officemain






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
title  office 2016/2019/365安装向导
mode con cols=98 lines=30


echo:
echo:
echo                   _____________________________________________________________
echo                  ^|                                                             ^| 
echo                  ^|                                                             ^|
echo                  ^|      [0] 返回                                               ^|
echo                  ^|                                                             ^|
echo                  ^|      [1] 安装32位 %officever% ( %officecn% )%verkg%%cnkg%  ^|
echo                  ^|                                                             ^|
echo                  ^|      [2] 安装64位 %officever% ( %officecn% )%verkg%%cnkg%  ^|
echo                  ^|                                                             ^|
echo                  ^|_____________________________________________________________^|
echo:          
choice /C:012 /N /M ">                         请在键盘中输入你的选择 [0、1、2] ："
if errorlevel  3 (set officeed=64) & goto %officever%
if errorlevel  2 (set officeed=32) & goto %officever%
if errorlevel  1 goto cn





:ed32main

cls
title  office 2016/2019/365安装向导
mode con cols=98 lines=30


echo:
echo:
echo                   _____________________________________________________________
echo                  ^|                                                             ^| 
echo                  ^|                                                             ^|
echo                  ^|      [0] 返回                                               ^|
echo                  ^|                                                             ^|
echo                  ^|      [1] 安装32位 %officever% ( %officecn% )%verkg%%cnkg%  ^|
echo                  ^|                                                             ^|
echo                  ^|_____________________________________________________________^|
echo:          
choice /C:012 /N /M ">                         请在键盘中输入你的选择 [0、1、2] ："
if errorlevel  2 (set officeed=32) & goto %officever%
if errorlevel  1 goto cn




:office2016mini
echo ^<Configuration^>>officetemp.xml
echo   ^<Add OfficeClientEdition="%officeed%" Channel="%officecn%" SourcePath="%officedir%:"^>>>officetemp.xml
echo     ^<Product ID="ProPlusRetail"^>>>officetemp.xml
echo       ^<Language ID="zh-cn" /^>>>officetemp.xml
echo       ^<ExcludeApp ID="Access" /^>>>officetemp.xml
echo       ^<ExcludeApp ID="Groove" /^>>>officetemp.xml
echo       ^<ExcludeApp ID="Lync" /^>>>officetemp.xml
echo       ^<ExcludeApp ID="OneDrive" /^>>>officetemp.xml
echo       ^<ExcludeApp ID="OneNote" /^>>>officetemp.xml
echo       ^<ExcludeApp ID="Outlook" /^>>>officetemp.xml
echo       ^<ExcludeApp ID="Publisher" /^>>>officetemp.xml
echo     ^</Product^>>>officetemp.xml
echo   ^</Add^>>>officetemp.xml
echo ^</Configuration^>>>officetemp.xml
goto officesetup
:office2019mini
echo ^<Configuration^>>officetemp.xml
echo   ^<Add OfficeClientEdition="%officeed%" Channel="%officecn%" SourcePath="%officedir%:"^>>>officetemp.xml
echo     ^<Product ID="ProPlus2019Retail"^>>>officetemp.xml
echo       ^<Language ID="zh-cn" /^>>>officetemp.xml
echo       ^<ExcludeApp ID="Access" /^>>>officetemp.xml
echo       ^<ExcludeApp ID="Groove" /^>>>officetemp.xml
echo       ^<ExcludeApp ID="Lync" /^>>>officetemp.xml
echo       ^<ExcludeApp ID="OneDrive" /^>>>officetemp.xml
echo       ^<ExcludeApp ID="OneNote" /^>>>officetemp.xml
echo       ^<ExcludeApp ID="Outlook" /^>>>officetemp.xml
echo       ^<ExcludeApp ID="Publisher" /^>>>officetemp.xml
echo     ^</Product^>>>officetemp.xml
echo   ^</Add^>>>officetemp.xml
echo ^</Configuration^>>>officetemp.xml
goto officesetup
:office365mini
echo ^<Configuration^>>officetemp.xml
echo   ^<Add OfficeClientEdition="%officeed%" Channel="%officecn%" SourcePath="%officedir%:"^>>>officetemp.xml
echo     ^<Product ID="MondoRetail"^>>>officetemp.xml
echo       ^<Language ID="zh-cn" /^>>>officetemp.xml
echo       ^<ExcludeApp ID="Access" /^>>>officetemp.xml
echo       ^<ExcludeApp ID="Groove" /^>>>officetemp.xml
echo       ^<ExcludeApp ID="Lync" /^>>>officetemp.xml
echo       ^<ExcludeApp ID="OneDrive" /^>>>officetemp.xml
echo       ^<ExcludeApp ID="OneNote" /^>>>officetemp.xml
echo       ^<ExcludeApp ID="Outlook" /^>>>officetemp.xml
echo       ^<ExcludeApp ID="Project" /^>>>officetemp.xml
echo       ^<ExcludeApp ID="Publisher" /^>>>officetemp.xml
echo       ^<ExcludeApp ID="Visio" /^>>>officetemp.xml
echo     ^</Product^>>>officetemp.xml
echo   ^</Add^>>>officetemp.xml
echo ^</Configuration^>>>officetemp.xml
goto officesetup
:office365all
echo ^<Configuration^>>officetemp.xml
echo   ^<Add OfficeClientEdition="%officeed%" Channel="%officecn%" SourcePath="%officedir%:"^>>>officetemp.xml
echo     ^<Product ID="MondoRetail"^>>>officetemp.xml
echo       ^<Language ID="zh-cn" /^>>>officetemp.xml
echo       ^<ExcludeApp ID="Groove" /^>>>officetemp.xml
echo       ^<ExcludeApp ID="Lync" /^>>>officetemp.xml
echo       ^<ExcludeApp ID="OneDrive" /^>>>officetemp.xml
echo     ^</Product^>>>officetemp.xml
echo   ^</Add^>>>officetemp.xml
echo ^</Configuration^>>>officetemp.xml
goto officesetup

:officesetup
setup.exe /configure officetemp.xml
if %errorlevel%==0 (
echo 已安装成功
) else (
echo 安装失败
)
del /f /q officetemp.xml >nul 2>nul
goto end
:officeno
echo 未检测到office镜像，请先挂载office镜像。
goto end


:end
pause
