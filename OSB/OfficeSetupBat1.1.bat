@echo off

%1 mshta vbscript:CreateObject("Shell.Application").ShellExecute("cmd.exe","/c %~s0 ::","","runas",1)(window.close)&&exit


cd /d %~dp0


for /f "tokens=3" %%i in ('echo list volume ^| diskpart ^| findstr "16.0"') do (
    if exist "%%i:\Office" set officedir=%%i&goto officemain
    
)
goto officeno


:officemain
cls
title  office 2016/2019/365��װ��
mode con cols=98 lines=30


echo:
echo:
echo                   _____________________________________________________________
echo                  ^|                                                             ^| 
echo                  ^|                                                             ^|
echo                  ^|      [0] �˳�                                               ^|
echo                  ^|                                                             ^|
echo                  ^|      [1] ��װ office2016��word��excle��powerpoint��         ^|
echo                  ^|                                                             ^|
echo                  ^|      [2] ��װ office2019��word��excle��powerpoint��         ^|
echo                  ^|                                                             ^|
echo                  ^|      [3] ��װ office365��word��excle��powerpoint��          ^|
echo                  ^|                                                             ^|
echo                  ^|      [4] ��װ office365��ALL��                              ^|
echo                  ^|                                                             ^|
echo                  ^|_____________________________________________________________^|
echo:          
choice /C:01234 /N /M ">                         ���ڼ������������ѡ�� [0��1��2��3��4] ��"
if errorlevel  5 (set officever=office365all) & set verkg=    & goto cn
if errorlevel  4 (set officever=office365mini) & set verkg=   & goto cn
if errorlevel  3 (set officever=office2019mini) & set verkg=  & goto cn
if errorlevel  2 (set officever=office2016mini) & set verkg=  & goto cn
if errorlevel  1 goto end


:cn
cls
title  office 2016/2019/365��װ��
mode con cols=98 lines=30


echo:
echo:
echo                   _____________________________________________________________
echo                  ^|                                                             ^| 
echo                  ^|                                                             ^|
echo                  ^|      [0] ����                                               ^|
echo                  ^|                                                             ^|
echo                  ^|      [1] ѡ�� %officever% ����ͨ������ǰͨ�� %verkg%          ^|
echo                  ^|                                                             ^|
echo                  ^|      [2] ѡ�� %officever% ����ͨ�����¶���ҵͨ�� %verkg%      ^|
echo                  ^|                                                             ^|
echo                  ^|      [3] ѡ�� %officever% ����ͨ�����������ҵͨ�� %verkg%    ^|
echo                  ^|                                                             ^|
echo                  ^|                                                             ^|
echo                  ^|_____________________________________________________________^|
echo:          
choice /C:01234 /N /M ">                         ���ڼ������������ѡ�� [0��1��2��3��4] ��"
if errorlevel  4 (set officecn=SemiAnnual) & set cnkg=         & goto ed
if errorlevel  3 (set officecn=MonthlyEnterprise) & set cnkg=  & goto ed
if errorlevel  2 (set officecn=Current) & set cnkg=            & goto ed
if errorlevel  1 goto officemain






:ed
if "%PROCESSOR_ARCHITEW6432%"=="AMD64" (
    goto ed64main
) else if "%PROCESSOR_ARCHITECTURE%"=="AMD64" (
    goto ed64main
) else if "%PROCESSOR_ARCHITECTURE%"=="x86" (
    goto ed32main
) else (
    echo ��֧�ֵ�ǰ�Ĳ���ϵͳ
    goto end
)

:ed64main

cls
title  office 2016/2019/365��װ��
mode con cols=98 lines=30


echo:
echo:
echo                   _____________________________________________________________
echo                  ^|                                                             ^| 
echo                  ^|                                                             ^|
echo                  ^|      [0] ����                                               ^|
echo                  ^|                                                             ^|
echo                  ^|      [1] ��װ32λ %officever% ( %officecn% )%verkg%%cnkg%  ^|
echo                  ^|                                                             ^|
echo                  ^|      [2] ��װ64λ %officever% ( %officecn% )%verkg%%cnkg%  ^|
echo                  ^|                                                             ^|
echo                  ^|_____________________________________________________________^|
echo:          
choice /C:012 /N /M ">                         ���ڼ������������ѡ�� [0��1��2] ��"
if errorlevel  3 (set officeed=64) & goto %officever%
if errorlevel  2 (set officeed=32) & goto %officever%
if errorlevel  1 goto cn





:ed32main

cls
title  office 2016/2019/365��װ��
mode con cols=98 lines=30


echo:
echo:
echo                   _____________________________________________________________
echo                  ^|                                                             ^| 
echo                  ^|                                                             ^|
echo                  ^|      [0] ����                                               ^|
echo                  ^|                                                             ^|
echo                  ^|      [1] ��װ32λ %officever% ( %officecn% )%verkg%%cnkg%  ^|
echo                  ^|                                                             ^|
echo                  ^|_____________________________________________________________^|
echo:          
choice /C:012 /N /M ">                         ���ڼ������������ѡ�� [0��1��2] ��"
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
    del /f /q officetemp.xml >nul 2>nul
    echo ��װ�����
    pause
    exit
) else (
    echo ��װʧ��,errorlevel:%errorlevel%
    goto end
)

:officeno
echo δ��⵽office�������ȹ���office����
goto end


:end
echo ��װ���˳�
pause