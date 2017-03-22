@echo off

REM Title:			Guided 'all_qa.txt' File Aggregator
REM Description:	This file recursively searches for all instances of 'all_qa.txt'
REM 				within a chosen input directory, then aggregates and outputs 
REM					them to a chosen output directory as a single text file.
REM Author:			Drewry Pope
REM Date:			09/12/2016
REM Version:		1.0.0.1

setlocal enabledelayedexpansion
set INPUT=%CD%
set OUTPUT=%CD%
:start
set END="Y"
if exist "all_qa_dirpath.txt" erase "all_qa_dirpath.txt"
if exist "all_qa_filepath.txt" erase "all_qa_filepath.txt"
echo What is the file path you want to search within? (Input Directory)
echo Current Default^: %INPUT%
echo Don't include / at end.
set /p INPUT=""
dir "%INPUT%" /s /b /ad > all_qa_dirpath.txt
echo. 2>all_qa_filepath.txt
for /f "delims=" %%i in (all_qa_dirpath.txt) do (
if exist "%%i\all_qa.txt" (echo "%%i\all_qa.txt" >> all_qa_filepath.txt)
)
for /f %%a in ('type "all_qa_filepath.txt" ^| find "" /v /c') do set /a COUNT=%%a
:redo
echo,
echo We found %COUNT% all_qa.txt files in that directory.
set CONTINUE="Y"
echo Would you like to aggregate them? (Default^: Yes)
echo (Y - Yes, N - No, L - List, R- Restart)
set /p CONTINUE=""
echo,
if "%CONTINUE%"=="l" goto list
if "%CONTINUE%"=="L" goto list
if "%CONTINUE%"=="n" goto skip
if "%CONTINUE%"=="N" goto skip
if "%CONTINUE%"=="no" goto skip
if "%CONTINUE%"=="No" goto skip
if "%CONTINUE%"=="NO" goto skip
if "%CONTINUE%"=="r" goto start
if "%CONTINUE%"=="R" goto start
echo Where would you like to aggregate them? (Output Directory)
echo Current Default^: %OUTPUT%
echo Don't include \ at end.
set /p OUTPUT=""
echo,
echo Aggregating...
for /f "delims=" %%i in (all_qa_filepath.txt) do type %%i >>"%OUTPUT%\all_qa_aggregated_%date:~4,2%%date:~7,2%%date:~-2,2%.txt"
echo Operation Completed.
echo,
:skip
if exist "all_qa_dirpath.txt" erase "all_qa_dirpath.txt"
if exist "all_qa_filepath.txt" erase "all_qa_filepath.txt"
echo E to Soft Exit. (if started from prompt)
echo R to Restart.
echo S to Split File.
echo Press Enter to Close Window. 
set /p END=""
if "%END%"=="r" goto start
if "%END%"=="R" goto start
if "%END%"=="e" exit /b
if "%END%"=="E" exit /b
if "%END%"=="s" goto split
if "%END%"=="S" goto split
exit

:list
for /f %%a in ('type "all_qa_filepath.txt" ^| find "" /v /c') do set /a COUNT=%%a
echo The %COUNT% sources of "all_qa.txt" within "%INPUT%" ^:
echo,
type all_qa_filepath.txt
echo,
goto redo

:split
set SPLIT_SIZE=150
echo How many MB would you like each split to be? (# only, don't include MB)
set /p SPLIT_SIZE=""
@REM Using http://stackoverflow.com/questions/19335004/how-to-run-a-powershell-script-from-a-batch-file
@REM and http://stackoverflow.com/questions/1001776/how-can-i-split-a-text-file-using-powershell
@PowerShell  ^
    $upperBound = %SPLIT_SIZE%MB;  ^
    $rootName = '%OUTPUT%\all_qa_aggregated_%date:~4,2%%date:~7,2%%date:~-2,2%';  ^
    $from = $rootName;  ^
	$ext = 'txt';  ^
	$from1 = '{0}.{1}' -f ($from, $ext);  ^
    $fromFile = [io.file]::OpenRead($from1);  ^
    $buff = new-object byte[] $upperBound;  ^
    $count = $idx = 0;  ^
    try {  ^
        do {  ^
            'Reading ' + $upperBound;  ^
            $count = $fromFile.Read($buff, 0, $buff.Length);  ^
            if ($count -gt 0) {  ^
                $to = '{0}.{1}.{2}' -f ($rootName, $idx, $ext);  ^
                $toFile = [io.file]::OpenWrite($to);  ^
                try {  ^
                    'Writing ' + $count + ' to ' + $to;  ^
                    $tofile.Write($buff, 0, $count);  ^
                } finally {  ^
                    $tofile.Close();  ^
                }  ^
            }  ^
            $idx ++;  ^
        } while ($count -gt 0);  ^
    }  ^
    finally {  ^
        $fromFile.Close();  ^
    }  ^
%End PowerShell%
echo Split Completed.
echo Press Enter to Close Window. 
set END="y"
set /p END=""
exit