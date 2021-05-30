@ECHO OFF
SET Self=%~n0
REM UNCOMMENT BELOW FOR DEBUG
::IF NOT DEFINED in_subprocess (CMD /K SET in_subprocess=y ^& %0 %*) & EXIT )
REM Set title, size, prompt and version ;)
TITLE %Self%
MODE CON: COLS=95 LINES=30
SET SPGprompt=SPG ^$
SET Ver=v2.01

:MAIN
COLOR
CALL:HEADER
ECHO.
ECHO.                          Press CTRL+C to Abort and Terminate Batch Job
ECHO.
SET RNDLength=0
ECHO  +----------------------------------------+ MAIN MENU +--------------------------------------+
ECHO  :                                                                                           :
ECHO  :                                [1] 10 Random Characters Long                              :
ECHO  :                                [2] 20 Random Characters Long                              :
ECHO  :                                [3] 30 Random Characters Long                              :
ECHO  :                                [4] 40 Random Characters Long                              :
ECHO  :                                [5] 50 Random Characters Long                              :
ECHO  :                                [6] Custom Random Length                                   :
ECHO  :                                                                                           :
ECHO  :                                                                                           :
ECHO  :                                [7] EXIT                                                   :
ECHO  :                                                                                           :
ECHO  +-------------------------------------------------------------------------------------------+
ECHO.
CHOICE /C 1234567 /M "%SPGprompt%: Make your selection:"
IF %ERRORLEVEL%==1 SET RNDLength=10
IF %ERRORLEVEL%==2 SET RNDLength=20
IF %ERRORLEVEL%==3 SET RNDLength=30
IF %ERRORLEVEL%==4 SET RNDLength=40
IF %ERRORLEVEL%==5 SET RNDLength=50
IF %ERRORLEVEL%==6 GOTO :CUSTOM
IF %ERRORLEVEL%==7 GOTO :EXIT

:PGEN
REM The random character engine was provided by TheOucaste.
REM Some modifications were made to suit the Simple Password Generator.
REM A thank you to TheOutcaste for the code snippet.
SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
SET Alphanumeric=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%&)(-_=+][}{\|'";:/?.,
SET Str=%Alphanumeric%9876543210
:LENGHTLOOP
IF NOT "%Str:~18%"=="" SET Str=%Str:~9%& SET /A Len+=9& GOTO :LENGHTLOOP
SET tmp=%Str:~9,1%
SET /A Len=Len+tmp
SET count=0
SET "RNDAlphaNum="
:LOOP
SET /A count+=1
SET RND=%Random%
SET /A RND=RND%%%Len%
SET RNDAlphaNum=!RNDAlphaNum!!Alphanumeric:~%RND%,1!
IF !count! lss %RNDLength% GOTO :LOOP
SET "PassWord="
ECHO !RNDAlphaNum!| CLIP
ECHO %SPGprompt%: INFO: You selected a password %RNDLength% characters long.
ECHO %SPGprompt%: INFO: Copied !RNDAlphaNum! to the clipboard^^!
ECHO  +-------------------------------------------------------------------------------------------+
ECHO  :          WARNING: THE CLIPBOARD WILL BE CLEARED ON PASSWORD REGENERATION OR EXIT.         :
ECHO  +-------------------------------------------------------------------------------------------+
SET PassWord=!RNDAlphaNum!
ENDLOCAL && SET PassWord=%PassWord%

:QUERYSAVEPASS
CHOICE /M "%SPGprompt%: Would you like to save the password to file"
IF %ERRORLEVEL%==1 CALL :SAVEPASSNAME
IF %ERRORLEVEL%==2 GOTO :REGEN

:REGEN
CHOICE /M "%SPGprompt%: Would you like to generate a new password"
IF %ERRORLEVEL%==1 ECHO OFF | CLIP && CLS && GOTO :MAIN
IF %ERRORLEVEL%==2 GOTO :EXIT

:QUERY
CHOICE /M "%SPGprompt%: Return to Main Menu"
IF %ERRORLEVEL%==1 CLS && GOTO :MAIN
IF %ERRORLEVEL%==2 GOTO :EXIT

:EXIT
ECHO OFF | CLIP
COLOR 0E
CALL :EXITH
ECHO %KthX%
ECHO %ExiT% . . .
PING 1.0.0.0 -n 1 -w 100 >NUL
CALL :EXITH
ECHO %KthX%
ECHO %ExiT% .
PING 1.0.0.0 -n 1 -w 100 >NUL
CALL :EXITH
ECHO %KthX%
ECHO %ExiT% . .
PING 1.0.0.0 -n 1 -w 100 >NUL
CALL :EXITH
ECHO %KthX%
ECHO %ExiT% . . .
PING 1.0.0.0 -n 1 -w 100 >NUL
EXIT

:CUSTOM
SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
SET "CustomLenQ="""
SET /P CustomLenQ="SPG $: Enter your custom password length: "
SET /A EvalCLen=CustomLenQ
IF %EvalCLen% EQU %CustomLenQ% (
	IF %CustomLenQ% GTR 999 ( GOTO :INVALID )
	IF %CustomLenQ% GTR 9 ( SET RNDLength=%CustomLenQ% & CALL:PGEN )
	IF %CustomLenQ% LSS 9 ( GOTO :INVALID )
	IF %CustomLenQ% EQU 9 ( GOTO :INVALID )
) ELSE ( GOTO :INVALID )
ENDLOCAL

:INVALID
ECHO SPG $: ERROR: Input is invalid.
ECHO SPG $: INFO : Enter a number greater than 9 and smaller than 999.
GOTO :CUSTOM

:SAVEPASSNAME
SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
SET "SavePassName="
SET /P SavePassName="SPG $: Enter your a reference for this password: "
IF "%SavePassName%"=="" (
	GOTO :SAVEPASSNAMEINVALID
) ELSE (
	ECHO SPG $: INFO: You reference for this password is [ %SavePassName% ].
	SET Reference=%SavePassName% && GOTO :PROCESSTOFILE
)
ENDLOCAL

:SAVEPASSNAMEINVALID
ECHO SPG $: ERROR: Enter a valid reference name or title for your password.
ECHO SPG $: INFO : This will help you remember where this password belongs.
GOTO :SAVEPASSNAME

:PROCESSTOFILE
SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
SET FileName=Simple-Password-Generator-List.txt
IF EXIST %FileName% (
	ECHO SPG $: INFO: %FileName% exists, appending data to file.
	ECHO.>>%FileName%
	ECHO Reference: %Reference%>>%FileName%
	ECHO.>>%FileName%
	ECHO Password: %PassWord%>>%FileName%
	ECHO.>>%FileName%
	CALL :TIMESTAMP
	ECHO.>>%FileName%
	ECHO  +-------+ SAVED SECURE PASSWORD LIST +--------------------------------------+ SPG %Ver% +---+>>%FileName%
) ELSE (
	ECHO SPG $: INFO: %FileName% Not found, generating new file.
	CALL :HEADER >%FileName%
	ECHO.>>%FileName%
	ECHO Reference: %Reference%>>%FileName%
	ECHO.>>%FileName%
	ECHO Password: %PassWord%>>%FileName%
	ECHO.>>%FileName%
	CALL :TIMESTAMP
	ECHO.>>%FileName%
	ECHO  +-------+ SAVED SECURE PASSWORD LIST +--------------------------------------+ SPG %Ver% +---+>>%FileName%
)
ECHO SPG $: INFO: Your password and reference have been saved to %FileName%.
ECHO SPG $:  TIP: Keep this file safe in a secure external drive.
GOTO :REGEN

:TIMESTAMP
SET TimeStamp=%DATE% -%TIME%
ECHO Generated: %TimeStamp%>>%FileName%
GOTO :EOF

:HEADER
ECHO.
ECHO  +-------------------------------------------------------------------------------------------+
ECHO  :                           * * *  SIMPLE PASSWORD GENERATOR  * * *                         :
ECHO  +-------------------------------------------------------------------------------+ %Ver% +---+
GOTO :EOF

:EXITH
SET KthX=%SPGprompt%: Thank you for using %Self% :)
SET ExiT=%SPGprompt%: Bye
CLS
CALL :HEADER
ECHO.

ENDLOCAL