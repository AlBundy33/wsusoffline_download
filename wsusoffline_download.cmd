@ECHO OFF
SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

:: =========================================================================================================================
:: user settings

:: 0 - append to log, 1 - override log, 2 - log-rotate
SET LOG_HANDLING=2
:: =========================================================================================================================

:: =========================================================================================================================
:: internal settings
SET WSUS_DIR=%~dp0
SET UPDATE_GENERATOR_INI=!WSUS_DIR!\UpdateGenerator.ini
SET DOWNLOAD_UPDATES_CMD=!WSUS_DIR!\cmd\DownloadUpdates.cmd
SET CREATE_ISO_IMAGE_CMD=!WSUS_DIR!\cmd\CreateISOImage.cmd
SET COPY_TO_TARGET_CMD=!WSUS_DIR!\cmd\CopyToTarget.cmd
SET LOG_DIR=!WSUS_DIR!\log
SET DOWNLOAD_LOG=!LOG_DIR!\download.log

SET LANGUAGES=enu fra esn jpn kor rus ptg ptb deu nld ita chs cht plk hun csy sve trk ell ara heb dan nor fin
SET ALL=all all-x86 all-x64
SET WIN_APPS=w2k3 w2k3-x64
SET OFC_APPS=o2k7 o2k7-x64 o2k10
SET GLOBAL_WIN_APPS=w60 w60-x64 w61 w61-x64 w62 w62-x64 w63 w63-x64
SET GLOBAL_OFC_APPS=ofc
SET APPLICATIONS=!WIN_APPS! !OFC_APPS!
SET GLOBAL_APPLICATIONS=!GLOBAL_WIN_APPS! !GLOBAL_OFC_APPS!
:: special handling for scripts other than DownloadUpdates.cmd
SET SCRIPT_APPLICATIONS=!WIN_APPS! ofc
SET SCRIPT_GLOBAL_APPLICATIONS=!GLOBAL_WIN_APPS!
:: =========================================================================================================================

ECHO.%~nx0 by AlBundy
ECHO.more information at: https://github.com/AlBundy33/wsusoffline_download
ECHO.

IF NOT EXIST "!WSUS_DIR!\UpdateGenerator.exe" (
	CALL :LOG_ERROR please copy this script into your wsusoffline-folder
	GOTO :EOF
)

IF NOT EXIST "!CREATE_ISO_IMAGE_CMD!" (
	CALL :LOG_ERROR "!CREATE_ISO_IMAGE_CMD!" not found.
	GOTO :EOF
)

IF NOT EXIST "!DOWNLOAD_UPDATES_CMD!" (
	CALL :LOG_ERROR "!DOWNLOAD_UPDATES_CMD!" not found.
	GOTO :EOF
)

IF NOT EXIST "!UPDATE_GENERATOR_INI!" (
	CALL :LOG_ERROR "!UPDATE_GENERATOR_INI!" not found.
	GOTO :EOF
)

IF NOT EXIST "!COPY_TO_TARGET_CMD!" (
	CALL :LOG_ERROR "!COPY_TO_TARGET_CMD!" not found.
	GOTO :EOF
)

IF EXIST "!DOWNLOAD_LOG!" (
	IF "!LOG_HANDLING!"=="1" (
		DEL /F /Q "!DOWNLOAD_LOG!"
	) ELSE IF "!LOG_HANDLING!"=="2" (
		PUSHD "!LOG_DIR!"
		SET /A LOG_COUNT=0
		FOR %%L IN (*.log) DO (
			SET /A LOG_COUNT+=1
		)
		REN "!DOWNLOAD_LOG!" "download[!LOG_COUNT!].log"
		POPD
	)
)

IF EXIST "!DOWNLOAD_LOG!" (
	CALL :LOG_SEPARATOR
	CALL :LOG_INFO starting %~nx0
	CALL :LOG_SEPARATOR
)

SET SHUTDOWN=0
SET RESTART=0
SET DEBUG=0

:: parse arguemnts
:PARSE_ARGUMENTS
	SET ARG=%~1
	SET SUBARGS=
	:PARSE_SUB_ARGS
		SET TEST=%~2
		IF "!TEST!"=="" GOTO :NOSUBARGS
		IF "!TEST:~0,2!"=="--" GOTO :NOSUBARGS
		SET SUBARGS=!SUBARGS! %2
		SHIFT /1
		GOTO :PARSE_SUB_ARGS
	:NOSUBARGS
	IF NOT "!ARG!"=="" (
		IF /I "!ARG!"=="--shutdown" (
			SET SHUTDOWN=1
		) ELSE IF /I "!ARG!"=="--restart" (
			SET RESTART=1
		) ELSE IF /I "!ARG!"=="--skipdownload" (
			SET SKIPDOWNLOAD=1
		) ELSE IF /I "!ARG!"=="--debug" (
			SET DEBUG=1
		) ELSE IF /I "!ARG!"=="--help" (
			CALL :SHOW_USAGE
			GOTO :EOF
		) ELSE IF /I "!ARG!"=="--ini" (
			SHIFT /1
			SET "UPDATE_GENERATOR_INI=%~1"
		) ELSE (
			CALL :SHOW_USAGE unknown argument: "!ARG!"
			GOTO :EOF
		)
		SHIFT /1
		GOTO :PARSE_ARGUMENTS
	)

:: read settings from ini
CALL :PARSE_INI

:: init variables
CALL :CHECK_SETTINGS
CALL :INIT
CALL :WORD_COUNT "!DOWNLOAD_LOG!" "Error" ERROR_COUNT_START
CALL :WORD_COUNT "!DOWNLOAD_LOG!" "Warning" WARNING_COUNT_START
IF NOT "!SKIPDOWNLOAD!"=="1" CALL :DOWNLOAD
CALL :CREATE_ISO
CALL :COPY_TO_TARGET
CALL :WORD_COUNT "!DOWNLOAD_LOG!" "Error" ERROR_COUNT_END
CALL :WORD_COUNT "!DOWNLOAD_LOG!" "Warning" WARNING_COUNT_END

SET /A ERROR_COUNT=!ERROR_COUNT_END! - !ERROR_COUNT_START!
SET /A WARNING_COUNT=!WARNING_COUNT_END! - !WARNING_COUNT_START!

CALL :LOG_SEPARATOR
CALL :LOG_INFO log contains !ERROR_COUNT! new error^(s^).
CALL :LOG_INFO log contains !WARNING_COUNT! new warning^(s^).
CALL :LOG_SEPARATOR

IF "!SHUTDOWN!"=="1" (
	CALL :RUN shutdown -s -f -t 30 -c "%~nx0 finished - shutting system down..."
) ELSE IF "!RESTART!"=="1" (
	CALL :RUN shutdown -r -f -t 30 -c "%~nx0 finished - restarting system..."
)

GOTO :EOF

:PARSE_INI
	:: see UpdateGenerator.au3 for ini_section mappings
	SET SECTION=
	SET ENABLED_LANGUAGES=
	FOR /F "tokens=*" %%L IN (!UPDATE_GENERATOR_INI!) DO (
		SET LINE=%%~L
		IF "!LINE:~0,1!"=="[" (
			IF "!LINE:~-1!"=="]" (
				IF /I "!LINE!"=="[Windows 2000]" (
					SET SECTION=w2k
				) ELSE IF /I "!LINE!"=="[Windows XP]" (
					SET SECTION=wxp
				) ELSE IF /I "!LINE!"=="[Windows Vista]" (
					SET SECTION=w60
				) ELSE IF /I "!LINE!"=="[Windows Vista x64]" (
					SET SECTION=w60-x64
				) ELSE IF /I "!LINE!"=="[Windows Server 2003]" (
					SET SECTION=w2k3
				) ELSE IF /I "!LINE!"=="[Windows Server 2003 x64]" (
					SET SECTION=w2k3-x64
				) ELSE IF /I "!LINE!"=="[Windows 7]" (
					SET SECTION=w61
				) ELSE IF /I "!LINE!"=="[Windows Server 2008 R2]" (
					SET SECTION=w61-x64
				) ELSE IF /I "!LINE!"=="[Office 2000]" (
					SET SECTION=skip
				) ELSE IF /I "!LINE!"=="[Office XP]" (
					SET SECTION=oxp
				) ELSE IF /I "!LINE!"=="[Office 2003]" (
					SET SECTION=o2k3
				) ELSE IF /I "!LINE!"=="[Office 2007]" (
					SET SECTION=o2k7
				) ELSE IF /I "!LINE!"=="[Office 2007 x64]" (
					SET SECTION=o2k7-x64
				) ELSE IF /I "!LINE!"=="[Office]" (
					SET SECTION=ofc
				) ELSE IF /I "!LINE!"=="[Office 2010]" (
					SET SECTION=o2k10
				) ELSE IF /I "!LINE!"=="[ISO Images]" (
					SET SECTION=iso
				) ELSE IF /I "!LINE!"=="[Miscellaneous]" (
					SET SECTION=misc
				) ELSE IF /I "!LINE!"=="[Options]" (
					SET SECTION=options
				) ELSE IF /I "!LINE!"=="[USB Images]" (
					SET SECTION=usb
				) ELSE IF /I "!LINE!"=="[Windows 8]" (
					SET SECTION=w62
				) ELSE IF /I "!LINE!"=="[Windows Server 2012]" (
					SET SECTION=w62-x64
				) ELSE IF /I "!LINE!"=="[Windows 8.1]" (
					SET SECTION=w63
				) ELSE IF /I "!LINE!"=="[Windows Server 2012 R2]" (
					SET SECTION=w63-x64
				) ELSE IF /I "!LINE!"=="[Windows 10]" (
					SET SECTION=w100
				) ELSE IF /I "!LINE!"=="[Windows Server 2016]" (
					SET SECTION=w100_x64
				) ELSE IF /I "!LINE!"=="[Office 2013]" (
					SET SECTION=o2k13
				) ELSE IF /I "!LINE!"=="[Office 2016]" (
					SET SECTION=o2k16
				) ELSE (
					CALL :LOG_WARNING unknown section found: !LINE!
					SET SECTION=!LINE:~1,-1!
					SET SECTION=!SECTION: =_!
					REM GOTO :EOF
				)
				SET !SECTION!_description=!LINE:~1,-1!
			)
		) ELSE IF /I NOT "!SECTION!" == "skip" (
			FOR /F "tokens=1,* delims==" %%K IN ("!LINE!") DO (
				SET KEY=%%~K
				SET VALUE=%%~L
				IF NOT "!SECTION!"=="" (
					SET !SECTION!_!KEY!=!VALUE!
				) ELSE (
					SET !KEY!=!VALUE!
				)
				IF NOT "!LANGUAGES:%%K=!"=="!LANGUAGES!" (
					IF /I NOT "!KEY!"=="glb" (
						IF /I "!VALUE!"=="Enabled" (
							IF "!ENABLED_LANGUAGES!"=="" (
								SET ENABLED_LANGUAGES=!ENABLED_LANGUAGES! %%K
							) ELSE IF "!ENABLED_LANGUAGES:%%K=!"=="!ENABLED_LANGUAGES!" (
								SET ENABLED_LANGUAGES=!ENABLED_LANGUAGES! %%K
							)
						)
					)
				)
			)
		)
	)
	GOTO :EOF

:DOWNLOAD
	IF EXIST "!WSUS_DIR!\client\md\*.txt" (
		PUSHD "!WSUS_DIR!\client\md"
		FOR %%F IN (*.txt) DO (
			IF "%%~zF"=="0" (
				CALL :LOG_WARNING size of "%%~fF" is 0Byte - deleting file...
				DEL /Q /F "%%~fF"
			) ELSE (
				FINDSTR /M /C:"%%%%" /C:"##" "%%~fF" >NUL
				IF NOT "!ERRORLEVEL!"=="0" (
					CALL :LOG_WARNING "%%~fF" seems to be invalid - deleting file...
					DEL /Q /F "%%~fF"
				)
			)
		)
		POPD
	)
	FOR %%A IN (!APPLICATIONS!) DO (
		FOR %%L IN (!LANGUAGES!) DO (
			IF /I "!%%A_%%L!"=="Enabled" (
				CALL :LOG_SEPARATOR
				CALL :LOG_INFO application: !%%A_description! ^(%%A^), language: %%L
				CALL :RUN "!DOWNLOAD_UPDATES_CMD!" %%A %%L !DEFAULT_DOWNLOAD_ARGS!
			)
		)
	)
	FOR %%A IN (!GLOBAL_APPLICATIONS!) DO (
		IF /I "!%%A_glb!"=="Enabled" (
			CALL :LOG_SEPARATOR
			CALL :LOG_INFO application: !%%A_description! ^(%%A^), language: glb
			CALL :RUN "!DOWNLOAD_UPDATES_CMD!" %%A glb !DEFAULT_DOWNLOAD_ARGS!
		)
	)
	GOTO :EOF

:CHECK_SETTINGS
	FOR %%L IN (!LANGUAGES!) DO (
		FOR %%O IN (!OFC_APPS! !GLOBAL_OFC_APPS!) DO (
			IF /I "!%%O_%%L!"=="Enabled" (
				:: force glb if an office-package is enabled
				SET ofc_glb=Enabled
				:: enable ofc for current language
				SET ofc_%%L=Enabled
			)
		)
	)
	GOTO :EOF

:INIT
	SET DEFAULT_DOWNLOAD_ARGS=
	SET DEFAULT_ISO_ARGS=
	SET DEFAULT_COPY_ARGS=
	IF /I "!options_includesp!"=="Disabled" (
		SET DEFAULT_DOWNLOAD_ARGS=!DEFAULT_DOWNLOAD_ARGS! /excludesp
		SET DEFAULT_ISO_ARGS=!DEFAULT_ISO_ARGS! /excludesp
		SET DEFAULT_COPY_ARGS=!DEFAULT_COPY_ARGS! /excludesp
	)
	IF /I "!misc_excludestatics!"=="Enabled" SET DEFAULT_DOWNLOAD_ARGS=!DEFAULT_DOWNLOAD_ARGS! /excludestatics
	IF /I "!options_cleanupdownloads!"=="Disabled" SET DEFAULT_DOWNLOAD_ARGS=!DEFAULT_DOWNLOAD_ARGS! /nocleanup
	IF /I "!options_includedotnet!"=="Enabled" (
		SET DEFAULT_DOWNLOAD_ARGS=!DEFAULT_DOWNLOAD_ARGS! /includedotnet
		SET DEFAULT_ISO_ARGS=!DEFAULT_ISO_ARGS! /includedotnet
		SET DEFAULT_COPY_ARGS=!DEFAULT_COPY_ARGS! /includedotnet
	)
	IF /I "!options_includemsse!"=="Enabled" (
		SET DEFAULT_DOWNLOAD_ARGS=!DEFAULT_DOWNLOAD_ARGS! /includemsse
		SET DEFAULT_ISO_ARGS=!DEFAULT_ISO_ARGS! /includemsse
		SET DEFAULT_COPY_ARGS=!DEFAULT_COPY_ARGS! /includemsse
	)
	IF /I "!options_includewddefs!"=="Enabled" (
		SET DEFAULT_DOWNLOAD_ARGS=!DEFAULT_DOWNLOAD_ARGS! /includewddefs
		SET DEFAULT_ISO_ARGS=!DEFAULT_ISO_ARGS! /includewddefs
		REM SET DEFAULT_COPY_ARGS=!DEFAULT_COPY_ARGS! /includewddefs
	)
	IF /I "!options_verifydownloads!"=="Enabled" SET DEFAULT_DOWNLOAD_ARGS=!DEFAULT_DOWNLOAD_ARGS! /verify
	IF NOT "!misc_proxy!"=="" SET DEFAULT_DOWNLOAD_ARGS=!DEFAULT_DOWNLOAD_ARGS! /proxy !misc_proxy!
	IF NOT "!misc_wsus!"=="" SET DEFAULT_DOWNLOAD_ARGS=!DEFAULT_DOWNLOAD_ARGS! /wsus !misc_wsus!
	GOTO :EOF

:CREATE_ISO
	IF /I "!iso_single!"=="Enabled" (
		FOR %%A IN (!SCRIPT_APPLICATIONS!) DO (
			FOR %%L IN (!LANGUAGES!) DO (
				IF /I "!%%A_%%L!"=="Enabled" (
					CALL :LOG_SEPARATOR
					CALL :LOG_INFO createing iso for: !%%A_description! ^(%%L^)
					CALL :RUN "!CREATE_ISO_IMAGE_CMD!" %%A %%L !DEFAULT_ISO_ARGS!
				)
			)
		)
		FOR %%A IN (!SCRIPT_GLOBAL_APPLICATIONS!) DO (
			IF /I "!%%A_glb!"=="Enabled" (
				CALL :LOG_SEPARATOR
				CALL :LOG_INFO createing iso for: !%%A_description! ^(glb^)
				CALL :RUN "!CREATE_ISO_IMAGE_CMD!" %%A glb !DEFAULT_ISO_ARGS!
			)
		)
	)
	IF /I "!iso_cross-platform!"=="Enabled" (
		FOR %%L IN (!ENABLED_LANGUAGES! !ALL!) DO (
			CALL :LOG_SEPARATOR
			CALL :LOG_INFO createing iso for language: %%L
			CALL :RUN "!CREATE_ISO_IMAGE_CMD!" %%L !DEFAULT_ISO_ARGS!
		)
	)
	GOTO :EOF

:COPY_TO_TARGET
	IF /I "!usb_copy!"=="Enabled" (
		IF "!usb_path!"=="" (
			CALL :LOG_ERROR target-dir not defined.
			GOTO :EOF
		)
		IF NOT EXIST "!usb_path!" (
			CALL :LOG_ERROR target-dir ^(!usb_path!^) does not exist.
			GOTO :EOF
		)
		FOR %%A IN (!SCRIPT_APPLICATIONS!) DO (
			FOR %%L IN (!LANGUAGES!) DO (
				IF /I "!%%A_%%L!"=="Enabled" (
					CALL :LOG_SEPARATOR
					CALL :LOG_INFO copying files for: !%%A_description! ^(%%L^)
					CALL :RUN "!COPY_TO_TARGET_CMD!" %%A %%L "!usb_path!" !DEFAULT_COPY_ARGS!
				)
			)
		)
		FOR %%A IN (!SCRIPT_GLOBAL_APPLICATIONS!) DO (
			IF /I "!%%A_glb!"=="Enabled" (
				CALL :LOG_SEPARATOR
				CALL :LOG_INFO copying files for: !%%A_description! ^(glb^)
				CALL :RUN "!COPY_TO_TARGET_CMD!" %%A glb "!usb_path!" !DEFAULT_COPY_ARGS!
			)
		)
	)
	GOTO :EOF

:WORD_COUNT
	IF "%~1"=="" (
		CALL :LOG_ERROR no filename given.
		GOTO :EOF
	)
	IF "%~2"=="" (
		CALL :LOG_ERROR no searchstring given.
		GOTO :EOF
	)
	IF "%~3"=="" (
		CALL :LOG_ERROR no variablename given.
		GOTO :EOF
	)
	SET /A %~3=0
	IF EXIST "%~1" (
		FOR /F "tokens=*" %%L IN ('FINDSTR /R /I /C:"\<%~2\>" "%~1"') DO (
			SET /A %~3+=1
		)
	)
	GOTO :EOF

:SHOW_USAGE
	ECHO.usage: %~nx0 [arguments]
	ECHO.arguments:
	ECHO.  --help          shows this text
	ECHO.  --shutdown      shuts the system down after skript is finished
	ECHO.  --restart       restarts the system after skript is finished
	ECHO.  --skipdownload  skip download of updates
	ECHO.  --ini FILE      read settings from given FILE
	ECHO.  --debug         print commands without executing
	ECHO.
	IF NOT "%~1"=="" (
		ECHO.%*
		ECHO.
	)
	GOTO :EOF

:RUN
	:: reset errorlevel
	VER >NUL 2>NUL
	CALL :LOG_INFO command-line: %*
	IF "!DEBUG!"=="0" (
		CALL %*
		IF NOT "!ERRORLEVEL!"=="0" (
			CALL :LOG_ERROR command ended with exit-code: !ERRORLEVEL!
		)
	)
	GOTO :EOF

:LOG
	SET MSG=!DATE! !TIME: =0! %*
	ECHO.!MSG!
	ECHO>>"!DOWNLOAD_LOG!" !MSG!
	GOTO :EOF

:LOG_ERROR
	CALL :LOG [ERROR] %*
	GOTO :EOF

:LOG_INFO
	CALL :LOG [INFO] %*
	GOTO :EOF

:LOG_WARNING
	CALL :LOG [WARNING] %*
	GOTO :EOF

:LOG_SEPARATOR
	ECHO.>>"!DOWNLOAD_LOG!"
	ECHO.-------------------------------------------------------------------------------- >>"!DOWNLOAD_LOG!"
	ECHO.>>"!DOWNLOAD_LOG!"
	GOTO :EOF

:CONTAINS
	SET FINDSTR_ARGS=/R "%~2"
	IF "%~3"=="1" SET FINDSTR_ARGS=/I !FINDSTR_ARGS!
	ECHO "%~1" | FINDSTR !FINDSTR_ARGS! >NUL 2>NUL
	GOTO :EOF

:CONTAINS_WORD
	CALL :CONTAINS %1 "\<%~2\>" %~3
	GOTO :EOF
