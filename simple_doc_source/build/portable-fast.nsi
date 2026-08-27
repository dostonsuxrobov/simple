!include "common.nsh"
!include "extractAppPackage.nsh"

# Keep the portable executable self-contained, but retain a versioned extracted
# runtime after its first launch. Electron Builder's stock portable launcher
# deletes and re-extracts Chromium on every run, which dominates startup time.

CRCCheck off
WindowIcon Off
AutoCloseWindow True
RequestExecutionLevel ${REQUEST_EXECUTION_LEVEL}

Var cacheDir

Function .onInit
  !ifndef SPLASH_IMAGE
    SetSilent silent
  !endif

  !insertmacro check64BitAndSetRegView
FunctionEnd

Function .onGUIInit
  InitPluginsDir

  !ifdef SPLASH_IMAGE
    File /oname=$PLUGINSDIR\splash.bmp "${SPLASH_IMAGE}"
    BgImage::SetBg $PLUGINSDIR\splash.bmp
    BgImage::Redraw
  !endif
FunctionEnd

Section
  !ifdef SPLASH_IMAGE
    HideWindow
  !endif

  !ifdef UNPACK_DIR_NAME
    StrCpy $cacheDir "$LOCALAPPDATA\Simple Docs\cache\${UNPACK_DIR_NAME}"
    IfFileExists "$cacheDir\.simple-docs-runtime-ready" 0 extractRuntime
    IfFileExists "$cacheDir\${APP_EXECUTABLE_FILENAME}" 0 extractRuntime
    IfFileExists "$cacheDir\resources\app.asar" useCachedRuntime extractRuntime
  !endif

extractRuntime:
  !ifdef UNPACK_DIR_NAME
    RMDir /r "$cacheDir"
  !endif

  # $PLUGINSDIR is unique to this launcher process, so extraction is isolated.
  # The complete staging directory is renamed into the cache only after its
  # ready marker has been written.
  StrCpy $INSTDIR "$PLUGINSDIR\app"
  SetOutPath $INSTDIR

  !ifdef APP_DIR_64
    !ifdef APP_DIR_ARM64
      !ifdef APP_DIR_32
        ${if} ${IsNativeARM64}
          File /r "${APP_DIR_ARM64}\*.*"
        ${elseif} ${RunningX64}
          File /r "${APP_DIR_64}\*.*"
        ${else}
          File /r "${APP_DIR_32}\*.*"
        ${endIf}
      !else
        ${if} ${IsNativeARM64}
          File /r "${APP_DIR_ARM64}\*.*"
        ${else}
          File /r "${APP_DIR_64}\*.*"
        ${endIf}
      !endif
    !else
      !ifdef APP_DIR_32
        ${if} ${RunningX64}
          File /r "${APP_DIR_64}\*.*"
        ${else}
          File /r "${APP_DIR_32}\*.*"
        ${endIf}
      !else
        File /r "${APP_DIR_64}\*.*"
      !endif
    !endif
  !else
    !ifdef APP_DIR_32
      File /r "${APP_DIR_32}\*.*"
    !else
      !insertmacro extractEmbeddedAppPackage
    !endif
  !endif

  !ifdef UNPACK_DIR_NAME
    FileOpen $R9 "$INSTDIR\.simple-docs-runtime-ready" w
    FileWrite $R9 "${UNPACK_DIR_NAME}"
    FileClose $R9
    CreateDirectory "$LOCALAPPDATA\Simple Docs"
    CreateDirectory "$LOCALAPPDATA\Simple Docs\cache"
    SetOutPath $PLUGINSDIR
    Rename "$INSTDIR" "$cacheDir"
    IfErrors cacheRaceWonByOtherProcess cacheCommitComplete

cacheRaceWonByOtherProcess:
    # A simultaneous first launch may have populated the same build cache.
    # Use it only after its complete marker and core files are visible.
    IfFileExists "$cacheDir\.simple-docs-runtime-ready" 0 useStagedRuntime
    IfFileExists "$cacheDir\${APP_EXECUTABLE_FILENAME}" 0 useStagedRuntime
    IfFileExists "$cacheDir\resources\app.asar" useCachedRuntime useStagedRuntime

cacheCommitComplete:
    StrCpy $INSTDIR "$cacheDir"
    Goto launchRuntime
  !endif

useStagedRuntime:
  # Cross-volume TEMP configurations can prevent an atomic directory rename.
  # Launch the complete staging copy; NSIS will clean it after this process.
  Goto launchRuntime

useCachedRuntime:
  StrCpy $INSTDIR "$cacheDir"

launchRuntime:
  System::Call 'Kernel32::SetEnvironmentVariable(t, t)i ("PORTABLE_EXECUTABLE_DIR", "$EXEDIR").r0'
  System::Call 'Kernel32::SetEnvironmentVariable(t, t)i ("PORTABLE_EXECUTABLE_FILE", "$EXEPATH").r0'
  System::Call 'Kernel32::SetEnvironmentVariable(t, t)i ("PORTABLE_EXECUTABLE_APP_FILENAME", "${APP_FILENAME}").r0'
  ${StdUtils.GetAllParameters} $R0 0

  !ifdef SPLASH_IMAGE
    BgImage::Destroy
  !endif

  ExecWait '"$INSTDIR\${APP_EXECUTABLE_FILENAME}" $R0' $0
  SetErrorLevel $0
  SetOutPath $EXEDIR
SectionEnd
