taskkill /F /IM KeepAwake.exe >nul 2>nul && (
  echo KeepAwakewas killed
) || (
  echo KeepAwake is not running
)