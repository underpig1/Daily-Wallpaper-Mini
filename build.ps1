install-module ps2exe
$env = Get-Content "./secrets.env" | ConvertFrom-StringData
$script = Get-Content "./src/daily_wallpaper.ps1"
foreach ($key in $env.Keys) {
    $script = $script.Replace("`$env.$key", $env.$key)
}
$script | out-file -FilePath "./dist/daily_wallpaper.ps1"
ps2exe "./dist/daily_wallpaper.ps1" "./dist/daily_wallpaper.exe" -noConsole -noOutput -noError -title "Daily Wallpaper"