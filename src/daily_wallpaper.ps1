param(
    [switch]$install = $false,
    [switch]$uninstall = $false
)

if (-not $(Get-ScheduledTask | Where-Object {$_.TaskName -like "Daily Wallpaper"})) {
    $install = $true
}

$exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
if ($install -or $uninstall) {
    if (!([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
        $mutex.Close()
        Start-Process -FilePath $exePath -Verb runas -ArgumentList $(If ($install) {"-install"} Else {"-uninstall"})
        exit
    }
    else {
        if ($install) {
            $Trigger = @(
                $(New-ScheduledTaskTrigger -AtLogOn),
                $(New-ScheduledTaskTrigger -AtStartup)
            )
            $Action = New-ScheduledTaskAction -Execute $exePath
            Register-ScheduledTask -TaskName "Daily Wallpaper" -Trigger $Trigger -Action $Action -Force
        }
        else {
            Unregister-ScheduledTask -TaskName "Daily Wallpaper" -Confirm:$false
            exit
        }
    }
}

$mutexName = "Global\DWMutex"
$createdNew = $false
$mutex = New-Object System.Threading.Mutex($false, $mutexName, [ref]$createdNew)
if (-not $createdNew) {
    exit
}

$client_id = $env.UNSPLASH_ID

[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')       | out-null
[System.Reflection.Assembly]::LoadWithPartialName('presentationframework')      | out-null
[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')          | out-null
[System.Reflection.Assembly]::LoadWithPartialName('WindowsFormsIntegration') | out-null
 
$icon = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\LiveCaptions.exe")

$Main_Tool_Icon = New-Object System.Windows.Forms.NotifyIcon
$Main_Tool_Icon.Text = "Daily Wallpaper"
$Main_Tool_Icon.Icon = $icon
$Main_Tool_Icon.Visible = $true

$global:query = "nature"
$global:overlay = $true
$global:image = $false
$global:setting = $false

$default_config = @{
    interval_text = "Randomize daily"
    interval_duration = 24 * 60 * 60 * 1000
    overlay = $true
    query = "nature"
}
$global:config = $default_config
$config_path = "$($env:USERPROFILE)\Pictures\Saved Wallpapers\dw_config.json"
$wallpaper_path = "$env:temp\\wallpaper.png"

$Menu_Refresh = New-Object System.Windows.Forms.MenuItem
$Menu_Refresh.Text = "Randomize"

$Menu_Save = New-Object System.Windows.Forms.MenuItem
$Menu_Save.Text = "Save wallpaper"

$Menu_interval = new-object System.Windows.Forms.MenuItem
$Menu_interval.Text = "Randomize daily"
$interval_never = new-object System.Windows.Forms.MenuItem -Property @{ Text = "Never" }
$interval_1day = new-object System.Windows.Forms.MenuItem -Property @{ Text = "Daily" }
$interval_1hr = new-object System.Windows.Forms.MenuItem -Property @{ Text = "Hourly" }
$interval_30min = new-object System.Windows.Forms.MenuItem -Property @{ Text = "Every 30 min" }
$interval_10min = new-object System.Windows.Forms.MenuItem -Property @{ Text = "Every 10 min" }
$Menu_interval.MenuItems.Add($interval_never)
$Menu_interval.MenuItems.Add($interval_1day)
$Menu_interval.MenuItems.Add($interval_1hr)
$Menu_interval.MenuItems.Add($interval_30min)
$Menu_interval.MenuItems.Add($interval_10min)

function set-interval {
    param (
        $text = "Randomize daily",
        $duration = 24 * 60 * 60 * 1000
    )
    $Menu_interval.Text = $text
    if ($text -eq "Never randomize") {
        $global:daily_timer.Stop()
    }
    else {
        $global:daily_timer.Start()
        $global:daily_timer.Interval = $duration
    }
    $global:config | Add-Member -MemberType NoteProperty -Name "interval_text" -Value $text -Force
    $global:config | Add-Member -MemberType NoteProperty -Name "interval_duration" -Value $duration -Force
    write-config
}

$interval_never.Add_Click({
    set-interval -text "Never randomize"
})
$interval_1day.Add_Click({
        set-interval -text "Randomize daily" -duration $(24 * 60 * 60 * 1000)
})
$interval_1hr.Add_Click({
        set-interval -text "Randomize hourly" -duration $(60 * 60 * 1000)
})
$interval_30min.Add_Click({
        set-interval -text "Randomize every 30 min" -duration $(30 * 60 * 1000)
})
$interval_10min.Add_Click({
        set-interval -text "Randomize every 10 min" -duration $(10 * 60 * 1000)
})

function write-config {
    if (-Not $(Test-Path "$($env:USERPROFILE)\\Pictures\\Saved Wallpapers" -PathType Container)) {
        New-Item -Path "$($env:USERPROFILE)\\Pictures" -Name "Saved Wallpapers" -ItemType "directory"
    }
    $global:config | ConvertTo-Json -depth 1 | Set-Content $config_path
}

$Menu_Overlay = New-Object System.Windows.Forms.MenuItem
$Menu_Overlay.Text = "Toggle overlay"

$Menu_Saved = New-Object System.Windows.Forms.MenuItem
$Menu_Saved.Text = "Select wallpaper"

$Menu_Saved_folder = New-Object System.Windows.Forms.MenuItem
$Menu_Saved_folder.Text = "Open saved folder"

$Menu_open = New-Object System.Windows.Forms.MenuItem
$Menu_open.Text = "Open in explorer"

$global:Menu_Query = New-Object System.Windows.Forms.MenuItem
$global:Menu_Query.Text = "Search for: $global:query"

$Menu_Exit = New-Object System.Windows.Forms.MenuItem
$Menu_Exit.Text = "Quit"
 
$Context_Menu = New-Object System.Windows.Forms.ContextMenu
$Main_Tool_Icon.ContextMenu = $Context_Menu
$Main_Tool_Icon.ContextMenu.MenuItems.AddRange($Menu_Refresh)
$Main_Tool_Icon.ContextMenu.MenuItems.AddRange($Menu_interval)
$Main_Tool_Icon.ContextMenu.MenuItems.AddRange($global:Menu_Query)
$Main_Tool_Icon.ContextMenu.MenuItems.AddRange($Menu_Overlay)
$Main_Tool_Icon.ContextMenu.MenuItems.Add("-")
$Main_Tool_Icon.ContextMenu.MenuItems.AddRange($Menu_Save)
$Main_Tool_Icon.ContextMenu.MenuItems.AddRange($Menu_Saved)
$Main_Tool_Icon.ContextMenu.MenuItems.AddRange($Menu_Saved_folder)
$Main_Tool_Icon.ContextMenu.MenuItems.Add("-")
$Main_Tool_Icon.ContextMenu.MenuItems.AddRange($Menu_Open)
$Main_Tool_Icon.ContextMenu.MenuItems.Add("-")
$Main_Tool_Icon.ContextMenu.MenuItems.AddRange($Menu_Exit)

$Main_Tool_Icon.Add_Click({
        If ($_.Button -eq [Windows.Forms.MouseButtons]::Left) {
            $Main_Tool_Icon.GetType().GetMethod("ShowContextMenu", [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic).Invoke($Main_Tool_Icon, $null)
        }
    })

$Menu_Refresh.add_Click({
        Set-Wallpaper($true)
    })

$Menu_Overlay.add_Click({
        $global:overlay = -Not $global:overlay
        $global:config | Add-Member -MemberType NoteProperty -Name "overlay" -Value $global:overlay -Force
        write-config
        Write-Overlay
    })

$Menu_Save.add_Click({
        Save-Wallpaper
    })

$Menu_Saved.add_Click({
        Select-Saved-Wallpaper
    })

$Menu_Saved_folder.add_Click({
        explorer /root,"$($env:USERPROFILE)\Pictures\Saved Wallpapers"
    })
 
$Menu_open.add_Click({
        explorer /select,$exePath
    })

function alert {
    param (
        $message = "Message",
        $title = "Daily Wallpaper"
    )
    [System.Windows.Forms.MessageBox]::Show($message, $title, 
        [System.Windows.Forms.MessageBoxButtons]::OK, 
        [System.Windows.Forms.MessageBoxIcon]::Information)
}

$global:Menu_Query.add_Click({
        Set-Wallpaper-Dialog
    })
 
$Menu_Exit.add_Click({
        $Main_Tool_Icon.Visible = $false
        $global:timer.Stop()
        $global:daily_timer.Stop()
        $global:overlay = $false
        Write-Overlay
        Stop-Process $pid
    })

function Set-Wallpaper {
    param(
        $manual = $false
    )
    $global:setting = $true
    if ($global:image) {
        $global:image.Dispose()
    }
    $response = Invoke-RestMethod -Uri "https://api.unsplash.com/photos/random?client_id=$client_id&query=$($global:query.Replace(' ','+'))&w=1920&h=1080&orientation=landscape"
    if ($manual) {
        $global:loading = $true
        Set-Loading-Wallpaper($response.color)
        $global:loading = $false
    }
    Invoke-WebRequest $response.urls.full -OutFile $wallpaper_path
    Preprocess-Image
    $global:setting = $false
    Write-Overlay
}

function Save-Wallpaper {
    if (-Not $(Test-Path "$($env:USERPROFILE)\\Pictures\\Saved Wallpapers" -PathType Container)) {
        New-Item -Path "$($env:USERPROFILE)\\Pictures" -Name "Saved Wallpapers" -ItemType "directory"
    }
    $id = 0
    $path = "$($env:USERPROFILE)\\Pictures\\Saved Wallpapers\\saved$id.png"
    while ($(Test-Path $path -PathType Leaf)) {
        $id++
        $path = "$($env:USERPROFILE)\\Pictures\\Saved Wallpapers\\saved$id.png"
    }
    Copy-Item -Path $wallpaper_path -Destination $path
    alert("Successfully saved wallpaper")
}

function Select-Saved-Wallpaper {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        InitialDirectory = "$($env:USERPROFILE)\Pictures\Saved Wallpapers"
        Filter           = "Images (*.png)|*.png"
    }
    $dialog.ShowDialog() | Out-Null
    if ($dialog.FileName) {
        if ($global:image) {
            $global:image.Dispose()
        }
        Copy-Item -Path $dialog.FileName -Destination $wallpaper_path
        Preprocess-Image
        Write-Overlay
    }
}

function Set-Wallpaper-Path {
    param (
        $imgPath = $wallpaper_path
    )
    $code = @'
    using System.Runtime.InteropServices;
    namespace Win32 { 
        public class Wallpaper {
            [DllImport("user32.dll", CharSet=CharSet.Auto)]
            static extern int SystemParametersInfo (int uAction, int uParam, string lpvParam, int fuWinIni);
            public static void SetWallpaper(string path) { 
                SystemParametersInfo(20, 0, path, 3);
            }
        }
    }
'@
    add-type $code
    [Win32.Wallpaper]::SetWallpaper($imgPath)
}

function Preprocess-Image {
    if ($global:image) {
        $global:image.Dispose()
    }
    $image = [System.Drawing.Image]::FromFile($wallpaper_path)
    $totalBrightness = 0
    $totalR = 0
    $totalG = 0
    $totalB = 0
    $sampleInterval = 10
    $totalSamples = 0
    for ($x = $image.Width/3; $x -lt $image.Width*2/3; $x += $sampleInterval) {
        for ($y = $image.Height / 5; $y -lt $image.Height / 3; $y += $sampleInterval) {
            $pixelColor = $image.GetPixel($x, $y)
            $brightness = ($pixelColor.R * 0.299) + ($pixelColor.G * 0.587) + ($pixelColor.B * 0.114)
            $totalBrightness += $brightness
            $totalR += $pixelColor.R
            $totalG += $pixelColor.G
            $totalB += $pixelColor.B
            $totalSamples++
        }
    }
    $image.Dispose()
    $averageBrightness = $totalBrightness / $totalSamples
    $averageR = [math]::Round($totalR / $totalSamples)
    $averageG = [math]::Round($totalG / $totalSamples)
    $averageB = [math]::Round($totalB / $totalSamples)
    $imageAccent = @($averageBrightness, $averageR, $averageG, $averageB)

    $offset = 120
    if ($averageBrightness -gt 128) {
        $r = [math]::Max(0, $imageAccent[1] - $offset)
        $g = [math]::Max(0, $imageAccent[2] - $offset)
        $b = [math]::Max(0, $imageAccent[3] - $offset)
    }
    else {
        $r = [math]::Min(255, $imageAccent[1] + $offset)
        $g = [math]::Min(255, $imageAccent[2] + $offset)
        $b = [math]::Min(255, $imageAccent[3] + $offset)
    }
    $color = [System.Drawing.Color]::FromArgb($r, $g, $b)
    $global:overlay_brush = new-object System.Drawing.SolidBrush($color)
}

function Write-Overlay {
    if ($global:overlay -and $(-not $global:loading) -and $global:overlay_brush -and $(-not $global:setting)) {
        $bitmap = New-Object System.Drawing.Bitmap(1920, 1080)
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAlias
        $image = [System.Drawing.Image]::FromFile($wallpaper_path)
        $graphics.DrawImage($image, 0, 0, 1920, 1080)

        $brush = $global:overlay_brush
        $primaryFontName = "SF Pro"
        $fallbackFontName = "Segoe UI"
        $primaryFontStyle = [System.Drawing.FontStyle]::Bold
        $fallbackFontStyle = [System.Drawing.FontStyle]::Regular
        $primaryFontExists = [System.Drawing.FontFamily]::Families.Name -contains $primaryFontName
        $fontName = if ($primaryFontExists) { $primaryFontName } else { $fallbackFontName }
        $fontStyle = if ($primaryFontExists) { $primaryFontStyle } else { $fallbackFontStyle }
        $font = new-object System.Drawing.Font($fontName, 96, $fontStyle)
        $subfont = new-object System.Drawing.Font($fontName, 18, $fontStyle)
        $stringFormat = new-object System.Drawing.StringFormat
        $stringFormat.Alignment = [System.Drawing.StringAlignment]::Center
        $stringFormat.LineAlignment = [System.Drawing.StringAlignment]::Center
        $rect = new-object System.Drawing.RectangleF(0, 0, 1920, 550)
        $time = Get-Date -Format "h:mm"
        $graphics.DrawString($time, $font, $brush, $rect, $stringFormat)
        $rect = new-object System.Drawing.RectangleF(0, 0, 1920, 700)
        $months = @("January", "Febuary", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
        $day = Get-Date -Format "dddd"
        $month = Get-Date -Format "MM"
        $date = Get-Date -UFormat "%e"
        $graphics.DrawString("$day, $($months[$month.ToString()]) $date", $subfont, $brush, $rect, $stringFormat)

        $bitmap.Save("$env:temp\\overlay_wallpaper.png", [System.Drawing.Imaging.ImageFormat]::Png)
        $graphics.Dispose()
        $bitmap.Dispose()
        $image.Dispose()
        Set-Wallpaper-Path("$env:temp\\overlay_wallpaper.png")
    }
    else {
        Set-Wallpaper-Path($wallpaper_path)
    }
}

function Set-Loading-Wallpaper {
    param(
        $color = "#ffffff"
    )
    $bitmap = New-Object System.Drawing.Bitmap 1920, 1080
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAlias
    $c = [System.Drawing.ColorTranslator]::FromHtml($color)
    $graphics.Clear($c)
    $brightness = ($c.R * 0.299) + ($c.G * 0.587) + ($c.B * 0.114)
    if ($brightness -gt 128) {
        $brush = new-object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(55, 55, 55))
    }
    else {
        $brush = new-object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(200, 200, 200))
    }
    $font = new-object System.Drawing.Font("SF Pro", 18, [System.Drawing.FontStyle]::Bold)
    $stringFormat = new-object System.Drawing.StringFormat
    $stringFormat.Alignment = [System.Drawing.StringAlignment]::Center
    $stringFormat.LineAlignment = [System.Drawing.StringAlignment]::Center
    $rect = new-object System.Drawing.RectangleF(0, 0, 1920, 1080)
    $graphics.DrawString("Fetching new wallpaper...", $font, $brush, $rect, $stringFormat)
    $graphics.Dispose()
    $tempPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "wallpaper.bmp")
    $bitmap.Save($tempPath)
    $bitmap.Dispose()
    Set-Wallpaper-Path($tempPath)
}

function Set-Wallpaper-Dialog {
    $global:dialog = New-Object System.Windows.Forms.Form
    $global:dialog.Visible = $true
    $global:dialog.width = 250
    $global:dialog.height = 135
    $global:dialog.FormBorderStyle = 3
    $global:dialog.StartPosition = 1
    $textLabel = New-Object System.Windows.Forms.Label
    $textLabel.Left = 10
    $textLabel.Top = 10
    $textLabel.Text = "Set search query"
    $global:textBox = New-Object System.Windows.Forms.TextBox
    $global:textBox.Left = 10
    $global:textBox.Top = 30
    $global:textBox.Width = 210
    $global:textBox.text = $global:query
    $confirmation = New-Object System.Windows.Forms.Button
    $confirmation.text = "Ok"
    $confirmation.left = 10
    $confirmation.width = 210
    $confirmation.top = 60
    $confirmation.DialogResult = 1
    $confirmation.add_Click({
            $global:query = $global:textBox.Text
            $global:Menu_Query.Text = "Search for: $global:query"
            $global:dialog.Close()
            $global:config | Add-Member -MemberType NoteProperty -Name "query" -Value $global:query -Force
            write-config
            Set-Wallpaper($true)
        })
    $global:dialog.Controls.Add($global:textBox)
    $global:dialog.Controls.Add($confirmation)
    $global:dialog.Controls.Add($textLabel)
    $global:dialog.AcceptButton = $confirmation
}

$global:timer = New-Object System.Windows.Forms.Timer
$global:timer.Interval = 1000
$global:prev_date = Get-Date -Format "dddd MM/dd/yyyy HH:mm K"
$global:timer.Add_Tick({
        $current_date = $(Get-Date -Format "dddd MM/dd/yyyy HH:mm K")
        if ($global:prev_date -ne $current_date) {
            Write-Overlay
        }
        $global:prev_date = Get-Date -Format "dddd MM/dd/yyyy HH:mm K"
    })
$global:timer.Start()

$global:daily_timer = new-object System.Windows.Forms.Timer
$global:daily_timer.Interval = 24 * 60 * 60 * 1000
$global:daily_timer.Add_Tick({
    Set-Wallpaper($false)
})
$global:daily_timer.Start()

if ($(Test-Path $config_path -PathType Leaf)) {
    $global:config = Get-Content $config_path | ConvertFrom-Json
    if ($global:config -ne $null) {
        if ($($global:config.PSobject.Properties.name -match "interval_text") -and $($global:config.PSobject.Properties.name -match "interval_duration")) {
            set-interval -text $global:config.interval_text -duration $global:config.interval_duration
        }
        else {
            $global:config | Add-Member -MemberType NoteProperty -Name "interval_text" -Value $default_config.interval_text -Force
            $global:config | Add-Member -MemberType NoteProperty -Name "interval_duration" -Value $default_config.interval_duration -Force
            write-config
        }
        if ($global:config.PSobject.Properties.name -match "overlay") {
            $global:overlay = $global:config.overlay
            Write-Overlay
        }
        else {
            $global:config | Add-Member -MemberType NoteProperty -Name "overlay" -Value $default_config.overlay -Force
            write-config
        }
        if ($global:config.PSobject.Properties.name -match "query") {
            $global:query = $global:config.query
        }
        else {
            $global:config | Add-Member -MemberType NoteProperty -Name "query" -Value $default_config.query -Force
            write-config
        }
        $global:Menu_Query.Text = "Search for: $global:query"
    }
    else {
        $global:config = $default_config
        write-config
    }

    if (-not $(Test-Path $wallpaper_path -PathType Leaf)) {
        Set-Wallpaper($false)
        Preprocess-Image
        Write-Overlay
    }
    else {
        Preprocess-Image
        Write-Overlay
    }
}
else {
    write-config
    Set-Wallpaper($false)
    Preprocess-Image
    Write-Overlay
}

$windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
$asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru
$null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)

[System.GC]::Collect()
 
$appContext = New-Object System.Windows.Forms.ApplicationContext
[void][System.Windows.Forms.Application]::Run($appContext)