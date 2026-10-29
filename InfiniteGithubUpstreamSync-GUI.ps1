#Requires -Version 5.1
<#
  InfiniteGithubUpstreamSync-GUI.ps1 —— 图形界面版（独立文件）

  与主程序 InfiniteGithubUpstreamSync.ps1 放在同一目录使用。
  功能与命令行版一致：仓库/间隔/排除规则设置、立即同步、托盘与自启控制、实时日志。

  推荐启动方式（无 PowerShell 黑窗，主程序菜单和托盘就是这样调用的）：
    powershell -STA -WindowStyle Hidden -File InfiniteGithubUpstreamSync-GUI.ps1
  直接运行也可以：脚本检测到非 STA 会自动以隐藏窗口方式重启自己。
#>
param()

# 非 STA（双击/直接运行）时，以隐藏控制台 + STA 方式重启自己 —— 不显示 PowerShell 黑窗
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Start-Process -FilePath 'powershell.exe' -WindowStyle Hidden -ArgumentList (
        '-STA -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}"' -f $PSCommandPath)
    return
}

$Root       = $PSScriptRoot
$ConfigFile = Join-Path $Root 'sync-upstream.config.json'
$LogFile    = Join-Path $Root 'sync-upstream.log'
$IconPath   = Join-Path $Root 'InfiniteGithubUpstreamSync.ico'
$VbsPath    = Join-Path $Root 'InfiniteGithubUpstreamSync-hidden.vbs'
$MainScript = Join-Path $Root 'InfiniteGithubUpstreamSync.ps1'
$RunKey     = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$RunValue   = 'InfiniteGithubUpstreamSync'

# ============================================================================
# 与主程序同款的辅助函数（独立封装，不依赖主脚本）
# ============================================================================
function Test-RepoUrl([string]$Url) {
    if ([string]::IsNullOrWhiteSpace($Url)) { return $false }
    $u = $Url.Trim().TrimEnd('/')
    return ($u -match '^https://github\.com/[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+(\.git)?$')
}

function Read-Config {
    if (-not (Test-Path -LiteralPath $ConfigFile)) { return $null }
    try { return Get-Content -LiteralPath $ConfigFile -Raw -Encoding UTF8 | ConvertFrom-Json } catch { return $null }
}

function Get-IntervalMinutes {
    $c = Read-Config
    if ($c -and $c.IntervalMinutes -and [int]$c.IntervalMinutes -gt 0) { return [int]$c.IntervalMinutes }
    if ($c -and $c.IntervalHours) { return [int]$c.IntervalHours * 60 }
    return 720
}

function Get-Excludes {
    $c = Read-Config
    if ($c -and $c.Excludes) { return @($c.Excludes | ForEach-Object { [string]$_ } | Where-Object { $_.Trim() -ne '' }) }
    return @()
}

function Save-GuiConfig([string]$Repo, [int]$Minutes, [string[]]$Excludes) {
    @{ RepoUrl = $Repo; IntervalMinutes = $Minutes; Excludes = @($Excludes) } |
        ConvertTo-Json | Set-Content -LiteralPath $ConfigFile -Encoding UTF8
}

# ---- 检测更新（与主脚本保持一致；版本号与更新源改这里） -----------------------
$ToolVersion = '1.0.0'

# 本工具的更新源（内置固定，不开放给用户修改；与主脚本的 $ToolRepo 保持一致）
$ToolRepo = 'https://github.com/ted-li-git/Infinite_gh_Upstream_Sync'

function ConvertTo-VersionNum([string]$v) {
    if ([string]::IsNullOrWhiteSpace($v)) { return $null }
    $m = [regex]::Match(($v.Trim() -replace '^[vV]', ''), '\d+(\.\d+){0,3}')
    if (-not $m.Success) { return $null }
    try { return [version]$m.Value } catch { return $null }
}

# 缓存目录按仓库地址区分（与主脚本同一算法，共用同一份缓存）
function Get-CacheDir([string]$RepoUrl) {
    $h = [System.BitConverter]::ToString(
        [System.Security.Cryptography.SHA1]::Create().ComputeHash(
            [System.Text.Encoding]::UTF8.GetBytes($RepoUrl))).Replace('-', '').Substring(0, 8)
    return Join-Path $env:TEMP ('HungerCat-upstream-cache-' + $h)
}

function Test-UpstreamUpdate([string]$RepoUrl) {
    $r = @{ Error = ''; Files = @(); Branch = '' }
    $cache = Get-CacheDir $RepoUrl
    if (-not (Test-Path -LiteralPath (Join-Path $cache '.git'))) {
        $r.Error = '本地还没有同步缓存，请先同步一次再检测'
        return $r
    }
    git -C $cache fetch origin --quiet 2>$null
    if ($LASTEXITCODE -ne 0) { $r.Error = '无法连接 GitHub（fetch 失败，请检查网络）'; return $r }
    $r.Branch = (& git -C $cache rev-parse --abbrev-ref HEAD).Trim()
    if ($r.Branch -eq 'HEAD' -or $r.Branch -eq '') { $r.Branch = 'main' }
    $local  = (& git -C $cache rev-parse HEAD).Trim()
    $remote = (& git -C $cache rev-parse "origin/$($r.Branch)" 2>$null).Trim()
    if (-not $remote) { $r.Error = "找不到远程分支 origin/$($r.Branch)"; return $r }
    if ($local -eq $remote) { return $r }
    $r.Files = @(& git -C $cache diff --name-only $local $remote 2>$null | Where-Object { $_ -ne '' })
    return $r
}

function Test-ToolUpdate {
    $r = @{ Error = ''; Latest = ''; Url = ''; HasNew = $false }
    if (-not (Test-RepoUrl $ToolRepo)) { $r.Error = "内置更新源格式不正确：$ToolRepo"; return $r }
    $slug = $ToolRepo.Trim().TrimEnd('/') -replace '^https://github\.com/', '' -replace '\.git$', ''
    $headers = @{ 'User-Agent' = 'InfiniteGithubUpstreamSync'; 'Accept' = 'application/vnd.github+json' }
    try {
        $rel = Invoke-RestMethod -Uri "https://api.github.com/repos/$slug/releases/latest" -Headers $headers -TimeoutSec 20
        $r.Latest = [string]$rel.tag_name
        $r.Url = [string]$rel.html_url
    } catch {
        try {
            $tags = Invoke-RestMethod -Uri "https://api.github.com/repos/$slug/tags" -Headers $headers -TimeoutSec 20
            if ($tags.Count -gt 0) {
                $r.Latest = [string]$tags[0].name
                $r.Url = "https://github.com/$slug/releases/tag/$($tags[0].name)"
            }
        } catch {
            $code = $_.Exception.Response.StatusCode.value__
            if ($code -eq 404) { $r.Error = '内置更新源尚不可用（仓库或 tag 还没有发布）' }
            else { $r.Error = '无法访问 GitHub API：' + $_.Exception.Message }
            return $r
        }
    }
    if ($r.Latest -eq '') { $r.Error = '该仓库还没有 release 或 tag，无法比较版本'; return $r }
    $cur = ConvertTo-VersionNum $ToolVersion
    $new = ConvertTo-VersionNum $r.Latest
    if ($cur -and $new) { $r.HasNew = ($new -gt $cur) } else { $r.HasNew = ($r.Latest.TrimStart('v', 'V') -ne $ToolVersion) }
    return $r
}

function Get-TrayProcess {
    @(Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
        Where-Object {
            $_.ProcessId -ne $PID -and $_.CommandLine -and
            $_.CommandLine -like '*InfiniteGithubUpstreamSync.ps1*' -and
            $_.CommandLine -notlike '*-Once*' -and $_.CommandLine -notlike '*-Menu*' -and $_.CommandLine -notlike '*-GUI*'
        })
}

function Test-Autostart { [bool](Get-ItemProperty -Path $RunKey -Name $RunValue -ErrorAction SilentlyContinue) }

# ---- 本工具自我更新（GUI 版，与主脚本 Invoke-ToolUpdate 同一套流程） --------
$ToolFiles = @(
    'InfiniteGithubUpstreamSync.ps1'
    'InfiniteGithubUpstreamSync-GUI.ps1'
    'InfiniteGithubUpstreamSync.ico'
    '双击启动同步.bat'
)

function Restart-TrayIfRunning {
    $procs = Get-TrayProcess
    if (-not $procs -or $procs.Count -eq 0) { return $false }
    $procs | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    Start-Sleep -Milliseconds 900
    if (Test-Path -LiteralPath $VbsPath) {
        Start-Process -FilePath 'wscript.exe' -ArgumentList ('//B //Nologo "{0}"' -f $VbsPath)
    }
    return $true
}

function Invoke-ToolUpdateGui {
    $t = Test-ToolUpdate
    if ($t.Error) { return @{ Ok = $false; Msg = '无法检查更新：' + $t.Error } }
    if (-not $t.HasNew) { return @{ Ok = $false; Msg = '已是最新版本（' + $t.Latest + '），无需更新' } }

    $slug = $ToolRepo.Trim().TrimEnd('/') -replace '^https://github\.com/', '' -replace '\.git$', ''
    $tag = $t.Latest
    $ts = Get-Date -Format 'yyyyMMdd-HHmmss'
    $tmp = Join-Path $env:TEMP ('igus-update-' + $ts)
    New-Item -ItemType Directory -Path $tmp -Force | Out-Null

    $got = @(); $fail = @()
    foreach ($f in $ToolFiles) {
        $dst = Join-Path $tmp $f
        $uri = 'https://raw.githubusercontent.com/' + $slug + '/' + $tag + '/' + [uri]::EscapeDataString($f)
        try {
            Invoke-WebRequest -Uri $uri -OutFile $dst -UseBasicParsing -TimeoutSec 40
            if ((Get-Item -LiteralPath $dst).Length -lt 100) { $fail += $f; continue }
            if ($f -like '*.ps1') {
                $errs = $null
                $null = [System.Management.Automation.Language.Parser]::ParseFile($dst, [ref]$null, [ref]$errs)
                if ($errs -and $errs.Count -gt 0) { $fail += ($f + '(语法检查未通过)'); continue }
            }
            $got += $f
        } catch { $fail += $f }
    }
    if ($got -notcontains 'InfiniteGithubUpstreamSync.ps1') {
        Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
        return @{ Ok = $false; Msg = '未能下载到主脚本，已放弃更新。失败项：' + ($fail -join '、') }
    }

    $backup = Join-Path $Root ('.sync-script-backup\pre-selfupdate-' + $ts)
    New-Item -ItemType Directory -Path $backup -Force | Out-Null
    foreach ($f in $got) {
        $cur = Join-Path $Root $f
        if (Test-Path -LiteralPath $cur) { Copy-Item -LiteralPath $cur -Destination (Join-Path $backup $f) -Force }
    }

    $written = @()
    try {
        foreach ($f in $got) {
            Copy-Item -LiteralPath (Join-Path $tmp $f) -Destination (Join-Path $Root $f) -Force
            $written += $f
        }
    } catch {
        foreach ($f in $written) {
            $bk = Join-Path $backup $f
            if (Test-Path -LiteralPath $bk) { Copy-Item -LiteralPath $bk -Destination (Join-Path $Root $f) -Force }
        }
        Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
        return @{ Ok = $false; Msg = '写入失败，已回滚到更新前：' + $_.Exception.Message }
    }
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue

    $restarted = Restart-TrayIfRunning
    $msg = '已更新到 ' + $tag + '，替换 ' + $got.Count + ' 个文件：' + ($got -join '、')
    if ($fail.Count -gt 0) { $msg += '；未获取：' + ($fail -join '、') }
    if ($restarted) { $msg += '；托盘已重启' }
    $msg += '；原文件备份在 ' + $backup
    return @{ Ok = $true; Msg = $msg }
}

function Install-Autostart {
    if (-not (Test-Path -LiteralPath $ConfigFile)) { throw '请先保存配置（仓库地址 + 同步间隔）' }
    if (-not (Test-Path -LiteralPath $MainScript)) { throw '未找到主程序 InfiniteGithubUpstreamSync.ps1，请确认两个文件在同一目录' }
    $vbs = @"
' 由 InfiniteGithubUpstreamSync-GUI.ps1 自动生成，请勿手动修改
Set sh = CreateObject("WScript.Shell")
sh.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ""$MainScript""", 0, False
"@
    [System.IO.File]::WriteAllText($VbsPath, $vbs, [System.Text.Encoding]::Default)
    if (-not (Test-Path -LiteralPath $RunKey)) { New-Item -Path $RunKey -Force | Out-Null }
    Set-ItemProperty -Path $RunKey -Name $RunValue -Value ('wscript.exe //B //Nologo "{0}"' -f $VbsPath)
    if (-not (Get-TrayProcess)) {
        Start-Process -FilePath 'wscript.exe' -ArgumentList ('//B //Nologo "{0}"' -f $VbsPath)
    }
}

function Remove-Autostart {
    Get-TrayProcess | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    Remove-ItemProperty -Path $RunKey -Name $RunValue -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $VbsPath) { Remove-Item -LiteralPath $VbsPath -Force }
}

function Write-GuiError([string]$msg) {
    try {
        Add-Content -LiteralPath (Join-Path $Root 'InfiniteGithubUpstreamSync-error.log') `
            -Value ('[{0}] GUI: {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $msg) -Encoding UTF8
    } catch {}
}

function Show-GuiMessage([string]$text, [string]$icon = 'Warning') {
    [void][System.Windows.Forms.MessageBox]::Show($text, 'InfiniteGithubUpstreamSync', 'OK', $icon)
}

# ============================================================================
# 界面
# ============================================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

try {
    $cfg = Read-Config
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'InfiniteGithubUpstreamSync'
    $form.ClientSize = New-Object System.Drawing.Size(664, 700)
    $form.StartPosition = 'CenterScreen'
    $form.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
    try { $form.Icon = New-Object System.Drawing.Icon($IconPath) } catch {}

    $mkLabel = { param($text, $x, $y, $w)
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $text; $l.Location = New-Object System.Drawing.Point($x, $y)
        $l.Size = New-Object System.Drawing.Size($w, 20); $l.AutoSize = $false
        $form.Controls.Add($l); return $l }
    $mkButton = { param($text, $x, $y, $w, $h)
        $b = New-Object System.Windows.Forms.Button
        $b.Text = $text; $b.Location = New-Object System.Drawing.Point($x, $y)
        $b.Size = New-Object System.Drawing.Size($w, $h)
        $form.Controls.Add($b); return $b }

    # ---- 设置区 ----
    & $mkLabel '仓库地址：' 12 15 80 | Out-Null
    $txtRepo = New-Object System.Windows.Forms.TextBox
    $txtRepo.Location = New-Object System.Drawing.Point(95, 12)
    $txtRepo.Size = New-Object System.Drawing.Size(555, 25)
    $txtRepo.Text = $(if ($cfg -and $cfg.RepoUrl) { $cfg.RepoUrl } else { '' })
    $form.Controls.Add($txtRepo)

    & $mkLabel '同步间隔（分钟）：' 12 50 130 | Out-Null
    $numInt = New-Object System.Windows.Forms.NumericUpDown
    $numInt.Location = New-Object System.Drawing.Point(145, 47)
    $numInt.Size = New-Object System.Drawing.Size(90, 25)
    $numInt.Minimum = 1; $numInt.Maximum = 100000
    $numInt.Value = [Math]::Max(1, (Get-IntervalMinutes))
    $form.Controls.Add($numInt)

    & $mkLabel '排除规则（点“浏览文件…”从资源管理器选择；也可直接输入通配规则，如 *.log、node_modules）：' 12 85 640 | Out-Null
    $lstEx = New-Object System.Windows.Forms.ListBox
    $lstEx.Location = New-Object System.Drawing.Point(12, 108)
    $lstEx.Size = New-Object System.Drawing.Size(530, 115)
    $lstEx.SelectionMode = 'MultiExtended'
    Get-Excludes | ForEach-Object { [void]$lstEx.Items.Add($_) }
    $form.Controls.Add($lstEx)

    $txtPat = New-Object System.Windows.Forms.TextBox
    $txtPat.Location = New-Object System.Drawing.Point(12, 230)
    $txtPat.Size = New-Object System.Drawing.Size(268, 25)
    $form.Controls.Add($txtPat)
    $btnAddEx = & $mkButton '添加输入' 288 228 88 27
    $btnPick  = & $mkButton '浏览文件…' 384 228 118 27
    $btnDelEx = & $mkButton '删除所选' 510 228 140 27

    # 把绝对路径转成相对同步目录的路径；不在同步目录内则返回 $null
    $toRel = {
        param($full)
        $rootFull = (Resolve-Path -LiteralPath $Root).Path.TrimEnd('\')
        if ($full.StartsWith($rootFull, [StringComparison]::OrdinalIgnoreCase)) {
            return $full.Substring($rootFull.Length).TrimStart('\')
        }
        return $null
    }

    # 浏览文件：打开资源管理器选择（可多选），自动转成相对路径加入排除列表
    $btnPick.add_Click({
            $dlg = New-Object System.Windows.Forms.OpenFileDialog
            $dlg.Title = '选择不参与同步的文件（可多选）'
            $dlg.InitialDirectory = $Root
            $dlg.Multiselect = $true
            $dlg.Filter = '所有文件 (*.*)|*.*'
            if ($dlg.ShowDialog() -ne 'OK') { return }
            $bad = 0; $added = 0
            foreach ($f in $dlg.FileNames) {
                $rel = & $toRel $f
                if (-not $rel) { $bad++; continue }
                if (-not $lstEx.Items.Contains($rel)) { [void]$lstEx.Items.Add($rel); $added++ }
            }
            if ($bad -gt 0) {
                Show-GuiMessage ("有 $bad 个文件不在同步目录 $Root 内，已忽略。`n排除规则只对同步目录内的文件生效。")
            }
            if ($added -gt 0) { $lblCfgMsg.ForeColor = [System.Drawing.Color]::DarkOrange; $lblCfgMsg.Text = "已添加 $added 条，记得点『保存配置』" }
        })

    $btnAddEx.add_Click({
            $p = $txtPat.Text.Trim()
            if ($p -ne '' -and -not $lstEx.Items.Contains($p)) { [void]$lstEx.Items.Add($p); $txtPat.Text = '' }
        })
    $btnDelEx.add_Click({
            $sel = @($lstEx.SelectedItems)
            foreach ($s in $sel) { $lstEx.Items.Remove($s) }
        })

    $btnSave = & $mkButton '保存配置' 12 268 150 32
    $lblCfgMsg = & $mkLabel '' 170 274 480

    # ---- 更新（更新源内置，不提供修改入口） ----
    $btnCheck = & $mkButton '检测更新' 12 303 110 28
    $btnUpdate = & $mkButton '更新到最新版' 130 303 120 28
    & $mkLabel ("当前版本 $ToolVersion") 262 308 390 | Out-Null

    # ---- 控制区 ----
    $btnSync  = & $mkButton '立即同步一次' 12 342 130 34
    $btnStart = & $mkButton '启动托盘 + 开机自启' 150 342 165 34
    $btnStop  = & $mkButton '停止托盘 + 取消自启' 323 342 170 34
    $btnExit  = & $mkButton '退出托盘' 501 342 100 34

    $chkAuto = New-Object System.Windows.Forms.CheckBox
    $chkAuto.Text = '开机自启'
    $chkAuto.Location = New-Object System.Drawing.Point(12, 386)
    $chkAuto.Size = New-Object System.Drawing.Size(90, 24)
    $chkAuto.Checked = (Test-Autostart)
    $form.Controls.Add($chkAuto)
    $btnConsole = & $mkButton '打开控制台' 110 382 110 28
    $btnConsole.add_Click({
            $bat = Join-Path $Root '双击启动同步.bat'
            if (Test-Path -LiteralPath $bat) { Start-Process -FilePath 'cmd.exe' -ArgumentList ('/c ""{0}""' -f $bat) }
        })
    $lblStatus = & $mkLabel '' 230 386 420

    # ---- 日志区（自动刷新） ----
    & $mkLabel '同步日志（自动刷新）：' 12 422 300 | Out-Null
    $txtLog = New-Object System.Windows.Forms.TextBox
    $txtLog.Location = New-Object System.Drawing.Point(12, 445)
    $txtLog.Size = New-Object System.Drawing.Size(640, 220)
    $txtLog.Multiline = $true
    $txtLog.ReadOnly = $true
    $txtLog.ScrollBars = 'Vertical'
    $txtLog.Font = New-Object System.Drawing.Font('Consolas', 8.5)
    $txtLog.BackColor = [System.Drawing.Color]::White
    $form.Controls.Add($txtLog)

    $script:guiSyncProc = $null
    $script:guiLastLogWrite = [DateTime]::MinValue

    $refreshLog = {
        if (Test-Path -LiteralPath $LogFile) {
            $lw = (Get-Item -LiteralPath $LogFile).LastWriteTime
            if ($lw -ne $script:guiLastLogWrite) {
                $script:guiLastLogWrite = $lw
                $txtLog.Lines = @(Get-Content -LiteralPath $LogFile -Tail 200 -Encoding UTF8)
                $txtLog.SelectionStart = $txtLog.TextLength
                $txtLog.ScrollToCaret()
            }
        }
    }
    $refreshStatus = {
        $tray = Get-TrayProcess
        $lblStatus.Text = '托盘：' + $(if ($tray) { "运行中（PID $($tray[0].ProcessId)）" } else { '未运行' }) +
            '    自启：' + $(if (Test-Autostart) { '已设置' } else { '未设置' })
        if ($chkAuto.Checked -ne (Test-Autostart)) { $chkAuto.Checked = (Test-Autostart) }
    }

    $btnSave.add_Click({
            if (-not (Test-RepoUrl $txtRepo.Text)) {
                Show-GuiMessage '仓库地址格式不正确，应形如 https://github.com/用户名/仓库名'
                return
            }
            Save-GuiConfig $txtRepo.Text.Trim() ([int]$numInt.Value) ([string[]]@($lstEx.Items | ForEach-Object { [string]$_ }))
            $lblCfgMsg.ForeColor = [System.Drawing.Color]::Green
            $lblCfgMsg.Text = '已保存 ' + (Get-Date -Format 'HH:mm:ss') + '（托盘的定时同步会在重启托盘后使用新间隔）'
        })

    # 检测更新：上游是否有新提交 + 本工具是否有新版本
    $btnCheck.add_Click({
            $btnCheck.Enabled = $false
            $oldText = $btnCheck.Text
            $btnCheck.Text = '检测中…'
            try {
                & $refreshStatus
                $dup = Test-UpstreamUpdate $txtRepo.Text.Trim()
                if ($dup.Error) { $uText = $dup.Error }
                elseif ($dup.Files.Count -eq 0) { $uText = '已是最新，没有新提交' }
                else {
                    $uText = "有 $($dup.Files.Count) 个文件将变化（分支 $($dup.Branch)）"
                    $dup.Files | Select-Object -First 10 | ForEach-Object { $uText += "`n    $_" }
                    if ($dup.Files.Count -gt 10) { $uText += "`n    ……其余 $($dup.Files.Count - 10) 个" }
                }

                $dt = Test-ToolUpdate
                if ($dt.Error) { $tText = $dt.Error }
                elseif ($dt.HasNew) { $tText = "发现新版本 $($dt.Latest)（当前 $ToolVersion）`n$($dt.Url)" }
                else { $tText = "已是最新（$($dt.Latest)）" }

                $msg = "本工具版本：$ToolVersion`n`n【上游仓库】`n$uText`n`n【工具更新】`n$tText"
                if ($dt.HasNew -and $dt.Url) {
                    $r = [System.Windows.Forms.MessageBox]::Show($msg + "`n`n现在下载并更新到最新版？（选『否』则打开下载页）", 'InfiniteGithubUpstreamSync',
                        [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Information)
                    if ($r -eq 'Yes') {
                        $ur = Invoke-ToolUpdateGui
                        [void][System.Windows.Forms.MessageBox]::Show($ur.Msg, 'InfiniteGithubUpstreamSync', 'OK',
                            $(if ($ur.Ok) { 'Information' } else { 'Warning' }))
                        if ($ur.Ok) { & $refreshStatus }
                    } else { Start-Process $dt.Url }
                } else {
                    [void][System.Windows.Forms.MessageBox]::Show($msg, 'InfiniteGithubUpstreamSync', 'OK', 'Information')
                }
            } catch {
                Write-GuiError ('检测更新失败: ' + $_.Exception.Message)
                Show-GuiMessage ('检测失败：' + $_.Exception.Message) 'Error'
            } finally {
                $btnCheck.Enabled = $true
                $btnCheck.Text = $oldText
            }
        })

    # 更新到最新版：检查 → 下载 → 校验 → 备份 → 替换 → 重启托盘
    $btnUpdate.add_Click({
            $btnUpdate.Enabled = $false
            $oldText = $btnUpdate.Text
            $btnUpdate.Text = '更新中…'
            try {
                $ur = Invoke-ToolUpdateGui
                [void][System.Windows.Forms.MessageBox]::Show($ur.Msg, 'InfiniteGithubUpstreamSync', 'OK',
                    $(if ($ur.Ok) { 'Information' } else { 'Warning' }))
                if ($ur.Ok) { & $refreshStatus }
            } catch {
                Write-GuiError ('自我更新失败: ' + $_.Exception.Message)
                Show-GuiMessage ('更新失败：' + $_.Exception.Message) 'Error'
            } finally {
                $btnUpdate.Enabled = $true
                $btnUpdate.Text = $oldText
            }
        })

    $btnSync.add_Click({
            if ($script:guiSyncProc -and -not $script:guiSyncProc.HasExited) { return }
            $btnSync.Enabled = $false
            $btnSync.Text = '同步中…'
            $arg = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}" -Once' -f $MainScript
            try { $script:guiSyncProc = Start-Process -FilePath 'powershell.exe' -ArgumentList $arg -WindowStyle Hidden -PassThru }
            catch { $btnSync.Enabled = $true; $btnSync.Text = '立即同步一次'; Show-GuiMessage $_.Exception.Message }
        })
    $btnStart.add_Click({
            try {
                if (-not (Test-Path -LiteralPath $ConfigFile)) {
                    if (-not (Test-RepoUrl $txtRepo.Text)) { Show-GuiMessage '请先填写正确的仓库地址并保存配置'; return }
                    Save-GuiConfig $txtRepo.Text.Trim() ([int]$numInt.Value) ([string[]]@($lstEx.Items | ForEach-Object { [string]$_ }))
                }
                Install-Autostart
            } catch { Show-GuiMessage $_.Exception.Message }
            & $refreshStatus
        })
    $btnStop.add_Click({ Remove-Autostart; & $refreshStatus })
    $btnExit.add_Click({
            Get-TrayProcess | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
            & $refreshStatus
        })
    $chkAuto.add_Click({
            try {
                if ($chkAuto.Checked) { Install-Autostart }
                else {
                    Remove-ItemProperty -Path $RunKey -Name $RunValue -ErrorAction SilentlyContinue
                    if (Test-Path -LiteralPath $VbsPath) { Remove-Item -LiteralPath $VbsPath -Force }
                }
            } catch { Show-GuiMessage $_.Exception.Message }
            & $refreshStatus
        })

    $guiTimer = New-Object System.Windows.Forms.Timer
    $guiTimer.Interval = 1500
    $guiTimer.add_Tick({
            & $refreshLog
            & $refreshStatus
            if ($script:guiSyncProc -and $script:guiSyncProc.HasExited) {
                $script:guiSyncProc = $null
                $btnSync.Enabled = $true
                $btnSync.Text = '立即同步一次'
            }
        })
    $guiTimer.Enabled = $true
    $form.add_FormClosing({ $guiTimer.Enabled = $false })

    & $refreshLog
    & $refreshStatus
    [void]$form.ShowDialog()
} catch {
    Write-GuiError ($_.Exception.Message + "`n" + $_.ScriptStackTrace)
    try { Show-GuiMessage "GUI 出错：$($_.Exception.Message)" 'Error' } catch {}
}
