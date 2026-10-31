#Requires -Version 5.1
<#
  InfiniteGithubUpstreamSync.ps1 —— 单文件版「GitHub 上游自动同步」

  一个文件包含全部功能，不再需要其他 ps1：
    · 同步内核：同名文件以远程为准覆盖，远程新增的复制过来，本地独有文件保留
    · 控制台菜单（双击 .bat 进来）
    · 托盘常驻：定时同步、立即同步、查看实时日志、回到控制台、退出
    · 开机自启（注册表 Run + 隐藏启动器 VBS，自动生成）

  参数：
    -Menu              打开控制台菜单（默认行为，供 .bat 调用）
    -Once              只同步一次就退出（托盘后台就是用这个）
    -DryRun            试运行：只列出将要做出的改动，不实际写入
    -Install           设置开机自启并启动托盘
    -Remove            停止托盘并取消开机自启
    -Exit              只退出托盘（保留开机自启）
    -Status            查看运行状态 / 自启状态
    -CheckUpdate       只检测更新：上游是否有新提交 + 本工具是否有新版本
    -Update            检查并更新本工具到最新版（下载 → 校验 → 备份 → 替换 → 重启托盘）
    -RepoUrl / -IntervalMinutes   临时覆盖配置（不写入配置文件）

  图形界面在独立文件 InfiniteGithubUpstreamSync-GUI.ps1（与本脚本同目录）；
  菜单选 8 或托盘右键即可打开，无 PowerShell 黑窗。

  排除规则：配置文件 sync-upstream.config.json 里的 Excludes 数组。
    命中规则的上游文件不会被同步（不覆盖、不新增）。支持 * 通配符；
    写目录名（如 node_modules）则该目录整体跳过；写文件名（如 *.log）按名称匹配。
#>
param(
    [string]$RepoUrl = '',
    [int]$IntervalMinutes = 0,
    [switch]$Once,
    [switch]$DryRun,
    [switch]$Menu,
    [switch]$Install,
    [switch]$Remove,
    [switch]$Status,
    [switch]$Exit,
    [switch]$CheckUpdate,
    [switch]$Update
)

$Root       = $PSScriptRoot
$ConfigFile = Join-Path $Root 'sync-upstream.config.json'
$LogFile    = Join-Path $Root 'sync-upstream.log'
$BackupRoot = Join-Path $Root '.sync-backup'
$VbsPath    = Join-Path $Root 'InfiniteGithubUpstreamSync-hidden.vbs'
$IconPath   = Join-Path $Root 'InfiniteGithubUpstreamSync.ico'
$RunKey     = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$RunValue   = 'InfiniteGithubUpstreamSync'
$MutexName  = 'Local\InfiniteGithubUpstreamSync'
$DefaultRepo = 'https://github.com/JOHNNY-beep456/HungerCat_MineCraft_Launcher.git'

# 本工具自身的版本号：发布新版时改这里（GUI/托盘「检测更新」会与 GitHub 上的 tag 对比）
$ToolVersion = '1.0.2'

# 本工具的更新源（内置固定，不开放给用户修改；更换发布仓库只需改这一行）
$ToolRepo = 'https://github.com/ted-li-git/Infinite_gh_Upstream_Sync'

# 自动更新时会替换的本工具文件（仓库里缺少某个文件就跳过该文件）
$ToolFiles = @(
    'InfiniteGithubUpstreamSync.ps1'
    'InfiniteGithubUpstreamSync-GUI.ps1'
    'InfiniteGithubUpstreamSync.ico'
    '双击启动同步.bat'
)

# ============================================================================
# 配置
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
    if ($IntervalMinutes -gt 0) { return $IntervalMinutes }
    $c = Read-Config
    if ($c -and $c.IntervalMinutes -and [int]$c.IntervalMinutes -gt 0) { return [int]$c.IntervalMinutes }
    if ($c -and $c.IntervalHours) { return [int]$c.IntervalHours * 60 }   # 兼容旧配置
    return 720
}

# ---- 排除规则（Excludes） -------------------------------------------------
function Get-Excludes {
    $c = Read-Config
    if ($c -and $c.Excludes) { return @($c.Excludes | ForEach-Object { [string]$_ } | Where-Object { $_.Trim() -ne '' }) }
    return @()
}

# 统一写配置：保留已有 Excludes，除非显式传入
function Save-Config([string[]]$Excludes) {
    if ($null -eq $Excludes) { $Excludes = Get-Excludes }
    @{ RepoUrl = $script:RepoUrl; IntervalMinutes = $script:IntervalMinutes; Excludes = @($Excludes) } |
        ConvertTo-Json | Set-Content -LiteralPath $ConfigFile -Encoding UTF8
}

# 判断相对路径是否命中排除规则：
#   整条路径通配（build\*）/ 文件名通配（*.log）/ 目录名整体跳过（node_modules）
function Test-Excluded([string]$Rel) {
    $pats = Get-Excludes
    if ($pats.Count -eq 0) { return $false }
    $relN = $Rel.Replace('/', '\')
    $name = Split-Path $relN -Leaf
    foreach ($p in $pats) {
        $pat = ([string]$p).Trim().Replace('/', '\').Trim('\')
        if ($pat -eq '') { continue }
        if ($relN -like $pat -or $relN -like "$pat\*" -or $name -like $pat) { return $true }
        foreach ($seg in $relN.Split('\')) { if ($seg -eq $pat) { return $true } }
    }
    return $false
}

# 解析配置；缺项时在控制台询问（-NonInteractive 则用默认值，不询问）
function Resolve-Config([switch]$NonInteractive) {
    $config = Read-Config
    if ($script:RepoUrl -eq '' -and $config -and $config.RepoUrl) { $script:RepoUrl = $config.RepoUrl }
    if ($script:IntervalMinutes -le 0) {
        if ($config -and $config.IntervalMinutes) { $script:IntervalMinutes = [int]$config.IntervalMinutes }
        elseif ($config -and $config.IntervalHours) { $script:IntervalMinutes = [int]$config.IntervalHours * 60 }
    }

    $needAsk = (-not (Test-RepoUrl $script:RepoUrl)) -or ($script:IntervalMinutes -le 0)
    if ($needAsk -and ($DryRun -or $NonInteractive)) {
        if (-not (Test-RepoUrl $script:RepoUrl)) { $script:RepoUrl = $DefaultRepo }
        if ($script:IntervalMinutes -le 0) { $script:IntervalMinutes = 720 }
        return $true
    }
    if ($needAsk) {
        Write-Host '===== 首次运行，请先设置同步配置 =====' -ForegroundColor Cyan
        while (-not (Test-RepoUrl $script:RepoUrl)) {
            $input1 = Read-Host '请输入要同步的 GitHub 仓库地址（格式：https://github.com/用户名/仓库名）'
            $script:RepoUrl = $input1.Trim()
            if ($script:RepoUrl -eq '') {
                Write-Host '仓库地址不能为空，请重新输入。' -ForegroundColor Yellow
            } elseif (-not (Test-RepoUrl $script:RepoUrl)) {
                Write-Host "仓库地址格式不正确，必须形如 https://github.com/用户名/仓库名（可用 .git 结尾）。当前输入：$($script:RepoUrl)" -ForegroundColor Yellow
            }
        }
        while ($script:IntervalMinutes -le 0) {
            $input2 = Read-Host '多久同步一次？请输入分钟数'
            $m = 0
            if ([int]::TryParse($input2.Trim(), [ref]$m) -and $m -gt 0) { $script:IntervalMinutes = $m; break }
            Write-Host '请输入大于 0 的整数分钟数。' -ForegroundColor Yellow
        }
        Save-Config
        Write-Host "配置已保存到 $ConfigFile，之后启动不再询问。" -ForegroundColor Green
    }
    if (-not (Test-RepoUrl $script:RepoUrl)) { throw "仓库地址格式不正确：$script:RepoUrl （应形如 https://github.com/用户名/仓库名）" }
    if ($script:IntervalMinutes -le 0) { $script:IntervalMinutes = 720 }
    return $true
}

# 缓存目录按仓库地址区分，换仓库时不会误用旧缓存
function Set-RepoHash {
    $script:RepoHash = [System.BitConverter]::ToString(
        [System.Security.Cryptography.SHA1]::Create().ComputeHash(
            [System.Text.Encoding]::UTF8.GetBytes($script:RepoUrl))).Replace('-', '').Substring(0, 8)
}

# ============================================================================
# 同步内核
# ============================================================================
function Write-Log([string]$msg, [string]$color = 'Gray') {
    $line = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $msg
    Write-Host $line -ForegroundColor $color
    Add-Content -LiteralPath $LogFile -Value $line -Encoding UTF8
}

# 托盘是隐藏运行的，出错时没有控制台可见，单独记一份错误日志便于排查
function Write-TrayError([string]$msg) {
    try {
        Add-Content -LiteralPath (Join-Path $Root 'InfiniteGithubUpstreamSync-error.log') `
            -Value ('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $msg) -Encoding UTF8
    } catch {}
}

function Update-UpstreamCache {
    $CacheDir = Join-Path $env:TEMP ('HungerCat-upstream-cache-' + $script:RepoHash)
    $ok = $false
    if (Test-Path -LiteralPath (Join-Path $CacheDir '.git')) {
        git -C $CacheDir fetch origin --quiet 2>$null
        if ($LASTEXITCODE -eq 0) {
            git -C $CacheDir pull --ff-only --quiet 2>$null
            if ($LASTEXITCODE -eq 0) { $ok = $true }
        }
    }
    if (-not $ok) {
        if (Test-Path -LiteralPath $CacheDir) { Remove-Item -LiteralPath $CacheDir -Recurse -Force }
        Write-Log '正在从 GitHub 克隆最新代码到缓存...' 'Cyan'
        git clone --depth 1 $RepoUrl $CacheDir --quiet
        if ($LASTEXITCODE -ne 0) { throw 'git clone 失败，请检查网络连接或仓库地址' }
    }
    $script:CacheDir = (Get-Item -LiteralPath $CacheDir).FullName
}

function Invoke-Sync {
    Update-UpstreamCache

    $srcLen    = $script:CacheDir.TrimEnd('\').Length
    $gitDir    = (Join-Path $script:CacheDir '.git') + '\'
    $stats     = @{ Overwrite = 0; New = 0; Same = 0; Skipped = 0 }
    $backupDir = $null

    Get-ChildItem -LiteralPath $script:CacheDir -Recurse -File -Force | Where-Object {
        -not $_.FullName.StartsWith($gitDir, [StringComparison]::OrdinalIgnoreCase)
    } | ForEach-Object {
        $rel = $_.FullName.Substring($srcLen + 1)

        # 命中排除规则的上游文件：不覆盖、不新增，原样保留本地文件
        if (Test-Excluded $rel) { $stats.Skipped++; return }

        $dst = Join-Path $Root $rel

        if (Test-Path -LiteralPath $dst) {
            $dstItem = Get-Item -LiteralPath $dst -Force
            $same = ($_.Length -eq $dstItem.Length)
            if ($same) {
                $h1 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA1).Hash
                $h2 = (Get-FileHash -LiteralPath $dst -Algorithm SHA1).Hash
                $same = ($h1 -eq $h2)
            }
            if ($same) {
                $stats.Same++
            } else {
                $stats.Overwrite++
                Write-Log "覆盖: $rel" 'Yellow'
                if (-not $DryRun) {
                    if (-not $backupDir) { $backupDir = Join-Path $BackupRoot (Get-Date -Format 'yyyyMMdd-HHmmss') }
                    $bk = Join-Path $backupDir $rel
                    New-Item -ItemType Directory -Path (Split-Path $bk -Parent) -Force | Out-Null
                    Copy-Item -LiteralPath $dst -Destination $bk -Force
                    Copy-Item -LiteralPath $_.FullName -Destination $dst -Force
                }
            }
        } else {
            $stats.New++
            Write-Log "新增: $rel" 'Green'
            if (-not $DryRun) {
                New-Item -ItemType Directory -Path (Split-Path $dst -Parent) -Force | Out-Null
                Copy-Item -LiteralPath $_.FullName -Destination $dst -Force
            }
        }
    }

    if ($DryRun) {
        Write-Log ("[试运行] 将覆盖 {0} 个文件，新增 {1} 个文件，{2} 个文件无变化，排除 {3} 个（未做任何实际修改）" -f $stats.Overwrite, $stats.New, $stats.Same, $stats.Skipped) 'Magenta'
    } else {
        Write-Log ("同步完成：覆盖 {0} 个，新增 {1} 个，无变化 {2} 个，排除 {3} 个" -f $stats.Overwrite, $stats.New, $stats.Same, $stats.Skipped) 'Cyan'
        if ($backupDir) { Write-Log "旧文件已备份到: $backupDir" 'Cyan' }
        if (Test-Path -LiteralPath $BackupRoot) {
            Get-ChildItem -LiteralPath $BackupRoot -Directory | Sort-Object Name -Descending |
                Select-Object -Skip 10 | Remove-Item -Recurse -Force
        }
    }
}

function Invoke-SyncOnce {
    Resolve-Config -NonInteractive | Out-Null
    Set-RepoHash
    Write-Log "===== 自动同步启动 ===== 仓库: $script:RepoUrl 间隔: $script:IntervalMinutes 分钟$(if ($DryRun) { ' [试运行模式]' })" 'Cyan'
    try { Invoke-Sync } catch { Write-Log ('出错: ' + $_.Exception.Message) 'Red' }
}

# ============================================================================
# 检测更新：① 上游仓库是否有新提交  ② 本工具自身是否有新版本
# ============================================================================
function Get-ToolRepoUrl {
    # 内置固定值，不通过配置/界面暴露，避免用户改坏导致更新失败
    return $ToolRepo
}

# 把 "v1.2.3-dev1" 这类 tag 解析成可比较的版本号
function ConvertTo-VersionNum([string]$v) {
    if ([string]::IsNullOrWhiteSpace($v)) { return $null }
    $m = [regex]::Match(($v.Trim() -replace '^[vV]', ''), '\d+(\.\d+){0,3}')
    if (-not $m.Success) { return $null }
    try { return [version]$m.Value } catch { return $null }
}

# 只 fetch 不 pull：检查上游是否比本地缓存新
function Test-UpstreamUpdate {
    $r = @{ Error = ''; Files = @(); Branch = '' }
    $cache = Join-Path $env:TEMP ('HungerCat-upstream-cache-' + $script:RepoHash)
    if (-not (Test-Path -LiteralPath (Join-Path $cache '.git'))) {
        $r.Error = '本地还没有同步缓存，请先同步一次再检测'
        return $r
    }
    git -C $cache fetch origin --quiet 2>$null
    if ($LASTEXITCODE -ne 0) { $r.Error = '无法连接 GitHub（fetch 失败，请检查网络）'; return $r }
    $r.Branch = (& git -C $cache rev-parse --abbrev-ref HEAD).Trim()
    if ($r.Branch -eq 'HEAD' -or $r.Branch -eq '') { $r.Branch = 'main' }   # 极端情况（detached）
    $local  = (& git -C $cache rev-parse HEAD).Trim()
    $remote = (& git -C $cache rev-parse "origin/$($r.Branch)" 2>$null).Trim()
    if (-not $remote) { $r.Error = "找不到远程分支 origin/$($r.Branch)"; return $r }
    if ($local -eq $remote) { return $r }
    $r.Files = @(& git -C $cache diff --name-only $local $remote 2>$null | Where-Object { $_ -ne '' })
    return $r
}

# 查本工具在 GitHub 上的最新 release / tag，并与 $ToolVersion 比较
function Test-ToolUpdate {
    $r = @{ Error = ''; Latest = ''; Url = ''; HasNew = $false }
    $url = Get-ToolRepoUrl
    if (-not (Test-RepoUrl $url)) { $r.Error = "内置更新源格式不正确：$url"; return $r }

    $slug = $url.Trim().TrimEnd('/') -replace '^https://github\.com/', '' -replace '\.git$', ''
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

# 控制台展示检测结果（菜单选项 9）
function Invoke-CheckUpdate {
    Resolve-Config -NonInteractive | Out-Null
    Set-RepoHash
    Write-Host ''
    Write-Host ('本工具版本：' + $ToolVersion) -ForegroundColor DarkGray

    Write-Host '上游仓库：' -NoNewline
    $u = Test-UpstreamUpdate
    if ($u.Error) { Write-Host $u.Error -ForegroundColor Yellow }
    elseif ($u.Files.Count -eq 0) { Write-Host '已是最新，没有新提交。' -ForegroundColor Green }
    else {
        Write-Host ('有更新，' + $u.Files.Count + ' 个文件将变化（分支 ' + $u.Branch + '）。') -ForegroundColor Yellow
        $u.Files | Select-Object -First 8 | ForEach-Object { Write-Host "    $_" }
        if ($u.Files.Count -gt 8) { Write-Host ("    ……其余 {0} 个" -f ($u.Files.Count - 8)) }
    }

    Write-Host '工具更新：' -NoNewline
    $t = Test-ToolUpdate
    if ($t.Error) { Write-Host $t.Error -ForegroundColor DarkGray }
    elseif ($t.HasNew) {
        Write-Host ('发现新版本 ' + $t.Latest + '（当前 ' + $ToolVersion + '）') -ForegroundColor Cyan
        Write-Host ('    下载页：' + $t.Url)
    } else {
        Write-Host ('已是最新（' + $t.Latest + '）') -ForegroundColor Green
    }

    $sum = '上游：' + $(if ($u.Error) { $u.Error } elseif ($u.Files.Count -eq 0) { '已是最新' } else { "$($u.Files.Count) 个文件有更新" }) +
        '；工具：' + $(if ($t.Error) { '未配置/不可用' } elseif ($t.HasNew) { "发现新版 $($t.Latest)" } else { '已是最新' })
    Write-Log ('检测更新 —— ' + $sum) 'DarkGray'
    return $t
}

# 把下载到临时目录的脚本里的版本号改成 tag 值（tag 是权威版本，避免脚本里写死的旧版本号）
function Set-ScriptVersion([string]$Path, [string]$Version) {
    $txt = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
    $new = [regex]::Replace($txt, '^\$ToolVersion = ''[^'']*''', {
            param($m) '$ToolVersion = ''' + $Version + ''''
        }, 'Multiline')
    if ($new -ne $txt) { [System.IO.File]::WriteAllText($Path, $new, (New-Object System.Text.UTF8Encoding $true)) }
}

# 工具自我更新：从工具仓库对应 tag 下载各文件 → 校验 → 备份 → 替换 → 重启托盘
# 返回 @{ Ok=bool; Msg=string }
function Invoke-ToolUpdate {
    $t = Test-ToolUpdate
    if ($t.Error) { return @{ Ok = $false; Msg = '无法检查更新：' + $t.Error } }
    if (-not $t.HasNew) { return @{ Ok = $false; Msg = '已是最新版本（' + $t.Latest + '），无需更新' } }

    $slug = (Get-ToolRepoUrl).Trim().TrimEnd('/') -replace '^https://github\.com/', '' -replace '\.git$', ''
    $tag = $t.Latest
    $ver = ([string]$tag).TrimStart('v', 'V')
    $ts = Get-Date -Format 'yyyyMMdd-HHmmss'
    $tmp = Join-Path $env:TEMP ('igus-update-' + $ts)
    New-Item -ItemType Directory -Path $tmp -Force | Out-Null

    # 1) 先全部下载到临时目录并校验，任何一个不合规都不动本地文件
    $got = @(); $fail = @()
    foreach ($f in $ToolFiles) {
        $dst = Join-Path $tmp $f
        $uri = 'https://raw.githubusercontent.com/' + $slug + '/' + $tag + '/' + [uri]::EscapeDataString($f)
        try {
            Invoke-WebRequest -Uri $uri -OutFile $dst -UseBasicParsing -TimeoutSec 40
            if ((Get-Item -LiteralPath $dst).Length -lt 100) { $fail += $f; continue }
            if ($f -like '*.ps1') {
                Set-ScriptVersion $dst $ver
                $errs = $null
                $null = [System.Management.Automation.Language.Parser]::ParseFile($dst, [ref]$null, [ref]$errs)
                if ($errs -and $errs.Count -gt 0) { $fail += ($f + '(语法检查未通过)'); continue }
            }
            $got += $f
        } catch { $fail += $f }
    }
    if ($got -notcontains 'InfiniteGithubUpstreamSync.ps1') {
        Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
        return @{ Ok = $false; Msg = '未能下载到主脚本，已放弃更新（可能 tag 与仓库文件不一致）。失败项：' + ($fail -join '、') }
    }

    # 2) 备份现有文件
    $backup = Join-Path $Root ('.sync-script-backup\pre-selfupdate-' + $ts)
    New-Item -ItemType Directory -Path $backup -Force | Out-Null
    foreach ($f in $got) {
        $cur = Join-Path $Root $f
        if (Test-Path -LiteralPath $cur) { Copy-Item -LiteralPath $cur -Destination (Join-Path $backup $f) -Force }
    }

    # 3) 替换；中途失败则回滚
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

    # 4) 托盘在跑就重启它，让新代码生效
    $restarted = Restart-TrayIfRunning

    $msg = '已更新到 ' + $tag + '（本地版本号已同步为 ' + $ver + '），替换 ' + $got.Count + ' 个文件：' + ($got -join '、')
    if ($fail.Count -gt 0) { $msg += '；未获取：' + ($fail -join '、') }
    if ($restarted) { $msg += '；托盘已重启' }
    $msg += '；原文件备份在 ' + $backup
    Write-Log ('工具更新 —— ' + $msg) 'Cyan'
    return @{ Ok = $true; Msg = $msg }
}

# ============================================================================
# 进程 / 自启
# ============================================================================
function Get-SyncProcess {
    Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.ProcessId -ne $PID -and $_.CommandLine -and $_.CommandLine -like '*InfiniteGithubUpstreamSync.ps1*' }
}

# 仅托盘常驻进程（排除 -Once 一次性同步、-Menu 菜单、-GUI 窗口）
function Get-TrayProcess {
    @(Get-SyncProcess | Where-Object {
            $_.CommandLine -notlike '*-Once*' -and $_.CommandLine -notlike '*-Menu*' -and $_.CommandLine -notlike '*-GUI*'
        })
}

function Test-Autostart { [bool](Get-ItemProperty -Path $RunKey -Name $RunValue -ErrorAction SilentlyContinue) }

# 托盘在运行就重启（更新完代码后用），返回是否重启了
function Restart-TrayIfRunning {
    $procs = Get-TrayProcess
    if (-not $procs -or $procs.Count -eq 0) { return $false }
    foreach ($p in $procs) { Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue }
    Start-Sleep -Milliseconds 900
    if (Test-Path -LiteralPath $VbsPath) {
        Start-Process -FilePath (Join-Path $env:WINDIR 'explorer.exe') -ArgumentList ('"' + $VbsPath + '"')
    } else {
        Start-Process -FilePath 'powershell.exe' -WindowStyle Hidden -ArgumentList ('-NoProfile -ExecutionPolicy Bypass -File "{0}"' -f $PSCommandPath)
    }
    return $true
}

function Show-Status {
    $procs = Get-TrayProcess
    Write-Host ('程序进程: ' + $(if ($procs) { "运行中（PID $($procs.ProcessId -join ', ')）" } else { '未运行' }))
    Write-Host ('开机自启: ' + $(if (Test-Autostart) { '已设置' } else { '未设置' }))
    Write-Host ('仓库地址: ' + $(try { (Read-Config).RepoUrl } catch { '未配置' }))
    Write-Host ('同步间隔: ' + (Get-IntervalMinutes) + ' 分钟')
    $ex = Get-Excludes
    Write-Host ('排除规则: ' + $(if ($ex.Count -gt 0) { ($ex -join ', ') } else { '（无）' }))
    if (Test-Path -LiteralPath $LogFile) {
        Write-Host '最近日志:'
        Get-Content -LiteralPath $LogFile -Tail 3 | ForEach-Object { Write-Host "  $_" }
    }
}

function Install-Autostart {
    if (-not (Test-Path -LiteralPath $ConfigFile)) {
        Write-Host '首次使用，先完成设置（会立即同步一次）...' -ForegroundColor Cyan
        Resolve-Config | Out-Null
        Set-RepoHash
        Write-Log "===== 自动同步启动 ===== 仓库: $script:RepoUrl 间隔: $script:IntervalMinutes 分钟" 'Cyan'
        try { Invoke-Sync } catch { Write-Log ('出错: ' + $_.Exception.Message) 'Red' }
        if (-not (Test-Path -LiteralPath $ConfigFile)) { throw '首次设置未完成（未生成配置文件）' }
    }
    $vbs = @"
' 由 InfiniteGithubUpstreamSync.ps1 自动生成，请勿手动修改
Set sh = CreateObject("WScript.Shell")
sh.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ""$PSCommandPath""", 0, False
"@
    [System.IO.File]::WriteAllText($VbsPath, $vbs, [System.Text.Encoding]::Default)
    if (-not (Test-Path -LiteralPath $RunKey)) { New-Item -Path $RunKey -Force | Out-Null }
    Set-ItemProperty -Path $RunKey -Name $RunValue -Value ('wscript.exe //B //Nologo "{0}"' -f $VbsPath)
    Write-Host "已设置开机自启（$RunValue）。" -ForegroundColor Green

    if (Get-ScheduledTask -TaskName 'HungerCat-Upstream-Sync' -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName 'HungerCat-Upstream-Sync' -Confirm:$false
        Write-Host '已移除旧的计划任务，改由本程序接管。' -ForegroundColor Yellow
    }
    foreach ($old in @('sync-hidden.vbs', 'tray-hidden.vbs', 'InfiniteGithubUpstreamSync-hidden.vbs.old')) {
        $p = Join-Path $Root $old
        if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Force }
    }

    $tray = Get-TrayProcess
    if (-not $tray) {
        Start-Process -FilePath 'wscript.exe' -ArgumentList ('//B //Nologo "{0}"' -f $VbsPath)
    }
    Write-Host ('InfiniteGithubUpstreamSync 已启动，间隔 ' + (Get-IntervalMinutes) + ' 分钟。右键托盘图标即可操作。') -ForegroundColor Cyan
}

function Remove-Autostart {
    Get-SyncProcess | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    Remove-ItemProperty -Path $RunKey -Name $RunValue -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $VbsPath) { Remove-Item -LiteralPath $VbsPath -Force }
    Write-Host '已停止 InfiniteGithubUpstreamSync，并取消开机自启。' -ForegroundColor Green
}

function Open-LiveLog {
    Start-Process -FilePath 'powershell.exe' -ArgumentList (
        '-NoExit -NoProfile -Command "Get-Content -LiteralPath ''{0}'' -Wait -Tail 40"' -f $LogFile)
}

# ============================================================================
# 命令行版：排除规则管理（菜单选项 7）
# ============================================================================
function Edit-Excludes {
    while ($true) {
        $ex = Get-Excludes
        Write-Host ''
        Write-Host '----- 排除规则（命中的上游文件不会被同步） -----' -ForegroundColor Cyan
        if ($ex.Count -eq 0) { Write-Host '  （当前为空）' }
        for ($i = 0; $i -lt $ex.Count; $i++) { Write-Host ("  [{0}] {1}" -f ($i + 1), $ex[$i]) }
        Write-Host '  输入新规则直接添加（如 node_modules、*.log、build\*）；'
        Write-Host '  输入 del 序号 删除（如 del 2）；输入 0 返回。'
        $in = (Read-Host '排除规则操作').Trim()
        if ($in -eq '0') { return }
        if ($in -match '^del\s+(\d+)$') {
            $idx = [int]$Matches[1] - 1
            if ($idx -ge 0 -and $idx -lt $ex.Count) {
                $ex = @($ex | Where-Object { $_ -ne $ex[$idx] })
                # 上面的 Where 在有重复值时会误删同值项，改用索引重组
                $ex = @(); $cur = Get-Excludes
                for ($i = 0; $i -lt $cur.Count; $i++) { if ($i -ne $idx) { $ex += $cur[$i] } }
                Resolve-Config -NonInteractive | Out-Null
                Save-Config $ex
                Write-Host '已删除。' -ForegroundColor Green
            } else { Write-Host '序号无效。' -ForegroundColor Yellow }
            continue
        }
        if ($in -ne '') {
            Resolve-Config -NonInteractive | Out-Null
            Save-Config (@(Get-Excludes) + $in)
            Write-Host "已添加排除规则：$in（下次同步生效）" -ForegroundColor Green
        }
    }
}

# ============================================================================
# 打开 GUI（独立文件，隐藏 PowerShell 窗口启动）
# ============================================================================
function Open-Gui {
    $gui = Join-Path $Root 'InfiniteGithubUpstreamSync-GUI.ps1'
    if (-not (Test-Path -LiteralPath $gui)) {
        throw '未找到 InfiniteGithubUpstreamSync-GUI.ps1，请确认它与主脚本在同一目录。'
    }
    Start-Process -FilePath 'powershell.exe' -WindowStyle Hidden -ArgumentList (
        '-STA -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}"' -f $gui)
}

# ============================================================================
# 模式分派
# ============================================================================
if ($Status) { Show-Status; return }
if ($Exit) {
    $procs = Get-TrayProcess
    if ($procs) {
        $procs | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
        Write-Host '已退出托盘程序（开机自启保持不变）。' -ForegroundColor Green
    } else {
        Write-Host '托盘程序未在运行。' -ForegroundColor Yellow
    }
    return
}
if ($Remove) { Remove-Autostart; return }
if ($Install) { Install-Autostart; return }
if ($Once -or $DryRun) { Invoke-SyncOnce; return }
if ($CheckUpdate) { Invoke-CheckUpdate | Out-Null; return }
if ($Update) {
    $r = Invoke-ToolUpdate
    Write-Host $r.Msg -ForegroundColor $(if ($r.Ok) { 'Green' } else { 'Yellow' })
    return
}

if ($Menu) {
    # 执行完任一操作后自动清屏并重新输出主菜单；
    # 对输出内容重要的操作（状态/同步），先按回车再清屏，避免结果一闪而过
    $pauseBack = { Write-Host ''; [void](Read-Host '按回车返回主菜单') }
    while ($true) {
        Clear-Host
        Write-Host ''
        Write-Host '==============================================' -ForegroundColor Cyan
        Write-Host '   InfiniteGithubUpstreamSync' -ForegroundColor Cyan
        Write-Host '==============================================' -ForegroundColor Cyan
        Write-Host '   1. 启动托盘 + 设置开机自启'
        Write-Host '   2. 查看状态'
        Write-Host '   3. 停止托盘 + 取消开机自启'
        Write-Host '   4. 立即同步一次'
        Write-Host '   5. 查看实时日志'
        Write-Host '   6. 退出托盘（保留开机自启）'
        Write-Host '   7. 管理排除规则'
        Write-Host '   8. 打开 GUI 窗口'
        Write-Host '   9. 检测更新（上游提交 / 工具新版）'
        Write-Host '  10. 更新到最新版（下载并替换本工具）'
        Write-Host '   0. 关闭菜单'
        Write-Host '==============================================' -ForegroundColor Cyan
        $choice = Read-Host '请输入序号'
        Clear-Host
        switch ($choice.Trim()) {
            '1' {
                try { Install-Autostart } catch { Write-Host $_.Exception.Message -ForegroundColor Red }
                Write-Host ''
                Write-Host '托盘已在后台常驻，本窗口即将关闭（关掉它不影响同步）。' -ForegroundColor Green
                return
            }
            '2' { Show-Status; & $pauseBack }
            '3' { Remove-Autostart; Start-Sleep -Milliseconds 1200 }
            '4' {
                Resolve-Config | Out-Null
                Set-RepoHash
                try { Invoke-Sync } catch { Write-Host ('出错: ' + $_.Exception.Message) -ForegroundColor Red }
                & $pauseBack
            }
            '5' { Open-LiveLog }
            '6' {
                $procs = Get-TrayProcess
                if ($procs) {
                    $procs | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
                    Write-Host '已退出托盘程序（开机自启保持不变）。' -ForegroundColor Green
                } else { Write-Host '托盘程序未在运行。' -ForegroundColor Yellow }
                Start-Sleep -Milliseconds 1200
            }
            '7' { Edit-Excludes }
            '8' { try { Open-Gui } catch { Write-Host $_.Exception.Message -ForegroundColor Red; Start-Sleep -Milliseconds 1500 } }
            '9' {
                $t = Invoke-CheckUpdate
                if ($t -and $t.HasNew) {
                    $ans = Read-Host ("发现新版本 {0}（当前 {1}），是否立即更新？(y/N)" -f $t.Latest, $ToolVersion)
                    if ($ans -match '^(y|Y)') {
                        $r = Invoke-ToolUpdate
                        Write-Host $r.Msg -ForegroundColor $(if ($r.Ok) { 'Green' } else { 'Yellow' })
                        if ($r.Ok) { Write-Host '更新后请重新打开菜单或托盘，以使用新版本。' -ForegroundColor DarkGray }
                    }
                }
                & $pauseBack
            }
            '10' {
                $r = Invoke-ToolUpdate
                Write-Host $r.Msg -ForegroundColor $(if ($r.Ok) { 'Green' } else { 'Yellow' })
                & $pauseBack
            }
            '0' { Clear-Host; return }
            default {
                Write-Host '无效输入，请重新选择。' -ForegroundColor Yellow
                Start-Sleep -Milliseconds 1200
            }
        }
    }
}

# ============================================================================
# 托盘本体（默认模式）
# ============================================================================
# 兜底：托盘隐藏运行，任何未处理异常都记进错误日志，便于排查
trap {
    Write-TrayError ("托盘异常: " + $_.Exception.Message + "`n" + $_.ScriptStackTrace)
    break
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$created = $false
$mutex = New-Object System.Threading.Mutex($true, $MutexName, [ref]$created)
if (-not $created) { return }

Resolve-Config -NonInteractive | Out-Null
Set-RepoHash

$Interval = $script:IntervalMinutes

$notify = New-Object System.Windows.Forms.NotifyIcon
try { $notify.Icon = New-Object System.Drawing.Icon($IconPath) } catch { $notify.Icon = [System.Drawing.SystemIcons]::Application }
$notify.Text = 'InfiniteGithubUpstreamSync'
$notify.Visible = $true

$script:syncProc = $null
$script:nextRun  = (Get-Date).AddMinutes($Interval)

function Show-Tip([string]$msg) {
    $notify.BalloonTipTitle = 'InfiniteGithubUpstreamSync'
    $notify.BalloonTipText = $msg
    $notify.ShowBalloonTip(4000)
}

function Start-SyncNow {
    if ($script:syncProc -and -not $script:syncProc.HasExited) {
        Show-Tip '正在同步中，请稍候…'
        return
    }
    $arg = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}" -Once' -f $PSCommandPath
    try {
        $script:syncProc = Start-Process -FilePath 'powershell.exe' -ArgumentList $arg -WindowStyle Hidden -PassThru
        $poll.Enabled = $true
    } catch {
        Show-Tip ('启动同步失败：' + $_.Exception.Message)
    }
}

# 注意：变量名不能叫 $menu —— 会与 -Menu 开关参数同名冲突，导致右键菜单异常
$ctxMenu = New-Object System.Windows.Forms.ContextMenuStrip

function Invoke-MenuAction([scriptblock]$Action, [string]$Name) {
    try { & $Action }
    catch {
        # 出错也要让用户看得见，而不是右键"没反应"
        Write-TrayError "菜单项[$Name]出错: $($_.Exception.Message)"
        [void][System.Windows.Forms.MessageBox]::Show("操作失败：$($_.Exception.Message)", 'InfiniteGithubUpstreamSync',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
}

$itemSync = $ctxMenu.Items.Add('立即同步一次')
$itemSync.add_Click({ Invoke-MenuAction { Start-SyncNow } '立即同步一次' })

$itemStatus = $ctxMenu.Items.Add('查看状态')
$itemStatus.add_Click({
        Invoke-MenuAction {
            $last = '尚未运行'
            if (Test-Path -LiteralPath $LogFile) { $last = (Get-Content -LiteralPath $LogFile -Tail 1) }
            Show-Tip ("仓库：" + $script:RepoUrl + "`n间隔：" + $Interval + " 分钟`n下次：" +
                $script:nextRun.ToString('MM-dd HH:mm') + "`n最近：" + $last)
        } '查看状态'
    })

$itemLog = $ctxMenu.Items.Add('查看实时日志')
$itemLog.add_Click({ Invoke-MenuAction { Open-LiveLog } '查看实时日志' })

$itemGui = $ctxMenu.Items.Add('打开 GUI 设置窗口')
$itemGui.add_Click({ Invoke-MenuAction { Open-Gui } '打开 GUI 设置窗口' })

$itemCheck = $ctxMenu.Items.Add('检测更新')
$itemCheck.add_Click({
        Invoke-MenuAction {
            Show-Tip '正在检测更新，请稍候…'
            Set-RepoHash
            $u = Test-UpstreamUpdate
            $t = Test-ToolUpdate
            $txt = '上游：' + $(if ($u.Error) { $u.Error } elseif ($u.Files.Count -eq 0) { '已是最新' } else { "$($u.Files.Count) 个文件有更新" })
            if ($t.Error) { $txt += "`n工具：未配置更新源" }
            elseif ($t.HasNew) { $txt += "`n工具：发现新版 $($t.Latest)"; Write-TrayError ('发现工具新版: ' + $t.Latest + ' ' + $t.Url) }
            else { $txt += "`n工具：已是最新" }
            Show-Tip $txt
        } '检测更新'
    })

$itemConsole = $ctxMenu.Items.Add('回到控制台')
$itemConsole.add_Click({
        Invoke-MenuAction {
            $bat = Join-Path $Root '双击启动同步.bat'
            if (Test-Path -LiteralPath $bat) { Start-Process -FilePath 'cmd.exe' -ArgumentList ('/c ""{0}""' -f $bat) }
            else { Open-LiveLog }
        } '回到控制台'
    })

[void]$ctxMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))

$itemExit = $ctxMenu.Items.Add('退出')
$itemExit.add_Click({
        Invoke-MenuAction {
            $notify.Visible = $false
            $notify.Dispose()
            [System.Windows.Forms.Application]::Exit()
        } '退出'
    })

$notify.ContextMenuStrip = $ctxMenu

$notify.add_MouseDoubleClick({ Open-LiveLog })

$poll = New-Object System.Windows.Forms.Timer
$poll.Interval = 2000
$poll.add_Tick({
        if ($script:syncProc -and $script:syncProc.HasExited) {
            $poll.Enabled = $false
            $code = $script:syncProc.ExitCode
            $script:syncProc = $null
            if ($code -eq 0) {
                $script:nextRun = (Get-Date).AddMinutes($Interval)
                Show-Tip '同步完成。'
            } else {
                Show-Tip "同步失败（退出码 $code），右键 → 查看实时日志 了解详情。"
            }
        }
    })

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = [int][Math]::Min($Interval * 60 * 1000, [int]::MaxValue)
$timer.add_Tick({ Start-SyncNow })
$timer.Enabled = $true

# 启动 5 秒后先同步一次（补做关机期间错过的同步）
$kick = New-Object System.Windows.Forms.Timer
$kick.Interval = 5000
$kick.add_Tick({ $kick.Enabled = $false; Start-SyncNow })
$kick.Enabled = $true

Show-Tip ('InfiniteGithubUpstreamSync 已就绪，每 ' + $Interval + ' 分钟同步一次。')
[System.Windows.Forms.Application]::Run()
