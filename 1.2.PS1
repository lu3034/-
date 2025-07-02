# 解决权限和执行策略问题
try {
    $ErrorActionPreference = "Stop"
    
    # 检查管理员权限
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
        $warningMsg = @"
警告：需要管理员权限运行！
-------------------------
1. 请右键点击此脚本
2. 选择"以管理员身份运行"
3. 然后按确定继续
"@
        Add-Type -AssemblyName Microsoft.VisualBasic
        $result = [Microsoft.VisualBasic.Interaction]::MsgBox($warningMsg, "OKOnly,SystemModal,Critical", "权限错误")
        exit
    }
    
    # 设置执行策略（只在当前进程有效）
    $null = Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
}
catch {
    Write-Host "初始化错误: $_"
    Read-Host "按回车键退出"
    exit
}

# 加载所需程序集
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
}
catch {
    $errorMsg = @"
程序集加载失败！
--------------------
错误信息: $_
--------------------
可能是.NET框架损坏或缺失组件。
请尝试：
1. 运行Windows更新
2. 安装最新.NET框架
3. 运行系统文件检查: sfc /scannow
"@
    Add-Type -AssemblyName Microsoft.VisualBasic
    $result = [Microsoft.VisualBasic.Interaction]::MsgBox($errorMsg, "OKOnly,SystemModal,Critical", "严重错误")
    exit
}

# 创建主窗体
$form = New-Object System.Windows.Forms.Form
$form.Text = "文件备份监控器V1.2"
$form.Size = New-Object System.Drawing.Size(750, 650) # 增大宽度以显示完整信息
$form.StartPosition = "CenterScreen"
$form.MinimizeBox = $true
$form.MaximizeBox = $false
$form.FormBorderStyle = "FixedDialog"
$form.BackColor = "#F0F0F0"

# 设置图标
try {
    $iconPath = "$env:WINDIR\System32\SHELL32.dll"
    $iconIndex = 15
    $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
    $form.Icon = $icon
}
catch {
    # 如果图标加载失败继续运行
}

# 标签和输入框 - 目标文件夹
$lblTarget = New-Object System.Windows.Forms.Label
$lblTarget.Location = New-Object System.Drawing.Point(20, 20)
$lblTarget.Size = New-Object System.Drawing.Size(120, 20)
$lblTarget.Text = "目标文件夹:"
$form.Controls.Add($lblTarget)

$txtTarget = New-Object System.Windows.Forms.TextBox
$txtTarget.Location = New-Object System.Drawing.Point(140, 20)
$txtTarget.Size = New-Object System.Drawing.Size(450, 20) # 增加宽度
$txtTarget.Text = "$env:USERPROFILE\Desktop"
$form.Controls.Add($txtTarget)

# 浏览按钮
$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Location = New-Object System.Drawing.Point(600, 19)
$btnBrowse.Size = New-Object System.Drawing.Size(100, 23)
$btnBrowse.Text = "浏览..."
$btnBrowse.BackColor = "White"
$btnBrowse.Add_Click({
    try {
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "选择目标文件夹"
        $folderBrowser.SelectedPath = $txtTarget.Text
        if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtTarget.Text = $folderBrowser.SelectedPath
        }
    }
    catch {
        Write-Log "文件夹浏览失败: $_" "Red"
    }
})
$form.Controls.Add($btnBrowse)

# 标签和输入框 - 检测间隔
$lblInterval = New-Object System.Windows.Forms.Label
$lblInterval.Location = New-Object System.Drawing.Point(20, 60)
$lblInterval.Size = New-Object System.Drawing.Size(120, 20)
$lblInterval.Text = "检测间隔(秒):"
$form.Controls.Add($lblInterval)

$numInterval = New-Object System.Windows.Forms.NumericUpDown
$numInterval.Location = New-Object System.Drawing.Point(140, 60)
$numInterval.Size = New-Object System.Drawing.Size(100, 20)
$numInterval.Value = 1
$numInterval.Minimum = 1
$numInterval.Maximum = 10
$numInterval.Increment = 0.5
$form.Controls.Add($numInterval)

# 开始按钮
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Location = New-Object System.Drawing.Point(260, 60)
$btnStart.Size = New-Object System.Drawing.Size(100, 23)
$btnStart.Text = "开始监控"
$btnStart.BackColor = "#4CAF50"
$btnStart.ForeColor = "White"
$btnStart.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8.25, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnStart)

# 停止按钮
$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Location = New-Object System.Drawing.Point(370, 60)
$btnStop.Size = New-Object System.Drawing.Size(100, 23)
$btnStop.Text = "停止监控"
$btnStop.BackColor = "#F44336"
$btnStop.ForeColor = "White"
$btnStop.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8.25, [System.Drawing.FontStyle]::Bold)
$btnStop.Enabled = $false
$form.Controls.Add($btnStop)

# 重置统计按钮
$btnReset = New-Object System.Windows.Forms.Button
$btnReset.Location = New-Object System.Drawing.Point(480, 59)
$btnReset.Size = New-Object System.Drawing.Size(100, 23)
$btnReset.Text = "重置统计"
$btnReset.BackColor = "#FFC107"
$form.Controls.Add($btnReset)

# 大小显示格式下拉框
$lblSizeFormat = New-Object System.Windows.Forms.Label
$lblSizeFormat.Location = New-Object System.Drawing.Point(590, 60)
$lblSizeFormat.Size = New-Object System.Drawing.Size(120, 20)
$lblSizeFormat.Text = "大小显示:"
$form.Controls.Add($lblSizeFormat)

$cmbSizeFormat = New-Object System.Windows.Forms.ComboBox
$cmbSizeFormat.Location = New-Object System.Drawing.Point(650, 60)
$cmbSizeFormat.Size = New-Object System.Drawing.Size(80, 20)
$cmbSizeFormat.Items.AddRange(("自动", "原始字节", "KB", "MB"))
$cmbSizeFormat.SelectedIndex = 0
$form.Controls.Add($cmbSizeFormat)

# 统计面板
$statsPanel = New-Object System.Windows.Forms.Panel
$statsPanel.Location = New-Object System.Drawing.Point(20, 95)
$statsPanel.Size = New-Object System.Drawing.Size(710, 40) # 增加宽度
$statsPanel.BackColor = "#E8F5E9"
$statsPanel.BorderStyle = "FixedSingle"
$form.Controls.Add($statsPanel)

# 统计标签
$lblFiles = New-Object System.Windows.Forms.Label
$lblFiles.Location = New-Object System.Drawing.Point(15, 12)
$lblFiles.Size = New-Object System.Drawing.Size(120, 20)
$lblFiles.Text = "复制文件: 0"
$lblFiles.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$statsPanel.Controls.Add($lblFiles)

$lblFolders = New-Object System.Windows.Forms.Label
$lblFolders.Location = New-Object System.Drawing.Point(160, 12)
$lblFolders.Size = New-Object System.Drawing.Size(140, 20)
$lblFolders.Text = "复制文件夹: 0"
$lblFolders.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$statsPanel.Controls.Add($lblFolders)

$lblSize = New-Object System.Windows.Forms.Label
$lblSize.Location = New-Object System.Drawing.Point(320, 12)
$lblSize.Size = New-Object System.Drawing.Size(200, 20)
$lblSize.Text = "总占用空间: 0 B"
$lblSize.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$statsPanel.Controls.Add($lblSize)

$lblLastSize = New-Object System.Windows.Forms.Label
$lblLastSize.Location = New-Object System.Drawing.Point(530, 12)
$lblLastSize.Size = New-Object System.Drawing.Size(170, 20)
$lblLastSize.Text = "最新项目: N/A"
$lblLastSize.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$statsPanel.Controls.Add($lblLastSize)

# 日志框
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Location = New-Object System.Drawing.Point(20, 145)
$lblLog.Size = New-Object System.Drawing.Size(120, 20)
$lblLog.Text = "详细日志:"
$form.Controls.Add($lblLog)

$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 165)
$txtLog.Size = New-Object System.Drawing.Size(710, 380) # 增加宽度和高度
$txtLog.ReadOnly = $true
$txtLog.BackColor = "White"
$txtLog.ScrollBars = "Vertical"
$form.Controls.Add($txtLog)

# 状态栏
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusBar.Dock = "Bottom"
$statusBar.BackColor = "#E0E0E0"
$form.Controls.Add($statusBar)

# 状态标签
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "准备就绪"
$statusLabel.BorderSides = "Right"
$statusBar.Items.Add($statusLabel)

# 实时内存标签
$memoryLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$memoryLabel.Spring = $true
$memoryLabel.TextAlign = "MiddleRight"
$memoryLabel.Text = "内存: 计算中..."
$statusBar.Items.Add($memoryLabel)

# 进度条
$progressBar = New-Object System.Windows.Forms.ToolStripProgressBar
$progressBar.Style = "Marquee"
$progressBar.AutoSize = $false
$progressBar.Width = 150
$progressBar.Visible = $false
$statusBar.Items.Add($progressBar)

# 初始化全局变量
$global:isMonitoring = $false
$global:clipboardHistory = ""
$global:timer = $null

# 统计变量
$global:copiedFiles = 0
$global:copiedFolders = 0
$global:totalSize = 0
$global:lastItemSize = 0
$global:lastItemName = ""

# 格式化文件大小函数
function Format-Size {
    param(
        [long]$bytes,
        [int]$formatOption = 0 # 0=自动, 1=原始字节, 2=KB, 3=MB
    )
    
    if ($formatOption -eq 1) { 
        return "$bytes B"
    }
    elseif ($formatOption -eq 2) {
        return "{0:N2} KB" -f ($bytes / 1KB)
    }
    elseif ($formatOption -eq 3) {
        return "{0:N2} MB" -f ($bytes / 1MB)
    }
    # 自动格式
    elseif ($bytes -gt 1GB) {
        return "{0:N2} GB" -f ($bytes / 1GB)
    }
    elseif ($bytes -gt 1MB) {
        return "{0:N2} MB" -f ($bytes / 1MB)
    }
    elseif ($bytes -gt 1KB) {
        return "{0:N2} KB" -f ($bytes / 1KB)
    }
    else {
        return "$bytes B"
    }
}

# 更新内存显示函数
function Update-MemoryUsage {
    try {
        $process = Get-Process -Id $pid
        $memoryUsage = [math]::Round($process.WorkingSet64 / 1MB, 2)
        $memoryLabel.Text = "内存: $memoryUsage MB"
    }
    catch {
        $memoryLabel.Text = "内存: 读取失败"
    }
}

# 写入日志
function Write-Log {
    param(
        [string]$message,
        [string]$color = "Black",
        [switch]$isBold,
        [switch]$isSize
    )
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    
    # 获取大小显示选项
    $sizeFormat = $cmbSizeFormat.SelectedIndex
    
    # 处理消息中的大小信息
    if ($isSize) {
        $value = [regex]::Match($message, "\[size:(\d+)\]").Groups[1].Value
        if (-not [string]::IsNullOrEmpty($value)) {
            $bytes = [long]$value
            $sizeText = Format-Size -bytes $bytes -formatOption $sizeFormat
            $message = $message -replace "\[size:\d+\]", $sizeText
        }
    }
    
    $logEntry = "[$timestamp] $message"
    
    $txtLog.SelectionStart = $txtLog.TextLength
    $txtLog.SelectionLength = 0
    
    if ($isBold) {
        $txtLog.SelectionFont = New-Object System.Drawing.Font($txtLog.Font, [System.Drawing.FontStyle]::Bold)
    }
    
    if (-not [string]::IsNullOrEmpty($color) -and $color -ne "Black") {
        $txtLog.SelectionColor = $color
    }
    
    $txtLog.AppendText($logEntry + "`n")
    
    # 重置样式
    $txtLog.SelectionFont = $txtLog.Font
    $txtLog.SelectionColor = "Black"
    $txtLog.ScrollToCaret()
}

# 复制文件/文件夹并更新统计
function Copy-Resource {
    param(
        [string]$sourcePath,
        [string]$targetPath
    )
    
    try {
        # 检测源是文件还是文件夹
        if (Test-Path -Path $sourcePath -PathType Container) {
            $folderName = Split-Path $sourcePath -Leaf
            $destFolder = Join-Path -Path $targetPath -ChildPath $folderName
            
            # 如果目标文件夹不存在则创建
            if (-not (Test-Path -Path $targetPath)) {
                New-Item -ItemType Directory -Path $targetPath -Force | Out-Null
                Write-Log "创建目标文件夹: $targetPath" "Green"
            }
            
            # 复制文件夹前获取文件夹大小
            try {
                $folderSize = (Get-ChildItem -Path $sourcePath -Recurse | 
                    Where-Object { -not $_.PSIsContainer } | 
                    Measure-Object -Property Length -Sum).Sum
            }
            catch {
                $folderSize = 0
                Write-Log "警告: 无法计算文件夹大小 ($folderName). 可能部分文件无访问权限: $_" "DarkOrange"
            }
            
            # 复制文件夹
            Write-Log "正在复制文件夹: $folderName [size:$folderSize]" -color Blue -isSize
            Write-Log "├─ 文件数: $((Get-ChildItem -Path $sourcePath -Recurse -File | Measure-Object).Count)" "Blue"
            Write-Log "├─ 子文件夹数: $((Get-ChildItem -Path $sourcePath -Recurse -Directory | Measure-Object).Count)" "Blue"
            
            # 复制操作
            $null = New-Item -ItemType Directory -Path $destFolder -Force
            Copy-Item -Path "$sourcePath\*" -Destination $destFolder -Recurse -Force
            
            # 更新统计
            $global:copiedFolders++
            $global:totalSize += $folderSize
            $global:lastItemSize = $folderSize
            $global:lastItemName = $folderName
            
            # 更新UI
            $lblFolders.Text = "复制文件夹: $global:copiedFolders"
            $lblSize.Text = "总占用空间: $(Format-Size $global:totalSize)"
            $lblLastSize.Text = "最新项目: $folderName ($(Format-Size $folderSize))"
            
            return "文件夹已复制 (大小: $(Format-Size $folderSize))"
        }
        else {
            $fileName = Split-Path $sourcePath -Leaf
            $destFile = Join-Path -Path $targetPath -ChildPath $fileName
            
            # 如果目标文件夹不存在则创建
            if (-not (Test-Path -Path $targetPath)) {
                New-Item -ItemType Directory -Path $targetPath -Force | Out-Null
                Write-Log "创建目标文件夹: $targetPath" "Green"
            }
            
            # 获取文件大小
            $fileSize = (Get-Item $sourcePath).Length
            
            # 复制文件
            Write-Log "正在复制文件: $fileName [size:$fileSize]" -color Blue -isSize
            Copy-Item -Path $sourcePath -Destination $targetPath -Force
            
            # 更新统计
            $global:copiedFiles++
            $global:totalSize += $fileSize
            $global:lastItemSize = $fileSize
            $global:lastItemName = $fileName
            
            # 更新UI
            $lblFiles.Text = "复制文件: $global:copiedFiles"
            $lblSize.Text = "总占用空间: $(Format-Size $global:totalSize)"
            $lblLastSize.Text = "最新项目: $fileName ($(Format-Size $fileSize))"
            
            return "文件已复制 (大小: $(Format-Size $fileSize))"
        }
    }
    catch {
        return "错误: $($_.Exception.Message)"
    }
}

# 重置统计
function Reset-Stats {
    $global:copiedFiles = 0
    $global:copiedFolders = 0
    $global:totalSize = 0
    $global:lastItemSize = 0
    $global:lastItemName = ""
    
    $lblFiles.Text = "复制文件: 0"
    $lblFolders.Text = "复制文件夹: 0"
    $lblSize.Text = "总占用空间: 0 B"
    $lblLastSize.Text = "最新项目: N/A"
}

# 监控操作
function Monitor-Clipboard {
    # 获取剪贴板内容
    Add-Type -AssemblyName System.Windows.Forms
    $clipboardData = [System.Windows.Forms.Clipboard]::GetFileDropList()
    
    # 如果剪贴板中有文件
    if ($clipboardData.Count -gt 0) {
        # 将文件路径排序并连接成字符串用于比较
        $sortedPaths = $clipboardData | Sort-Object
        $currentClip = $sortedPaths -join '|'
        
        # 如果与上一次相同则跳过
        if ($currentClip -eq $global:clipboardHistory) {
            return
        }
        
        $global:clipboardHistory = $currentClip
        
        # 获取目标文件夹路径
        $targetPath = $txtTarget.Text
        
        # 在日志中添加分隔线
        Write-Log "------------------------------------"
        Write-Log "检测到新的剪贴板内容 ($($clipboardData.Count)个项目)" -color DarkBlue -isBold
        
        # 处理每个剪贴板项
        foreach ($item in $clipboardData) {
            $result = Copy-Resource -sourcePath $item -targetPath $targetPath
            
            if ($result -like "错误:*") {
                Write-Log $result "Red"
            }
            else {
                Write-Log $result "Green"
            }
        }
        
        Write-Log "复制操作完成" -color DarkGreen -isBold
        Write-Log "------------------------------------"
        
        # 更新内存显示
        Update-MemoryUsage
    }
}

# 开始监控
$btnStart.Add_Click({
    # 验证目标文件夹
    $targetPath = $txtTarget.Text
    if ([string]::IsNullOrWhiteSpace($targetPath)) {
        [System.Windows.Forms.MessageBox]::Show("请输入目标文件夹路径", "错误", "OK", "Error")
        return
    }
    
    # 启用/禁用按钮
    $btnStart.Enabled = $false
    $btnStop.Enabled = $true
    $btnReset.Enabled = $false
    
    # 更新状态
    $statusLabel.Text = "监控中 (间隔: $($numInterval.Value)秒)"
    
    # 显示进度条
    $progressBar.Visible = $true
    
    # 初始化全局变量
    $global:isMonitoring = $true
    $global:clipboardHistory = ""
    
    Write-Log "================================" -color DarkBlue
    Write-Log "监控已启动" -color DarkGreen -isBold
    Write-Log "目标目录: $targetPath"
    Write-Log "检测间隔: $($numInterval.Value)秒"
    Write-Log "按停止监控按钮终止程序"
    Write-Log "------------------------------------"
    
    # 创建并启动定时器
    $global:timer = New-Object System.Windows.Forms.Timer
    $global:timer.Interval = $numInterval.Value * 1000
    $global:timer.Add_Tick({
        Monitor-Clipboard
        Update-MemoryUsage
    })
    $global:timer.Start()
    
    # 初始化内存显示
    Update-MemoryUsage
})

# 停止监控
$btnStop.Add_Click({
    # 停止定时器
    if ($global:timer -ne $null) {
        $global:timer.Stop()
        $global:timer.Dispose()
        $global:timer = $null
    }
    
    # 启用/禁用按钮
    $btnStart.Enabled = $true
    $btnStop.Enabled = $false
    $btnReset.Enabled = $true
    
    # 更新状态
    $statusLabel.Text = "监控已停止"
    
    # 隐藏进度条
    $progressBar.Visible = $false
    
    # 更新全局变量
    $global:isMonitoring = $false
    
    Write-Log "------------------------------------"
    Write-Log "监控已停止" -color DarkRed -isBold
    Write-Log "================================" -color DarkBlue
    
    # 更新内存显示
    Update-MemoryUsage
})

# 重置统计
$btnReset.Add_Click({
    Reset-Stats
    Write-Log "统计信息已重置" "Blue"
})

# 应用程序关闭前事件
$form.Add_FormClosing({
    # 停止定时器
    if ($global:timer -ne $null) {
        $global:timer.Stop()
        $global:timer.Dispose()
    }
})

# 显示主窗体
$form.Add_Shown({ 
    $form.Activate() 
    Reset-Stats
    Write-Log "文件备份监控器已启动" -color DarkGreen
    Update-MemoryUsage
})
[void]$form.ShowDialog()