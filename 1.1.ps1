Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# 创建主窗体
$form = New-Object System.Windows.Forms.Form
$form.Text = "文件备份监控器V1.1"
$form.Size = New-Object System.Drawing.Size(650, 550)
$form.StartPosition = "CenterScreen"
$form.MinimizeBox = $false
$form.MaximizeBox = $false
$form.FormBorderStyle = "FixedDialog"
$form.BackColor = "#F0F0F0"

# 图标
$icon = [System.Drawing.Icon]::ExtractAssociatedIcon("$env:WINDIR\system32\imageres.dll")
$form.Icon = $icon

# 标签和输入框 - 目标文件夹
$lblTarget = New-Object System.Windows.Forms.Label
$lblTarget.Location = New-Object System.Drawing.Point(20, 20)
$lblTarget.Size = New-Object System.Drawing.Size(120, 20)
$lblTarget.Text = "目标文件夹:"
$form.Controls.Add($lblTarget)

$txtTarget = New-Object System.Windows.Forms.TextBox
$txtTarget.Location = New-Object System.Drawing.Point(140, 20)
$txtTarget.Size = New-Object System.Drawing.Size(350, 20)
$txtTarget.Text = ""
$form.Controls.Add($txtTarget)

# 浏览按钮
$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Location = New-Object System.Drawing.Point(500, 19)
$btnBrowse.Size = New-Object System.Drawing.Size(100, 23)
$btnBrowse.Text = "浏览..."
$btnBrowse.BackColor = "White"
$btnBrowse.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "选择目标文件夹"
    $folderBrowser.SelectedPath = $txtTarget.Text
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtTarget.Text = $folderBrowser.SelectedPath
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

# 统计面板
$statsPanel = New-Object System.Windows.Forms.Panel
$statsPanel.Location = New-Object System.Drawing.Point(20, 95)
$statsPanel.Size = New-Object System.Drawing.Size(580, 40)
$statsPanel.BackColor = "#E8F5E9"
$statsPanel.BorderStyle = "FixedSingle"
$form.Controls.Add($statsPanel)

# 统计标签
$lblFiles = New-Object System.Windows.Forms.Label
$lblFiles.Location = New-Object System.Drawing.Point(15, 12)
$lblFiles.Size = New-Object System.Drawing.Size(160, 20)
$lblFiles.Text = "复制文件: 0"
$lblFiles.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$statsPanel.Controls.Add($lblFiles)

$lblFolders = New-Object System.Windows.Forms.Label
$lblFolders.Location = New-Object System.Drawing.Point(185, 12)
$lblFolders.Size = New-Object System.Drawing.Size(160, 20)
$lblFolders.Text = "复制文件夹: 0"
$lblFolders.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$statsPanel.Controls.Add($lblFolders)

$lblSize = New-Object System.Windows.Forms.Label
$lblSize.Location = New-Object System.Drawing.Point(355, 12)
$lblSize.Size = New-Object System.Drawing.Size(200, 20)
$lblSize.Text = "总占用空间: 0 B"
$lblSize.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$statsPanel.Controls.Add($lblSize)

# 日志框
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Location = New-Object System.Drawing.Point(20, 145)
$lblLog.Size = New-Object System.Drawing.Size(120, 20)
$lblLog.Text = "操作日志:"
$form.Controls.Add($lblLog)

$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 165)
$txtLog.Size = New-Object System.Drawing.Size(580, 280)
$txtLog.ReadOnly = $true
$txtLog.BackColor = "White"
$txtLog.ScrollBars = "Vertical"
$form.Controls.Add($txtLog)

# 状态栏
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusBar.Dock = "Bottom"
$form.Controls.Add($statusBar)

# 状态标签
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "准备就绪"
$statusBar.Items.Add($statusLabel)

# 内存状态标签
$memoryLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$memoryLabel.Spring = $true
$memoryLabel.TextAlign = "MiddleRight"
$memoryLabel.Text = "内存: 0 MB"
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

# 格式化文件大小函数
function Format-Size {
    param([long]$bytes)
    if ($bytes -gt 1GB) {
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
    $process = Get-Process -Id $pid
    $memoryUsage = [math]::Round($process.WorkingSet64 / 1MB, 2)
    $memoryLabel.Text = "内存: $memoryUsage MB"
}

# 写入日志
function Write-Log {
    param([string]$message, [string]$color = "Black")
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    $logEntry = "[$timestamp] $message"
    
    $txtLog.SelectionStart = $txtLog.TextLength
    $txtLog.SelectionLength = 0
    $txtLog.SelectionColor = $color
    $txtLog.AppendText($logEntry + "`n")
    $txtLog.SelectionColor = "Black"
    $txtLog.ScrollToCaret()
    
    # 更新内存显示
    Update-MemoryUsage
}

# 复制文件/文件夹并更新统计
function Copy-Resource {
    param([string]$sourcePath, [string]$targetPath)
    
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
            
            # 复制文件夹
            Write-Log "正在复制文件夹: $folderName" "Blue"
            
            # 复制前获取文件夹大小（大约）
            $folderSize = (Get-ChildItem -Path $sourcePath -Recurse | Measure-Object -Property Length -Sum).Sum
            
            # 复制操作
            $null = New-Item -ItemType Directory -Path $destFolder -Force
            Copy-Item -Path "$sourcePath\*" -Destination $destFolder -Recurse -Force
            
            # 更新统计
            $global:copiedFolders++
            $global:totalSize += $folderSize
            
            # 更新UI
            $lblFolders.Text = "复制文件夹: $global:copiedFolders"
            $lblSize.Text = "总占用空间: $(Format-Size $global:totalSize)"
            
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
            Write-Log "正在复制文件: $fileName (大小: $(Format-Size $fileSize))" "Blue"
            Copy-Item -Path $sourcePath -Destination $targetPath -Force
            
            # 更新统计
            $global:copiedFiles++
            $global:totalSize += $fileSize
            
            # 更新UI
            $lblFiles.Text = "复制文件: $global:copiedFiles"
            $lblSize.Text = "总占用空间: $(Format-Size $global:totalSize)"
            
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
    
    $lblFiles.Text = "复制文件: 0"
    $lblFolders.Text = "复制文件夹: 0"
    $lblSize.Text = "总占用空间: 0 B"
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
        
        # 处理每个剪贴板项
        foreach ($item in $clipboardData) {
            $result = Copy-Resource -sourcePath $item -targetPath $targetPath
            
            if ($result -like "错误:*") {
                Write-Log $result "Red"
            }
            else {
                Write-Log "完成: $result" "Green"
            }
            
            # 更新内存显示
            Update-MemoryUsage
        }
        
        Write-Log "------------------------------------"
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
    
    # 更新状态
    $statusLabel.Text = "监控中 (间隔: $($numInterval.Value)秒)"
    
    # 显示进度条
    $progressBar.Visible = $true
    
    # 初始化全局变量
    $global:isMonitoring = $true
    $global:clipboardHistory = ""
    
    Write-Log "监控已启动..."
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
    
    # 更新状态
    $statusLabel.Text = "监控已停止"
    
    # 隐藏进度条
    $progressBar.Visible = $false
    
    # 更新全局变量
    $global:isMonitoring = $false
    
    Write-Log "监控已停止"
    Write-Log "------------------------------------"
    
    # 更新内存显示
    Update-MemoryUsage
})

# 添加重置统计按钮
$btnReset = New-Object System.Windows.Forms.Button
$btnReset.Location = New-Object System.Drawing.Point(480, 59)
$btnReset.Size = New-Object System.Drawing.Size(100, 23)
$btnReset.Text = "重置统计"
$btnReset.BackColor = "#FFC107"
$btnReset.Add_Click({
    Reset-Stats
    Write-Log "统计信息已重置" "Blue"
})
$form.Controls.Add($btnReset)

# 应用程序关闭前事件
$form.Add_FormClosing({
    # 停止定时器
    if ($global:timer -ne $null) {
        $global:timer.Stop()
        $global:timer.Dispose()
    }
})

# 显示主窗体
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()