# 请求管理员权限
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process PowerShell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# 加载所需程序集
try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName PresentationFramework
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
    [System.Windows.Forms.MessageBox]::Show($errorMsg, "严重错误", "OK", "Error")
    exit
}

# 创建主窗体
$form = New-Object System.Windows.Forms.Form
$form.Text = "文件备份监控器V1.3"
$form.Size = New-Object System.Drawing.Size(750, 720)
$form.StartPosition = "CenterScreen"
$form.MinimizeBox = $true
$form.MaximizeBox = $false
$form.FormBorderStyle = "FixedDialog"
$form.BackColor = "#F0F0F0"

# 设置图标
try {
    $iconPath = "$env:WINDIR\System32\SHELL32.dll"
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
$txtTarget.Size = New-Object System.Drawing.Size(450, 20)
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

# 控制按钮面板
$controlPanel = New-Object System.Windows.Forms.Panel
$controlPanel.Location = New-Object System.Drawing.Point(20, 90)
$controlPanel.Size = New-Object System.Drawing.Size(710, 35)
$controlPanel.BackColor = "#E0F7FA"
$controlPanel.BorderStyle = "FixedSingle"
$form.Controls.Add($controlPanel)

# 开始按钮
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Location = New-Object System.Drawing.Point(5, 5)
$btnStart.Size = New-Object System.Drawing.Size(80, 25)
$btnStart.Text = "开始监控"
$btnStart.BackColor = "#4CAF50"
$btnStart.ForeColor = "White"
$controlPanel.Controls.Add($btnStart)

# 停止按钮
$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Location = New-Object System.Drawing.Point(90, 5)
$btnStop.Size = New-Object System.Drawing.Size(80, 25)
$btnStop.Text = "停止监控"
$btnStop.BackColor = "#F44336"
$btnStop.ForeColor = "White"
$btnStop.Enabled = $false
$controlPanel.Controls.Add($btnStop)

# 暂停按钮
$btnPause = New-Object System.Windows.Forms.Button
$btnPause.Location = New-Object System.Drawing.Point(175, 5)
$btnPause.Size = New-Object System.Drawing.Size(80, 25)
$btnPause.Text = "暂停监控"
$btnPause.BackColor = "#FF9800"
$btnPause.ForeColor = "White"
$btnPause.Enabled = $false
$controlPanel.Controls.Add($btnPause)

# 继续按钮
$btnResume = New-Object System.Windows.Forms.Button
$btnResume.Location = New-Object System.Drawing.Point(260, 5)
$btnResume.Size = New-Object System.Drawing.Size(80, 25)
$btnResume.Text = "继续监控"
$btnResume.BackColor = "#2196F3"
$btnResume.ForeColor = "White"
$btnResume.Enabled = $false
$controlPanel.Controls.Add($btnResume)

# 重置统计按钮
$btnReset = New-Object System.Windows.Forms.Button
$btnReset.Location = New-Object System.Drawing.Point(345, 5)
$btnReset.Size = New-Object System.Drawing.Size(80, 25)
$btnReset.Text = "重置统计"
$btnReset.BackColor = "#FFC107"
$btnReset.ForeColor = "Black"
$controlPanel.Controls.Add($btnReset)

# 撤销按钮
$btnUndo = New-Object System.Windows.Forms.Button
$btnUndo.Location = New-Object System.Drawing.Point(430, 5)
$btnUndo.Size = New-Object System.Drawing.Size(100, 25)
$btnUndo.Text = "撤销最新操作"
$btnUndo.BackColor = "#9C27B0"
$btnUndo.ForeColor = "White"
$btnUndo.Enabled = $false
$controlPanel.Controls.Add($btnUndo)

# 状态指示器
$statusIndicator = New-Object System.Windows.Forms.Label
$statusIndicator.Location = New-Object System.Drawing.Point(535, 7)
$statusIndicator.Size = New-Object System.Drawing.Size(170, 20)
$statusIndicator.Text = "状态: 准备就绪"
$controlPanel.Controls.Add($statusIndicator)

# 统计面板
$statsPanel = New-Object System.Windows.Forms.Panel
$statsPanel.Location = New-Object System.Drawing.Point(20, 135)
$statsPanel.Size = New-Object System.Drawing.Size(710, 40)
$statsPanel.BackColor = "#E8F5E9"
$statsPanel.BorderStyle = "FixedSingle"
$form.Controls.Add($statsPanel)

# 统计标签
$lblFiles = New-Object System.Windows.Forms.Label
$lblFiles.Location = New-Object System.Drawing.Point(15, 12)
$lblFiles.Size = New-Object System.Drawing.Size(120, 20)
$lblFiles.Text = "复制文件: 0"
$statsPanel.Controls.Add($lblFiles)

$lblFolders = New-Object System.Windows.Forms.Label
$lblFolders.Location = New-Object System.Drawing.Point(150, 12)
$lblFolders.Size = New-Object System.Drawing.Size(140, 20)
$lblFolders.Text = "复制文件夹: 0"
$statsPanel.Controls.Add($lblFolders)

$lblSize = New-Object System.Windows.Forms.Label
$lblSize.Location = New-Object System.Drawing.Point(310, 12)
$lblSize.Size = New-Object System.Drawing.Size(200, 20)
$lblSize.Text = "总占用空间: 0 B"
$statsPanel.Controls.Add($lblSize)

$lblLastSize = New-Object System.Windows.Forms.Label
$lblLastSize.Location = New-Object System.Drawing.Point(520, 12)
$lblLastSize.Size = New-Object System.Drawing.Size(170, 20)
$lblLastSize.Text = "最新项目: N/A"
$statsPanel.Controls.Add($lblLastSize)

# 操作历史面板
$historyPanel = New-Object System.Windows.Forms.Panel
$historyPanel.Location = New-Object System.Drawing.Point(20, 185)
$historyPanel.Size = New-Object System.Drawing.Size(710, 50)
$historyPanel.BackColor = "#FFF8E1"
$historyPanel.BorderStyle = "FixedSingle"
$form.Controls.Add($historyPanel)

$lblHistory = New-Object System.Windows.Forms.Label
$lblHistory.Location = New-Object System.Drawing.Point(15, 5)
$lblHistory.Size = New-Object System.Drawing.Size(680, 20)
$lblHistory.Text = "操作历史 (最新 5 项):"
$historyPanel.Controls.Add($lblHistory)

$lblOperations = New-Object System.Windows.Forms.Label
$lblOperations.Location = New-Object System.Drawing.Point(15, 25)
$lblOperations.Size = New-Object System.Drawing.Size(680, 20)
$lblOperations.Text = "无操作历史"
$historyPanel.Controls.Add($lblOperations)

# 日志框
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Location = New-Object System.Drawing.Point(20, 245)
$lblLog.Size = New-Object System.Drawing.Size(120, 20)
$lblLog.Text = "详细日志:"
$form.Controls.Add($lblLog)

$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 265)
$txtLog.Size = New-Object System.Drawing.Size(710, 400)
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
$statusBar.Items.Add($statusLabel)

# 实时内存标签
$memoryLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$memoryLabel.Text = "内存: 计算中..."
$statusBar.Items.Add($memoryLabel)

# 进度条
$progressBar = New-Object System.Windows.Forms.ToolStripProgressBar
$progressBar.Style = "Marquee"
$progressBar.Visible = $false
$statusBar.Items.Add($progressBar)

# 初始化全局变量
$global:isMonitoring = $false
$global:isPaused = $false
$global:clipboardHistory = ""
$global:timer = $null
$global:copiedFiles = 0
$global:copiedFolders = 0
$global:totalSize = 0
$global:lastItemSize = 0
$global:lastItemName = ""
$global:lastItemPath = ""
$global:lastItemType = ""
$global:operationsHistory = New-Object System.Collections.ArrayList
$global:maxHistoryCount = 5

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
    param([string]$message, [string]$color = "Black")
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    $txtLog.SelectionColor = $color
    $txtLog.AppendText("[$timestamp] $message`n")
    $txtLog.ScrollToCaret()
}

# 添加操作历史
function Add-OperationHistory {
    param([string]$action, [string]$itemPath, [long]$size)
    
    $historyItem = @{
        Action = $action
        ItemPath = $itemPath
        Size = $size
        Timestamp = Get-Date -Format "HH:mm:ss"
    }
    
    $global:operationsHistory.Insert(0, $historyItem)
    if ($global:operationsHistory.Count -gt $global:maxHistoryCount) {
        $global:operationsHistory.RemoveAt($global:maxHistoryCount)
    }
    
    $historyText = ""
    foreach ($item in $global:operationsHistory) {
        $historyText += "$($item.Timestamp) $($item.Action): $($item.ItemPath)`n"
    }
    
    $lblOperations.Text = $historyText.Trim()
}

# 复制文件/文件夹
function Copy-Resource {
    param([string]$sourcePath, [string]$targetPath)
    
    try {
        if (Test-Path -Path $sourcePath -PathType Container) {
            $folderName = Split-Path $sourcePath -Leaf
            $destFolder = Join-Path -Path $targetPath -ChildPath $folderName
            
            if (-not (Test-Path -Path $targetPath)) {
                New-Item -ItemType Directory -Path $targetPath -Force | Out-Null
                Write-Log "创建目标文件夹: $targetPath" "Green"
            }
            
            Write-Log "正在复制文件夹: $folderName" "Blue"
            
            $folderSize = (Get-ChildItem -Path $sourcePath -Recurse | Measure-Object -Property Length -Sum).Sum
            
            $null = New-Item -ItemType Directory -Path $destFolder -Force
            Copy-Item -Path "$sourcePath\*" -Destination $destFolder -Recurse -Force
            
            $global:copiedFolders++
            $global:totalSize += $folderSize
            $global:lastItemSize = $folderSize
            $global:lastItemName = $folderName
            $global:lastItemPath = $destFolder
            $global:lastItemType = "folder"
            
            $lblFolders.Text = "复制文件夹: $global:copiedFolders"
            $lblSize.Text = "总占用空间: $(Format-Size $global:totalSize)"
            $lblLastSize.Text = "最新项目: $folderName ($(Format-Size $folderSize))"
            
            Add-OperationHistory -action "复制" -itemPath $destFolder -size $folderSize
            return "文件夹已复制 (大小: $(Format-Size $folderSize))"
        }
        else {
            $fileName = Split-Path $sourcePath -Leaf
            $destFile = Join-Path -Path $targetPath -ChildPath $fileName
            
            if (-not (Test-Path -Path $targetPath)) {
                New-Item -ItemType Directory -Path $targetPath -Force | Out-Null
                Write-Log "创建目标文件夹: $targetPath" "Green"
            }
            
            $fileSize = (Get-Item $sourcePath).Length
            Write-Log "正在复制文件: $fileName" "Blue"
            Copy-Item -Path $sourcePath -Destination $targetPath -Force
            
            $global:copiedFiles++
            $global:totalSize += $fileSize
            $global:lastItemSize = $fileSize
            $global:lastItemName = $fileName
            $global:lastItemPath = $destFile
            $global:lastItemType = "file"
            
            $lblFiles.Text = "复制文件: $global:copiedFiles"
            $lblSize.Text = "总占用空间: $(Format-Size $global:totalSize)"
            $lblLastSize.Text = "最新项目: $fileName ($(Format-Size $fileSize))"
            
            Add-OperationHistory -action "复制" -itemPath $destFile -size $fileSize
            return "文件已复制 (大小: $(Format-Size $fileSize))"
        }
    }
    catch {
        return "错误: $($_.Exception.Message)"
    }
}

# 撤销最新操作
function Undo-LastOperation {
    if (-not $global:lastItemPath) {
        Write-Log "无操作可撤销" "DarkOrange"
        return $false
    }
    
    try {
        if ($global:lastItemType -eq "file") {
            if (Test-Path -Path $global:lastItemPath -PathType Leaf) {
                Write-Log "正在撤销文件复制: $global:lastItemName" "Purple"
                Remove-Item -Path $global:lastItemPath -Force
                
                $global:copiedFiles--
                $global:totalSize -= $global:lastItemSize
                
                $lblFiles.Text = "复制文件: $global:copiedFiles"
                $lblSize.Text = "总占用空间: $(Format-Size $global:totalSize)"
                
                Add-OperationHistory -action "撤销" -itemPath $global:lastItemPath -size $global:lastItemSize
                Write-Log "成功撤销文件: $global:lastItemName" "DarkMagenta"
                return $true
            }
        }
        elseif ($global:lastItemType -eq "folder") {
            if (Test-Path -Path $global:lastItemPath -PathType Container) {
                Write-Log "正在撤销文件夹复制: $global:lastItemName" "Purple"
                Remove-Item -Path $global:lastItemPath -Recurse -Force
                
                $global:copiedFolders--
                $global:totalSize -= $global:lastItemSize
                
                $lblFolders.Text = "复制文件夹: $global:copiedFolders"
                $lblSize.Text = "总占用空间: $(Format-Size $global:totalSize)"
                
                Add-OperationHistory -action "撤销" -itemPath $global:lastItemPath -size $global:lastItemSize
                Write-Log "成功撤销文件夹: $global:lastItemName" "DarkMagenta"
                return $true
            }
        }
    }
    catch {
        Write-Log "撤销失败: $($_.Exception.Message)" "Red"
        return $false
    }
    
    Write-Log "找不到要撤销的项目: $global:lastItemPath" "DarkOrange"
    return $false
}

# 重置统计
function Reset-Stats {
    $global:copiedFiles = 0
    $global:copiedFolders = 0
    $global:totalSize = 0
    $global:lastItemSize = 0
    $global:lastItemName = ""
    $global:lastItemPath = ""
    $global:lastItemType = ""
    $global:operationsHistory.Clear()
    $lblOperations.Text = "无操作历史"
    
    $lblFiles.Text = "复制文件: 0"
    $lblFolders.Text = "复制文件夹: 0"
    $lblSize.Text = "总占用空间: 0 B"
    $lblLastSize.Text = "最新项目: N/A"
}

# 监控操作
function Monitor-Clipboard {
    if ($global:isPaused) {
        return
    }
    
    Add-Type -AssemblyName System.Windows.Forms
    $clipboardData = [System.Windows.Forms.Clipboard]::GetFileDropList()
    
    if ($clipboardData.Count -gt 0) {
        $sortedPaths = $clipboardData | Sort-Object
        $currentClip = $sortedPaths -join '|'
        
        if ($currentClip -eq $global:clipboardHistory) {
            return
        }
        
        $global:clipboardHistory = $currentClip
        $targetPath = $txtTarget.Text
        
        Write-Log "检测到新的剪贴板内容 ($($clipboardData.Count)个项目)" "DarkBlue"
        
        foreach ($item in $clipboardData) {
            $result = Copy-Resource -sourcePath $item -targetPath $targetPath
            
            if ($result -like "错误:*") {
                Write-Log $result "Red"
            }
            else {
                Write-Log $result "Green"
            }
        }
        
        Write-Log "复制操作完成" "DarkGreen"
        Update-MemoryUsage
        $btnUndo.Enabled = $true
    }
}

# 开始监控按钮
$btnStart.Add_Click({
    $targetPath = $txtTarget.Text
    if ([string]::IsNullOrWhiteSpace($targetPath)) {
        [System.Windows.Forms.MessageBox]::Show("请输入目标文件夹路径", "错误", "OK", "Error")
        return
    }
    
    $btnStart.Enabled = $false
    $btnStop.Enabled = $true
    $btnPause.Enabled = $true
    $statusIndicator.Text = "状态: 运行中"
    $statusLabel.Text = "监控中 (间隔: $($numInterval.Value)秒)"
    $progressBar.Visible = $true
    
    $global:isMonitoring = $true
    $global:isPaused = $false
    $global:clipboardHistory = ""
    
    Write-Log "监控已启动" "DarkGreen"
    Write-Log "目标目录: $targetPath"
    Write-Log "检测间隔: $($numInterval.Value)秒"
    
    $global:timer = New-Object System.Windows.Forms.Timer
    $global:timer.Interval = $numInterval.Value * 1000
    $global:timer.Add_Tick({
        Monitor-Clipboard
        Update-MemoryUsage
    })
    $global:timer.Start()
    Update-MemoryUsage
})

# 停止监控按钮
$btnStop.Add_Click({
    if ($global:timer -ne $null) {
        $global:timer.Stop()
        $global:timer.Dispose()
        $global:timer = $null
    }
    
    $btnStart.Enabled = $true
    $btnStop.Enabled = $false
    $btnPause.Enabled = $false
    $btnResume.Enabled = $false
    $btnReset.Enabled = $true
    
    $statusIndicator.Text = "状态: 已停止"
    $statusLabel.Text = "监控已停止"
    $progressBar.Visible = $false
    
    $global:isMonitoring = $false
    $global:isPaused = $false
    
    Write-Log "监控已停止" "DarkRed"
})

# 暂停监控按钮
$btnPause.Add_Click({
    if ($global:isMonitoring -and -not $global:isPaused) {
        $global:isPaused = $true
        $btnPause.Enabled = $false
        $btnResume.Enabled = $true
        $statusIndicator.Text = "状态: 已暂停"
        Write-Log "监控已暂停" "DarkOrange"
    }
})

# 继续监控按钮
$btnResume.Add_Click({
    if ($global:isMonitoring -and $global:isPaused) {
        $global:isPaused = $false
        $btnPause.Enabled = $true
        $btnResume.Enabled = $false
        $statusIndicator.Text = "状态: 运行中"
        Write-Log "监控继续运行" "Green"
    }
})

# 撤销按钮
$btnUndo.Add_Click({
    $success = Undo-LastOperation
    if ($success -and -not $global:lastItemPath) {
        $btnUndo.Enabled = $false
    }
})

# 重置统计按钮
$btnReset.Add_Click({
    Reset-Stats
    $btnUndo.Enabled = $false
    Write-Log "统计信息已重置" "Blue"
})

# 窗体关闭事件
$form.Add_FormClosing({
    if ($global:timer -ne $null) {
        $global:timer.Stop()
        $global:timer.Dispose()
    }
})

# 初始化
$form.Add_Shown({ 
    $form.Activate() 
    Reset-Stats
    Write-Log "文件备份监控器已启动V1.3" "DarkGreen"
    Update-MemoryUsage
})
[void]$form.ShowDialog()