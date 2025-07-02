Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# ����������
$form = New-Object System.Windows.Forms.Form
$form.Text = "�ļ����ݼ����V1.1"
$form.Size = New-Object System.Drawing.Size(650, 550)
$form.StartPosition = "CenterScreen"
$form.MinimizeBox = $false
$form.MaximizeBox = $false
$form.FormBorderStyle = "FixedDialog"
$form.BackColor = "#F0F0F0"

# ͼ��
$icon = [System.Drawing.Icon]::ExtractAssociatedIcon("$env:WINDIR\system32\imageres.dll")
$form.Icon = $icon

# ��ǩ������� - Ŀ���ļ���
$lblTarget = New-Object System.Windows.Forms.Label
$lblTarget.Location = New-Object System.Drawing.Point(20, 20)
$lblTarget.Size = New-Object System.Drawing.Size(120, 20)
$lblTarget.Text = "Ŀ���ļ���:"
$form.Controls.Add($lblTarget)

$txtTarget = New-Object System.Windows.Forms.TextBox
$txtTarget.Location = New-Object System.Drawing.Point(140, 20)
$txtTarget.Size = New-Object System.Drawing.Size(350, 20)
$txtTarget.Text = ""
$form.Controls.Add($txtTarget)

# �����ť
$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Location = New-Object System.Drawing.Point(500, 19)
$btnBrowse.Size = New-Object System.Drawing.Size(100, 23)
$btnBrowse.Text = "���..."
$btnBrowse.BackColor = "White"
$btnBrowse.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "ѡ��Ŀ���ļ���"
    $folderBrowser.SelectedPath = $txtTarget.Text
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtTarget.Text = $folderBrowser.SelectedPath
    }
})
$form.Controls.Add($btnBrowse)

# ��ǩ������� - �����
$lblInterval = New-Object System.Windows.Forms.Label
$lblInterval.Location = New-Object System.Drawing.Point(20, 60)
$lblInterval.Size = New-Object System.Drawing.Size(120, 20)
$lblInterval.Text = "�����(��):"
$form.Controls.Add($lblInterval)

$numInterval = New-Object System.Windows.Forms.NumericUpDown
$numInterval.Location = New-Object System.Drawing.Point(140, 60)
$numInterval.Size = New-Object System.Drawing.Size(100, 20)
$numInterval.Value = 1
$numInterval.Minimum = 1
$numInterval.Maximum = 10
$numInterval.Increment = 0.5
$form.Controls.Add($numInterval)

# ��ʼ��ť
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Location = New-Object System.Drawing.Point(260, 60)
$btnStart.Size = New-Object System.Drawing.Size(100, 23)
$btnStart.Text = "��ʼ���"
$btnStart.BackColor = "#4CAF50"
$btnStart.ForeColor = "White"
$btnStart.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8.25, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnStart)

# ֹͣ��ť
$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Location = New-Object System.Drawing.Point(370, 60)
$btnStop.Size = New-Object System.Drawing.Size(100, 23)
$btnStop.Text = "ֹͣ���"
$btnStop.BackColor = "#F44336"
$btnStop.ForeColor = "White"
$btnStop.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8.25, [System.Drawing.FontStyle]::Bold)
$btnStop.Enabled = $false
$form.Controls.Add($btnStop)

# ͳ�����
$statsPanel = New-Object System.Windows.Forms.Panel
$statsPanel.Location = New-Object System.Drawing.Point(20, 95)
$statsPanel.Size = New-Object System.Drawing.Size(580, 40)
$statsPanel.BackColor = "#E8F5E9"
$statsPanel.BorderStyle = "FixedSingle"
$form.Controls.Add($statsPanel)

# ͳ�Ʊ�ǩ
$lblFiles = New-Object System.Windows.Forms.Label
$lblFiles.Location = New-Object System.Drawing.Point(15, 12)
$lblFiles.Size = New-Object System.Drawing.Size(160, 20)
$lblFiles.Text = "�����ļ�: 0"
$lblFiles.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$statsPanel.Controls.Add($lblFiles)

$lblFolders = New-Object System.Windows.Forms.Label
$lblFolders.Location = New-Object System.Drawing.Point(185, 12)
$lblFolders.Size = New-Object System.Drawing.Size(160, 20)
$lblFolders.Text = "�����ļ���: 0"
$lblFolders.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$statsPanel.Controls.Add($lblFolders)

$lblSize = New-Object System.Windows.Forms.Label
$lblSize.Location = New-Object System.Drawing.Point(355, 12)
$lblSize.Size = New-Object System.Drawing.Size(200, 20)
$lblSize.Text = "��ռ�ÿռ�: 0 B"
$lblSize.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$statsPanel.Controls.Add($lblSize)

# ��־��
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Location = New-Object System.Drawing.Point(20, 145)
$lblLog.Size = New-Object System.Drawing.Size(120, 20)
$lblLog.Text = "������־:"
$form.Controls.Add($lblLog)

$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 165)
$txtLog.Size = New-Object System.Drawing.Size(580, 280)
$txtLog.ReadOnly = $true
$txtLog.BackColor = "White"
$txtLog.ScrollBars = "Vertical"
$form.Controls.Add($txtLog)

# ״̬��
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusBar.Dock = "Bottom"
$form.Controls.Add($statusBar)

# ״̬��ǩ
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "׼������"
$statusBar.Items.Add($statusLabel)

# �ڴ�״̬��ǩ
$memoryLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$memoryLabel.Spring = $true
$memoryLabel.TextAlign = "MiddleRight"
$memoryLabel.Text = "�ڴ�: 0 MB"
$statusBar.Items.Add($memoryLabel)

# ������
$progressBar = New-Object System.Windows.Forms.ToolStripProgressBar
$progressBar.Style = "Marquee"
$progressBar.AutoSize = $false
$progressBar.Width = 150
$progressBar.Visible = $false
$statusBar.Items.Add($progressBar)

# ��ʼ��ȫ�ֱ���
$global:isMonitoring = $false
$global:clipboardHistory = ""
$global:timer = $null

# ͳ�Ʊ���
$global:copiedFiles = 0
$global:copiedFolders = 0
$global:totalSize = 0

# ��ʽ���ļ���С����
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

# �����ڴ���ʾ����
function Update-MemoryUsage {
    $process = Get-Process -Id $pid
    $memoryUsage = [math]::Round($process.WorkingSet64 / 1MB, 2)
    $memoryLabel.Text = "�ڴ�: $memoryUsage MB"
}

# д����־
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
    
    # �����ڴ���ʾ
    Update-MemoryUsage
}

# �����ļ�/�ļ��в�����ͳ��
function Copy-Resource {
    param([string]$sourcePath, [string]$targetPath)
    
    try {
        # ���Դ���ļ������ļ���
        if (Test-Path -Path $sourcePath -PathType Container) {
            $folderName = Split-Path $sourcePath -Leaf
            $destFolder = Join-Path -Path $targetPath -ChildPath $folderName
            
            # ���Ŀ���ļ��в������򴴽�
            if (-not (Test-Path -Path $targetPath)) {
                New-Item -ItemType Directory -Path $targetPath -Force | Out-Null
                Write-Log "����Ŀ���ļ���: $targetPath" "Green"
            }
            
            # �����ļ���
            Write-Log "���ڸ����ļ���: $folderName" "Blue"
            
            # ����ǰ��ȡ�ļ��д�С����Լ��
            $folderSize = (Get-ChildItem -Path $sourcePath -Recurse | Measure-Object -Property Length -Sum).Sum
            
            # ���Ʋ���
            $null = New-Item -ItemType Directory -Path $destFolder -Force
            Copy-Item -Path "$sourcePath\*" -Destination $destFolder -Recurse -Force
            
            # ����ͳ��
            $global:copiedFolders++
            $global:totalSize += $folderSize
            
            # ����UI
            $lblFolders.Text = "�����ļ���: $global:copiedFolders"
            $lblSize.Text = "��ռ�ÿռ�: $(Format-Size $global:totalSize)"
            
            return "�ļ����Ѹ��� (��С: $(Format-Size $folderSize))"
        }
        else {
            $fileName = Split-Path $sourcePath -Leaf
            $destFile = Join-Path -Path $targetPath -ChildPath $fileName
            
            # ���Ŀ���ļ��в������򴴽�
            if (-not (Test-Path -Path $targetPath)) {
                New-Item -ItemType Directory -Path $targetPath -Force | Out-Null
                Write-Log "����Ŀ���ļ���: $targetPath" "Green"
            }
            
            # ��ȡ�ļ���С
            $fileSize = (Get-Item $sourcePath).Length
            
            # �����ļ�
            Write-Log "���ڸ����ļ�: $fileName (��С: $(Format-Size $fileSize))" "Blue"
            Copy-Item -Path $sourcePath -Destination $targetPath -Force
            
            # ����ͳ��
            $global:copiedFiles++
            $global:totalSize += $fileSize
            
            # ����UI
            $lblFiles.Text = "�����ļ�: $global:copiedFiles"
            $lblSize.Text = "��ռ�ÿռ�: $(Format-Size $global:totalSize)"
            
            return "�ļ��Ѹ��� (��С: $(Format-Size $fileSize))"
        }
    }
    catch {
        return "����: $($_.Exception.Message)"
    }
}

# ����ͳ��
function Reset-Stats {
    $global:copiedFiles = 0
    $global:copiedFolders = 0
    $global:totalSize = 0
    
    $lblFiles.Text = "�����ļ�: 0"
    $lblFolders.Text = "�����ļ���: 0"
    $lblSize.Text = "��ռ�ÿռ�: 0 B"
}

# ��ز���
function Monitor-Clipboard {
    # ��ȡ����������
    Add-Type -AssemblyName System.Windows.Forms
    $clipboardData = [System.Windows.Forms.Clipboard]::GetFileDropList()
    
    # ��������������ļ�
    if ($clipboardData.Count -gt 0) {
        # ���ļ�·���������ӳ��ַ������ڱȽ�
        $sortedPaths = $clipboardData | Sort-Object
        $currentClip = $sortedPaths -join '|'
        
        # �������һ����ͬ������
        if ($currentClip -eq $global:clipboardHistory) {
            return
        }
        
        $global:clipboardHistory = $currentClip
        
        # ��ȡĿ���ļ���·��
        $targetPath = $txtTarget.Text
        
        # ����ÿ����������
        foreach ($item in $clipboardData) {
            $result = Copy-Resource -sourcePath $item -targetPath $targetPath
            
            if ($result -like "����:*") {
                Write-Log $result "Red"
            }
            else {
                Write-Log "���: $result" "Green"
            }
            
            # �����ڴ���ʾ
            Update-MemoryUsage
        }
        
        Write-Log "------------------------------------"
    }
}

# ��ʼ���
$btnStart.Add_Click({
    # ��֤Ŀ���ļ���
    $targetPath = $txtTarget.Text
    if ([string]::IsNullOrWhiteSpace($targetPath)) {
        [System.Windows.Forms.MessageBox]::Show("������Ŀ���ļ���·��", "����", "OK", "Error")
        return
    }
    
    # ����/���ð�ť
    $btnStart.Enabled = $false
    $btnStop.Enabled = $true
    
    # ����״̬
    $statusLabel.Text = "����� (���: $($numInterval.Value)��)"
    
    # ��ʾ������
    $progressBar.Visible = $true
    
    # ��ʼ��ȫ�ֱ���
    $global:isMonitoring = $true
    $global:clipboardHistory = ""
    
    Write-Log "���������..."
    Write-Log "Ŀ��Ŀ¼: $targetPath"
    Write-Log "�����: $($numInterval.Value)��"
    Write-Log "��ֹͣ��ذ�ť��ֹ����"
    Write-Log "------------------------------------"
    
    # ������������ʱ��
    $global:timer = New-Object System.Windows.Forms.Timer
    $global:timer.Interval = $numInterval.Value * 1000
    $global:timer.Add_Tick({
        Monitor-Clipboard
        Update-MemoryUsage
    })
    $global:timer.Start()
    
    # ��ʼ���ڴ���ʾ
    Update-MemoryUsage
})

# ֹͣ���
$btnStop.Add_Click({
    # ֹͣ��ʱ��
    if ($global:timer -ne $null) {
        $global:timer.Stop()
        $global:timer.Dispose()
        $global:timer = $null
    }
    
    # ����/���ð�ť
    $btnStart.Enabled = $true
    $btnStop.Enabled = $false
    
    # ����״̬
    $statusLabel.Text = "�����ֹͣ"
    
    # ���ؽ�����
    $progressBar.Visible = $false
    
    # ����ȫ�ֱ���
    $global:isMonitoring = $false
    
    Write-Log "�����ֹͣ"
    Write-Log "------------------------------------"
    
    # �����ڴ���ʾ
    Update-MemoryUsage
})

# �������ͳ�ư�ť
$btnReset = New-Object System.Windows.Forms.Button
$btnReset.Location = New-Object System.Drawing.Point(480, 59)
$btnReset.Size = New-Object System.Drawing.Size(100, 23)
$btnReset.Text = "����ͳ��"
$btnReset.BackColor = "#FFC107"
$btnReset.Add_Click({
    Reset-Stats
    Write-Log "ͳ����Ϣ������" "Blue"
})
$form.Controls.Add($btnReset)

# Ӧ�ó���ر�ǰ�¼�
$form.Add_FormClosing({
    # ֹͣ��ʱ��
    if ($global:timer -ne $null) {
        $global:timer.Stop()
        $global:timer.Dispose()
    }
})

# ��ʾ������
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()