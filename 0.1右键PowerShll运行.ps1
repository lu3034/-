Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ����������
$form = New-Object System.Windows.Forms.Form
$form.Text = "�ļ����ݼ����V0.1"
$form.Size = New-Object System.Drawing.Size(650, 500)
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

# ��־��
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Location = New-Object System.Drawing.Point(20, 100)
$lblLog.Size = New-Object System.Drawing.Size(120, 20)
$lblLog.Text = "������־:"
$form.Controls.Add($lblLog)

$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 120)
$txtLog.Size = New-Object System.Drawing.Size(580, 280)
$txtLog.ReadOnly = $true
$txtLog.BackColor = "White"
$txtLog.ScrollBars = "Vertical"
$form.Controls.Add($txtLog)

# ״̬��ǩ
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(20, 410)
$lblStatus.Size = New-Object System.Drawing.Size(580, 20)
$lblStatus.Text = "״̬: ׼������"
$lblStatus.TextAlign = "MiddleLeft"
$form.Controls.Add($lblStatus)

# ������
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 440)
$progressBar.Size = New-Object System.Drawing.Size(580, 20)
$progressBar.Style = "Marquee"
$progressBar.MarqueeAnimationSpeed = 30
$progressBar.Visible = $false
$form.Controls.Add($progressBar)

# ��ʼ��ȫ�ֱ���
$global:isMonitoring = $false
$global:clipboardHistory = ""
$global:timer = $null

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
}

# �����ļ�/�ļ���
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
            Copy-Item -Path $sourcePath -Destination $destFolder -Recurse -Force
            
            return "�ļ����Ѹ���"
        }
        else {
            $fileName = Split-Path $sourcePath -Leaf
            $destFile = Join-Path -Path $targetPath -ChildPath $fileName
            
            # ���Ŀ���ļ��в������򴴽�
            if (-not (Test-Path -Path $targetPath)) {
                New-Item -ItemType Directory -Path $targetPath -Force | Out-Null
                Write-Log "����Ŀ���ļ���: $targetPath" "Green"
            }
            
            # �����ļ�
            Write-Log "���ڸ����ļ�: $fileName" "Blue"
            Copy-Item -Path $sourcePath -Destination $targetPath -Force
            
            return "�ļ��Ѹ���"
        }
    }
    catch {
        return "����: $($_.Exception.Message)"
    }
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
                Write-Log "���: $result ($(Split-Path $item -Leaf))" "Green"
            }
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
    $lblStatus.Text = "״̬: ����� (���: $($numInterval.Value)��)"
    $lblStatus.ForeColor = "Green"
    
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
    $global:timer.Add_Tick({ Monitor-Clipboard })
    $global:timer.Start()
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
    $lblStatus.Text = "״̬: �����ֹͣ"
    $lblStatus.ForeColor = "Red"
    
    # ���ؽ�����
    $progressBar.Visible = $false
    
    # ����ȫ�ֱ���
    $global:isMonitoring = $false
    
    Write-Log "�����ֹͣ"
    Write-Log "------------------------------------"
})

# Ӧ�ó���ر�ǰ�¼�
$form.Add_FormClosing({
    # ֹͣ��ʱ��
    if ($global:timer -ne $null) {
        $global:timer.Stop()
        $global:timer.Dispose()
    }
})

# ��ʾ������
$form.ShowDialog()