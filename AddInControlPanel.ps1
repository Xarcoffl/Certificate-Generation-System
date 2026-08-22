Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ProjectDir = $PSScriptRoot
$Port = 3000

# Simple TCP probe avoids depending on the NetTCPIP module being present.
function Test-PortOpen {
    param([int]$Port)
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $result = $client.BeginConnect("127.0.0.1", $Port, $null, $null)
        $connected = $result.AsyncWaitHandle.WaitOne(300, $false) -and $client.Connected
        $client.Close()
        return $connected
    } catch {
        return $false
    }
}

function Invoke-NpmCommand {
    param([string]$NpmArgs)
    Start-Process -FilePath "cmd.exe" -ArgumentList "/c cd /d ""$ProjectDir"" && npm $NpmArgs" -WindowStyle Normal
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Certificate Add-in Control Panel"
$form.Size = New-Object System.Drawing.Size(360, 200)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Status: Checking..."
$statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$statusLabel.AutoSize = $true
$statusLabel.Location = New-Object System.Drawing.Point(20, 20)
$form.Controls.Add($statusLabel)

$startButton = New-Object System.Windows.Forms.Button
$startButton.Text = "Start"
$startButton.Size = New-Object System.Drawing.Size(140, 40)
$startButton.Location = New-Object System.Drawing.Point(20, 70)
$form.Controls.Add($startButton)

$stopButton = New-Object System.Windows.Forms.Button
$stopButton.Text = "Stop"
$stopButton.Size = New-Object System.Drawing.Size(140, 40)
$stopButton.Location = New-Object System.Drawing.Point(180, 70)
$form.Controls.Add($stopButton)

$logLabel = New-Object System.Windows.Forms.Label
$logLabel.Text = ""
$logLabel.AutoSize = $true
$logLabel.ForeColor = [System.Drawing.Color]::Gray
$logLabel.Location = New-Object System.Drawing.Point(20, 130)
$form.Controls.Add($logLabel)

function Update-Status {
    if (Test-PortOpen -Port $Port) {
        $statusLabel.Text = "Status: Running"
        $statusLabel.ForeColor = [System.Drawing.Color]::Green
        $startButton.Enabled = $false
        $stopButton.Enabled = $true
    } else {
        $statusLabel.Text = "Status: Stopped"
        $statusLabel.ForeColor = [System.Drawing.Color]::Red
        $startButton.Enabled = $true
        $stopButton.Enabled = $false
    }
}

$startButton.Add_Click({
    $logLabel.Text = "Starting dev server..."
    Invoke-NpmCommand -NpmArgs "start"
    Start-Sleep -Seconds 2
    Update-Status
})

$stopButton.Add_Click({
    $logLabel.Text = "Stopping dev server..."
    Invoke-NpmCommand -NpmArgs "stop"
    Start-Sleep -Seconds 2
    Update-Status
    $logLabel.Text = ""
})

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 3000
$timer.Add_Tick({ Update-Status })
$timer.Start()

Update-Status
[void]$form.ShowDialog()
