# Tối Ưu CUKCUK - PowerShell Tool
# Tạo bởi Phạm Ngọc Thạch đẹp trai - ONB GPBL HCM
# Build: 2024-06
# Mô tả: Công cụ GUI WinForms tối ưu Windows 10/11, mạng, bảo mật, sao lưu, kiểm tra, máy in
# Yêu cầu: PowerShell 5.1, Windows 10/11, chạy với quyền admin

Add-Type -AssemblyName System.Windows.Forms,System.Drawing,System.Management.Automation

# Kiểm tra môi trường
function Check-Environment {
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole] "Administrator")
    $psVersion = $PSVersionTable.PSVersion
    $osVersion = (Get-CimInstance Win32_OperatingSystem).Version
    $isApartment = $Host.Runspace.ApartmentState -eq 'STA'
    return [PSCustomObject]@{
        IsAdmin = $isAdmin
        PSVersion = $psVersion
        OSVersion = $osVersion
        IsApartmentSTA = $isApartment
    }
}

# Backup registry key
function Backup-RegistryKey {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$BackupFolder
    )
    try {
        $safePath = $Path -replace '[:\]', '_'
        $backupFile = Join-Path $BackupFolder "$safePath.reg"
        reg export $Path $backupFile /y | Out-Null
        return $backupFile
    } catch {
        return $null
    }
}

# Restore registry key
function Restore-RegistryKey {
    param(
        [Parameter(Mandatory)][string]$BackupFile
    )
    if (Test-Path $BackupFile) {
        reg import $BackupFile | Out-Null
        return $true
    }
    return $false
}

# GUI creation
function Create-GUI {
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Tối Ưu CUKCUK - Phạm Ngọc Thạch đẹp trai - ONB GPBL HCM"
    $form.Size = New-Object System.Drawing.Size(1000, 700)
    $form.MinimumSize = New-Object System.Drawing.Size(900, 600)
    $form.StartPosition = "CenterScreen"
    $form.BackColor = [System.Drawing.Color]::FromArgb(18,18,18)
    $form.ForeColor = [System.Drawing.Color]::White
    $form.Font = New-Object System.Drawing.Font("Segoe UI",10)

    # Sidebar panel
    $sidebar = New-Object System.Windows.Forms.Panel
    $sidebar.Dock = "Left"
    $sidebar.Width = 220
    $sidebar.BackColor = [System.Drawing.Color]::FromArgb(34,34,34)
    $form.Controls.Add($sidebar)

    # Buttons for tabs
    $tabs = @(
        @{Name='Tối ưu Windows'; Id='tabWindows'},
        @{Name='Mạng & DNS'; Id='tabNetwork'},
        @{Name='Bảo mật'; Id='tabSecurity'},
        @{Name='Sao lưu / Tải ứng dụng'; Id='tabBackup'},
        @{Name='Kiểm tra sau setup'; Id='tabCheck'},
        @{Name='Máy in'; Id='tabPrinter'}
    )
    $btnHeight = 50
    $btns = @()
    for ($i=0; $i -lt $tabs.Count; $i++) {
        $btn = New-Object System.Windows.Forms.Button
        $btn.Text = $tabs[$i].Name
        $btn.Name = $tabs[$i].Id
        $btn.Width = $sidebar.Width
        $btn.Height = $btnHeight
        $btn.Location = New-Object System.Drawing.Point(0, $i * $btnHeight)
        $btn.FlatStyle = "Flat"
        $btn.ForeColor = [System.Drawing.Color]::LightGray
        $btn.BackColor = [System.Drawing.Color]::FromArgb(34,34,34)
        $btn.FlatAppearance.BorderSize = 0
        $btn.Cursor = "Hand"
        $btn.Tag = $i
        $sidebar.Controls.Add($btn)
        $btns += $btn
    }

    # Highlight first button
    function Set-ActiveButton {
        param([System.Windows.Forms.Button]$activeBtn)
        foreach ($b in $btns) {
            if ($b -eq $activeBtn) {
                $b.BackColor = [System.Drawing.Color]::FromArgb(10,132,255)
                $b.ForeColor = [System.Drawing.Color]::White
                $b.FlatAppearance.BorderSize = 4
                $b.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(10,132,255)
            } else {
                $b.BackColor = [System.Drawing.Color]::FromArgb(34,34,34)
                $b.ForeColor = [System.Drawing.Color]::LightGray
                $b.FlatAppearance.BorderSize = 0
            }
        }
    }
    Set-ActiveButton $btns[0]

    # Main panel for content
    $mainPanel = New-Object System.Windows.Forms.Panel
    $mainPanel.Dock = "Fill"
    $mainPanel.BackColor = [System.Drawing.Color]::FromArgb(24,24,24)
    $form.Controls.Add($mainPanel)

    # Label header
    $headerLabel = New-Object System.Windows.Forms.Label
    $headerLabel.Text = $tabs[0].Name
    $headerLabel.Font = New-Object System.Drawing.Font("Segoe UI",14,[System.Drawing.FontStyle]::Bold)
    $headerLabel.ForeColor = [System.Drawing.Color]::White
    $headerLabel.Dock = "Top"
    $headerLabel.Height = 40
    $headerLabel.TextAlign = 'MiddleLeft'
    $headerLabel.Padding = '10,0,0,0'
    $mainPanel.Controls.Add($headerLabel)

    # Tab content panels
    $tabPanels = @{}
    foreach ($tab in $tabs) {
        $p = New-Object System.Windows.Forms.Panel
        $p.Dock = "Fill"
        $p.Visible = $false
        $mainPanel.Controls.Add($p)
        $tabPanels[$tab.Id] = $p
    }
    $tabPanels[$tabs[0].Id].Visible = $true

    # Example: Tối ưu Windows tab content
    $winPanel = $tabPanels['tabWindows']

    # Add a DataGridView for detailed table
    $dgv = New-Object System.Windows.Forms.DataGridView
    $dgv.Dock = "Fill"
    $dgv.BackgroundColor = [System.Drawing.Color]::FromArgb(30,30,30)
    $dgv.ForeColor = [System.Drawing.Color]::White
    $dgv.GridColor = [System.Drawing.Color]::FromArgb(50,50,50)
    $dgv.EnableHeadersVisualStyles = $false
    $dgv.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(40,40,40)
    $dgv.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $dgv.RowHeadersVisible = $false
    $dgv.SelectionMode = 'FullRowSelect'
    $dgv.MultiSelect = $false
    $dgv.AllowUserToAddRows = $false
    $dgv.AllowUserToDeleteRows = $false
    $dgv.ReadOnly = $true
    $dgv.Font = New-Object System.Drawing.Font("Consolas",9)
    $winPanel.Controls.Add($dgv)

    # Define columns: Nhóm, Tool sẽ làm gì, Tác dụng, Icon, Hành động, Kết quả, Lý do
    $cols = @(
        @{Name='Group'; Header='Nhóm'; Width=120},
        @{Name='Tool'; Header='Tool sẽ làm gì'; Width=280},
        @{Name='Effect'; Header='Tác dụng'; Width=280},
        @{Name='Icon'; Header='Icon'; Width=50},
        @{Name='Action'; Header='Hành động'; Width=120},
        @{Name='Result'; Header='Kết quả'; Width=120},
        @{Name='Reason'; Header='Lý do'; Width=280}
    )
    foreach ($col in $cols) {
        $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $column.Name = $col.Name
        $column.HeaderText = $col.Header
        $column.Width = $col.Width
        $dgv.Columns.Add($column) | Out-Null
    }

    # Sample data rows (17 mục tối ưu Windows)
    $rows = @(
        @{Group='Windows'; Tool='Tắt Activity history'; Effect='Giảm thu thập dữ liệu người dùng'; Icon='🛑'; Action='Tắt dịch vụ và registry'; Result='Thành công/Thất bại'; Reason='Chính sách bảo mật'},
        @{Group='Windows'; Tool='Tắt Consumer Features'; Effect='Tắt tính năng không cần thiết'; Icon='🚫'; Action='Tắt registry và dịch vụ'; Result='Thành công/Thất bại'; Reason='Tiết kiệm tài nguyên'},
        @{Group='Windows'; Tool='Tắt P2P Delivery Optimization'; Effect='Giảm băng thông mạng'; Icon='📡'; Action='Tắt dịch vụ'; Result='Thành công/Thất bại'; Reason='Giảm tải mạng'},
        @{Group='Windows'; Tool='Bật taskbar End task'; Effect='Tăng hiệu quả quản lý tác vụ'; Icon='⚙️'; Action='Thay đổi registry'; Result='Chỉ Windows hỗ trợ'; Reason='Tính năng mới'},
        @{Group='Windows'; Tool='Tắt folder type autodiscovery'; Effect='Tăng tốc Explorer'; Icon='📁'; Action='Tắt registry'; Result='Thành công'; Reason='Tối ưu Explorer'},
        @{Group='Windows'; Tool='Tắt location'; Effect='Bảo mật vị trí'; Icon='📍'; Action='Tắt dịch vụ và registry'; Result='Thành công'; Reason='Bảo mật'},
        @{Group='Windows'; Tool='Tắt Store suggestions tìm kiếm'; Effect='Loại bỏ quảng cáo'; Icon='🔍'; Action='Tắt policy và registry'; Result='Đã kiểm chứng'; Reason='Trải nghiệm người dùng'},
        @{Group='Windows'; Tool='Tạo restore point'; Effect='Phục hồi hệ thống'; Icon='🛡️'; Action='Tạo điểm khôi phục'; Result='Thành công'; Reason='An toàn'},
        @{Group='Windows'; Tool='Manual chỉ MapsBroker WerSvc SysMain TrkWks'; Effect='Tối ưu dịch vụ'; Icon='🔧'; Action='Tắt dịch vụ không cần thiết'; Result='Không tắt WSearch'; Reason='Ổn định hệ thống'},
        @{Group='Windows'; Tool='Start layout trước đây'; Effect='Khôi phục giao diện'; Icon='🖥️'; Action='Thay đổi registry'; Result='Chỉ khi đáng tin'; Reason='Tùy chỉnh giao diện'},
        @{Group='Windows'; Tool='Telemetry mức edition hỗ trợ'; Effect='Giảm thu thập dữ liệu'; Icon='📊'; Action='Tắt telemetry'; Result='Phân biệt theo edition'; Reason='Bảo mật'},
        @{Group='Windows'; Tool='Xóa temp >7 ngày'; Effect='Giải phóng dung lượng'; Icon='🗑️'; Action='Xóa file temp'; Result='Bỏ qua file đang dùng'; Reason='An toàn dữ liệu'},
        @{Group='Windows'; Tool='Tắt/gỡ Widgets WebExperience'; Effect='Tắt widget không cần thiết'; Icon='🧩'; Action='Gỡ package và tắt dịch vụ'; Result='Phạm vi current user'; Reason='Tối ưu trải nghiệm'},
        @{Group='Windows'; Tool='WPBT'; Effect='Bảo mật boot'; Icon='🔐'; Action='Kiểm tra và cấu hình'; Result='Đã xác minh'; Reason='Bảo mật hệ thống'},
        @{Group='Windows'; Tool='Gỡ OneDrive'; Effect='Loại bỏ đồng bộ không cần thiết'; Icon='☁️'; Action='Dùng uninstaller chính thức'; Result='Kiểm tra đồng bộ trước'; Reason='Bảo mật và hiệu năng'},
        @{Group='Windows'; Tool='Classic context menu'; Effect='Khôi phục menu cổ điển'; Icon='📋'; Action='Thay đổi registry'; Result='Thành công'; Reason='Tùy chỉnh giao diện'},
        @{Group='Windows'; Tool='AI xử lý Copilot và Recall'; Effect='Tùy chọn tính năng AI'; Icon='🤖'; Action='Bật/tắt feature theo tài liệu Microsoft'; Result='Phân biệt package đúng'; Reason='Tối ưu AI'}
    )
    foreach ($r in $rows) {
        $idx = $dgv.Rows.Add()
        $row = $dgv.Rows[$idx]
        $row.Cells['Group'].Value = $r.Group
        $row.Cells['Tool'].Value = $r.Tool
        $row.Cells['Effect'].Value = $r.Effect
        $row.Cells['Icon'].Value = $r.Icon
        $row.Cells['Action'].Value = $r.Action
        $row.Cells['Result'].Value = $r.Result
        $row.Cells['Reason'].Value = $r.Reason
    }

    # Event handlers for sidebar buttons
    foreach ($btn in $btns) {
        $btn.Add_Click({
            $sender = $_.Sender
            Set-ActiveButton $sender
            $headerLabel.Text = $tabs[$sender.Tag].Name
            foreach ($key in $tabPanels.Keys) {
                $tabPanels[$key].Visible = $false
            }
            $tabPanels[$sender.Name].Visible = $true
        })
    }

    # Return form object
    return $form
}

# Main execution
$envCheck = Check-Environment
if (-not $envCheck.IsAdmin) {
    [System.Windows.Forms.MessageBox]::Show("Vui lòng chạy PowerShell với quyền Administrator để sử dụng công cụ này.","Quyền truy cập","OK","Warning") | Out-Null
    exit
}
if ($envCheck.PSVersion.Major -lt 5) {
    [System.Windows.Forms.MessageBox]::Show("PowerShell phiên bản 5.1 trở lên là bắt buộc.","Phiên bản PowerShell","OK","Warning") | Out-Null
    exit
}
if (-not $envCheck.IsApartmentSTA) {
    [System.Windows.Forms.MessageBox]::Show("Runspace phải ở trạng thái STA để chạy GUI.","Runspace Apartment","OK","Warning") | Out-Null
    exit
}

$form = Create-GUI
[void]$form.ShowDialog()

# Kết thúc script
# Tối Ưu CUKCUK - PowerShell Tool
# Tạo bởi Phạm Ngọc Thạch đẹp trai - ONB GPBL HCM
# Build: 2024-06
# Mô tả: Công cụ GUI WinForms tối ưu Windows 10/11, mạng, bảo mật, sao lưu, kiểm tra, máy in
# Yêu cầu: PowerShell 5.1, Windows 10/11, chạy với quyền admin

Add-Type -AssemblyName System.Windows.Forms,System.Drawing,System.Management.Automation

# Kiểm tra môi trường
function Check-Environment {
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole] "Administrator")
    $psVersion = $PSVersionTable.PSVersion
    $osVersion = (Get-CimInstance Win32_OperatingSystem).Version
    $isApartment = $Host.Runspace.ApartmentState -eq 'STA'
    return [PSCustomObject]@{
        IsAdmin = $isAdmin
        PSVersion = $psVersion
        OSVersion = $osVersion
        IsApartmentSTA = $isApartment
    }
}

# Backup registry key
function Backup-RegistryKey {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$BackupFolder
    )
    try {
        $safePath = $Path -replace '[:\]', '_'
        $backupFile = Join-Path $BackupFolder "$safePath.reg"
        reg export $Path $backupFile /y | Out-Null
        return $backupFile
    } catch {
        return $null
    }
}

# Restore registry key
function Restore-RegistryKey {
    param(
        [Parameter(Mandatory)][string]$BackupFile
    )
    if (Test-Path $BackupFile) {
        reg import $BackupFile | Out-Null
        return $true
    }
    return $false
}

# GUI creation
function Create-GUI {
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Tối Ưu CUKCUK - Phạm Ngọc Thạch đẹp trai - ONB GPBL HCM"
    $form.Size = New-Object System.Drawing.Size(1000, 700)
    $form.MinimumSize = New-Object System.Drawing.Size(900, 600)
    $form.StartPosition = "CenterScreen"
    $form.BackColor = [System.Drawing.Color]::FromArgb(18,18,18)
    $form.ForeColor = [System.Drawing.Color]::White
    $form.Font = New-Object System.Drawing.Font("Segoe UI",10)

    # Sidebar panel
    $sidebar = New-Object System.Windows.Forms.Panel
    $sidebar.Dock = "Left"
    $sidebar.Width = 220
    $sidebar.BackColor = [System.Drawing.Color]::FromArgb(34,34,34)
    $form.Controls.Add($sidebar)

    # Buttons for tabs
    $tabs = @(
        @{Name='Tối ưu Windows'; Id='tabWindows'},
        @{Name='Mạng & DNS'; Id='tabNetwork'},
        @{Name='Bảo mật'; Id='tabSecurity'},
        @{Name='Sao lưu / Tải ứng dụng'; Id='tabBackup'},
        @{Name='Kiểm tra sau setup'; Id='tabCheck'},
        @{Name='Máy in'; Id='tabPrinter'}
    )
    $btnHeight = 50
    $btns = @()
    for ($i=0; $i -lt $tabs.Count; $i++) {
        $btn = New-Object System.Windows.Forms.Button
        $btn.Text = $tabs[$i].Name
        $btn.Name = $tabs[$i].Id
        $btn.Width = $sidebar.Width
        $btn.Height = $btnHeight
        $btn.Location = New-Object System.Drawing.Point(0, $i * $btnHeight)
        $btn.FlatStyle = "Flat"
        $btn.ForeColor = [System.Drawing.Color]::LightGray
        $btn.BackColor = [System.Drawing.Color]::FromArgb(34,34,34)
        $btn.FlatAppearance.BorderSize = 0
        $btn.Cursor = "Hand"
        $btn.Tag = $i
        $sidebar.Controls.Add($btn)
        $btns += $btn
    }

    # Highlight first button
    function Set-ActiveButton {
        param([System.Windows.Forms.Button]$activeBtn)
        foreach ($b in $btns) {
            if ($b -eq $activeBtn) {
                $b.BackColor = [System.Drawing.Color]::FromArgb(10,132,255)
                $b.ForeColor = [System.Drawing.Color]::White
                $b.FlatAppearance.BorderSize = 4
                $b.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(10,132,255)
            } else {
                $b.BackColor = [System.Drawing.Color]::FromArgb(34,34,34)
                $b.ForeColor = [System.Drawing.Color]::LightGray
                $b.FlatAppearance.BorderSize = 0
            }
        }
    }
    Set-ActiveButton $btns[0]

    # Main panel for content
    $mainPanel = New-Object System.Windows.Forms.Panel
    $mainPanel.Dock = "Fill"
    $mainPanel.BackColor = [System.Drawing.Color]::FromArgb(24,24,24)
    $form.Controls.Add($mainPanel)

    # Label header
    $headerLabel = New-Object System.Windows.Forms.Label
    $headerLabel.Text = $tabs[0].Name
    $headerLabel.Font = New-Object System.Drawing.Font("Segoe UI",14,[System.Drawing.FontStyle]::Bold)
    $headerLabel.ForeColor = [System.Drawing.Color]::White
    $headerLabel.Dock = "Top"
    $headerLabel.Height = 40
    $headerLabel.TextAlign = 'MiddleLeft'
    $headerLabel.Padding = '10,0,0,0'
    $mainPanel.Controls.Add($headerLabel)

    # Tab content panels
    $tabPanels = @{}
    foreach ($tab in $tabs) {
        $p = New-Object System.Windows.Forms.Panel
        $p.Dock = "Fill"
        $p.Visible = $false
        $mainPanel.Controls.Add($p)
        $tabPanels[$tab.Id] = $p
    }
    $tabPanels[$tabs[0].Id].Visible = $true

    # Example: Tối ưu Windows tab content
    $winPanel = $tabPanels['tabWindows']

    # Add a DataGridView for detailed table
    $dgv = New-Object System.Windows.Forms.DataGridView
    $dgv.Dock = "Fill"
    $dgv.BackgroundColor = [System.Drawing.Color]::FromArgb(30,30,30)
    $dgv.ForeColor = [System.Drawing.Color]::White
    $dgv.GridColor = [System.Drawing.Color]::FromArgb(50,50,50)
    $dgv.EnableHeadersVisualStyles = $false
    $dgv.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(40,40,40)
    $dgv.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $dgv.RowHeadersVisible = $false
    $dgv.SelectionMode = 'FullRowSelect'
    $dgv.MultiSelect = $false
    $dgv.AllowUserToAddRows = $false
    $dgv.AllowUserToDeleteRows = $false
    $dgv.ReadOnly = $true
    $dgv.Font = New-Object System.Drawing.Font("Consolas",9)
    $winPanel.Controls.Add($dgv)

    # Define columns: Nhóm, Tool sẽ làm gì, Tác dụng, Icon, Hành động, Kết quả, Lý do
    $cols = @(
        @{Name='Group'; Header='Nhóm'; Width=120},
        @{Name='Tool'; Header='Tool sẽ làm gì'; Width=280},
        @{Name='Effect'; Header='Tác dụng'; Width=280},
        @{Name='Icon'; Header='Icon'; Width=50},
        @{Name='Action'; Header='Hành động'; Width=120},
        @{Name='Result'; Header='Kết quả'; Width=120},
        @{Name='Reason'; Header='Lý do'; Width=280}
    )
    foreach ($col in $cols) {
        $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $column.Name = $col.Name
        $column.HeaderText = $col.Header
        $column.Width = $col.Width
        $dgv.Columns.Add($column) | Out-Null
    }

    # Sample data rows (17 mục tối ưu Windows)
    $rows = @(
        @{Group='Windows'; Tool='Tắt Activity history'; Effect='Giảm thu thập dữ liệu người dùng'; Icon='🛑'; Action='Tắt dịch vụ và registry'; Result='Thành công/Thất bại'; Reason='Chính sách bảo mật'},
        @{Group='Windows'; Tool='Tắt Consumer Features'; Effect='Tắt tính năng không cần thiết'; Icon='🚫'; Action='Tắt registry và dịch vụ'; Result='Thành công/Thất bại'; Reason='Tiết kiệm tài nguyên'},
        @{Group='Windows'; Tool='Tắt P2P Delivery Optimization'; Effect='Giảm băng thông mạng'; Icon='📡'; Action='Tắt dịch vụ'; Result='Thành công/Thất bại'; Reason='Giảm tải mạng'},
        @{Group='Windows'; Tool='Bật taskbar End task'; Effect='Tăng hiệu quả quản lý tác vụ'; Icon='⚙️'; Action='Thay đổi registry'; Result='Chỉ Windows hỗ trợ'; Reason='Tính năng mới'},
        @{Group='Windows'; Tool='Tắt folder type autodiscovery'; Effect='Tăng tốc Explorer'; Icon='📁'; Action='Tắt registry'; Result='Thành công'; Reason='Tối ưu Explorer'},
        @{Group='Windows'; Tool='Tắt location'; Effect='Bảo mật vị trí'; Icon='📍'; Action='Tắt dịch vụ và registry'; Result='Thành công'; Reason='Bảo mật'},
        @{Group='Windows'; Tool='Tắt Store suggestions tìm kiếm'; Effect='Loại bỏ quảng cáo'; Icon='🔍'; Action='Tắt policy và registry'; Result='Đã kiểm chứng'; Reason='Trải nghiệm người dùng'},
        @{Group='Windows'; Tool='Tạo restore point'; Effect='Phục hồi hệ thống'; Icon='🛡️'; Action='Tạo điểm khôi phục'; Result='Thành công'; Reason='An toàn'},
        @{Group='Windows'; Tool='Manual chỉ MapsBroker WerSvc SysMain TrkWks'; Effect='Tối ưu dịch vụ'; Icon='🔧'; Action='Tắt dịch vụ không cần thiết'; Result='Không tắt WSearch'; Reason='Ổn định hệ thống'},
        @{Group='Windows'; Tool='Start layout trước đây'; Effect='Khôi phục giao diện'; Icon='🖥️'; Action='Thay đổi registry'; Result='Chỉ khi đáng tin'; Reason='Tùy chỉnh giao diện'},
        @{Group='Windows'; Tool='Telemetry mức edition hỗ trợ'; Effect='Giảm thu thập dữ liệu'; Icon='📊'; Action='Tắt telemetry'; Result='Phân biệt theo edition'; Reason='Bảo mật'},
        @{Group='Windows'; Tool='Xóa temp >7 ngày'; Effect='Giải phóng dung lượng'; Icon='🗑️'; Action='Xóa file temp'; Result='Bỏ qua file đang dùng'; Reason='An toàn dữ liệu'},
        @{Group='Windows'; Tool='Tắt/gỡ Widgets WebExperience'; Effect='Tắt widget không cần thiết'; Icon='🧩'; Action='Gỡ package và tắt dịch vụ'; Result='Phạm vi current user'; Reason='Tối ưu trải nghiệm'},
        @{Group='Windows'; Tool='WPBT'; Effect='Bảo mật boot'; Icon='🔐'; Action='Kiểm tra và cấu hình'; Result='Đã xác minh'; Reason='Bảo mật hệ thống'},
        @{Group='Windows'; Tool='Gỡ OneDrive'; Effect='Loại bỏ đồng bộ không cần thiết'; Icon='☁️'; Action='Dùng uninstaller chính thức'; Result='Kiểm tra đồng bộ trước'; Reason='Bảo mật và hiệu năng'},
        @{Group='Windows'; Tool='Classic context menu'; Effect='Khôi phục menu cổ điển'; Icon='📋'; Action='Thay đổi registry'; Result='Thành công'; Reason='Tùy chỉnh giao diện'},
        @{Group='Windows'; Tool='AI xử lý Copilot và Recall'; Effect='Tùy chọn tính năng AI'; Icon='🤖'; Action='Bật/tắt feature theo tài liệu Microsoft'; Result='Phân biệt package đúng'; Reason='Tối ưu AI'}
    )
    foreach ($r in $rows) {
        $idx = $dgv.Rows.Add()
        $row = $dgv.Rows[$idx]
        $row.Cells['Group'].Value = $r.Group
        $row.Cells['Tool'].Value = $r.Tool
        $row.Cells['Effect'].Value = $r.Effect
        $row.Cells['Icon'].Value = $r.Icon
        $row.Cells['Action'].Value = $r.Action
        $row.Cells['Result'].Value = $r.Result
        $row.Cells['Reason'].Value = $r.Reason
    }

    # Event handlers for sidebar buttons
    foreach ($btn in $btns) {
        $btn.Add_Click({
            $sender = $_.Sender
            Set-ActiveButton $sender
            $headerLabel.Text = $tabs[$sender.Tag].Name
            foreach ($key in $tabPanels.Keys) {
                $tabPanels[$key].Visible = $false
            }
            $tabPanels[$sender.Name].Visible = $true
        })
    }

    # Return form object
    return $form
}

# Main execution
$envCheck = Check-Environment
if (-not $envCheck.IsAdmin) {
    [System.Windows.Forms.MessageBox]::Show("Vui lòng chạy PowerShell với quyền Administrator để sử dụng công cụ này.","Quyền truy cập","OK","Warning") | Out-Null
    exit
}
if ($envCheck.PSVersion.Major -lt 5) {
    [System.Windows.Forms.MessageBox]::Show("PowerShell phiên bản 5.1 trở lên là bắt buộc.","Phiên bản PowerShell","OK","Warning") | Out-Null
    exit
}
if (-not $envCheck.IsApartmentSTA) {
    [System.Windows.Forms.MessageBox]::Show("Runspace phải ở trạng thái STA để chạy GUI.","Runspace Apartment","OK","Warning") | Out-Null
    exit
}

$form = Create-GUI
[void]$form.ShowDialog()

# Kết thúc script
