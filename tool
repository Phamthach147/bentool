RowPanel 2
    [void]$actions.Children.Add((Button '↻  Lấy IP hiện tại' {
        StartTask 'Network' 'Tổng quan mạng' @{} 'Read' $true
    }))
    [void]$actions.Children.Add((Button 'Áp dụng IP tĩnh' {
        $a=SelectedAdapter
        if ($null -eq $a) { return }
        if (Confirm (
            "Đổi IP card: $($a.Name)`n" +
            "IP: $($script:IPBox.Text)/$($script:PrefixBox.Text)`n" +
            "Gateway: $($script:GatewayBox.Text)`n`n" +
            "Có thể mất kết nối từ xa. Đây là IP theo card, không theo SSID.`n" +
            "Card nhiều IPv4 sẽ không được tool sửa tự động.`n" +
            "Anh đã kiểm tra IP chưa bị thiết bị khác sử dụng và đồng ý tiếp tục?"
        )) {
            $d=AdapterData $a
            $d.IP=$script:IPBox.Text.Trim()
            $d.Prefix=$script:PrefixBox.Text.Trim()
            $d.Gateway=$script:GatewayBox.Text.Trim()
            StartTask 'StaticIP' 'Đặt IP tĩnh' $d 'Config'
        }
    } '#075985'))
    [void]$actions.Children.Add((Button 'Trả IP/DNS về DHCP' {
        $a=SelectedAdapter
        if ($null -eq $a) { return }
        if (Confirm ("Trả IP và DNS card $($a.Name) về tự động?`nCó thể mất kết nối. Card nhiều IPv4 cần xử lý thủ công.")) {
            StartTask 'DHCP' 'Trả IP/DNS về DHCP' (AdapterData $a) 'Config'
        }
    }))
    [void]$actions.Children.Add((Button 'Cấu hình Windows' {
        Start-Process 'ncpa.cpl'
    }))
    [void]$s.Children.Add($actions)
    [void]$s.Children.Add((Button 'Ưu tiên Wi-Fi 5 GHz' {
        $a=SelectedAdapter
        if ($null -eq $a) { return }
        if (Confirm ("Ưu tiên 5 GHz trên card $($a.Name)?`nKhông tắt 2.4 GHz, không tự restart card. Chỉ làm khi driver có lựa chọn rõ.")) {
            StartTask 'WiFi5' 'Ưu tiên Wi-Fi 5 GHz' (AdapterData $a) 'Config'
        }
    }))
    [void]$s.Children.Add((Text 'IP theo từng Wi-Fi: ưu tiên DHCP reservation trên router. Tool không cài bộ theo dõi SSID chạy nền.' 12 '#FCD34D'))

    $script:AdapterBox.Add_SelectionChanged({
        $a=$script:AdapterBox.SelectedItem
        if ($null -eq $a) { return }
        $all=@($a.IPs)
        $chosen=$all | Where-Object IPAddress -eq $a.Source | Select-Object -First 1
        if ($null -eq $chosen -and $all.Count -eq 1) { $chosen=$all[0] }
        $script:IPBox.Text=''
        if ($chosen) {
            $script:IPBox.Text=[string]$chosen.IPAddress
            $script:PrefixBox.Text=[string]$chosen.PrefixLength
        }
        $script:GatewayBox.Text=[string]$a.Gateway
        $script:NetInfo.Text='Card: {0} | MAC: {1} | IPv4: {2}' -f
            $a.Name,$a.Mac,
            (($all | ForEach-Object { '{0}/{1}' -f $_.IPAddress,$_.PrefixLength }) -join ', ')
    })

    $s=$script:NetPanels['DNS']
    [void]$s.Children.Add((Text 'Chọn nhà cung cấp không tự thay đổi DNS. Bấm Áp dụng mới thực hiện.' 13))
    $script:DNSMode=Combo
    foreach ($mode in @('Tự nhập','Cloudflare','Google','Nhanh nhất','Tự động / DHCP')) {
        [void]$script:DNSMode.Items.Add($mode)
    }
    $script:DNSMode.SelectedIndex=0
    [void]$s.Children.Add($script:DNSMode)
    [void]$s.Children.Add((Text 'DNS chính'))
    $script:DNS1=Box
    [void]$s.Children.Add($script:DNS1)
    [void]$s.Children.Add((Text 'DNS phụ'))
    $script:DNS2=Box
    [void]$s.Children.Add($script:DNS2)
    $Xin lỗi ông, lần này tôi viết **mã PowerShell trực tiếp bên dưới**, không tạo file TXT.

**Nói rõ trước khi copy:** tôi chưa có môi trường Windows để chạy thử hoặc xác nhận bằng trình kiểm tra cú pháp. Bản này có giao diện và các tác vụ thực, nhưng **chưa đáp ứng trọn vẹn đặc tả**: chưa có lệnh cấu hình nâng cao riêng cho XP-80C, chưa tự đổi IP khi card có nhiều IPv4, chưa khôi phục tự động mọi thay đổi. Những trường hợp đó được thông báo rõ, không báo thành công giả. DNS và Recall sử dụng hướng xử lý được Microsoft tài liệu hóa. ([learn.microsoft.com](https://learn.microsoft.com/en-us/powershell/module/dnsclient/set-dnsclientserveraddress?view=windowsserver2025-ps))

**Copy toàn bộ khối dưới đây** vào file `toolcukcuk` trên GitHub, thay toàn bộ nội dung cũ:

```powershell
#requires -Version 5.1
<#
TOI UU CUKCUK
Build: 2026.10.24-r1
Nen tang muc tieu: Windows PowerShell 5.1, Windows 10/11.

CHUA DUOC CHAY THU HOAC KIEM TRA PARSER TREN WINDOWS.

Pham vi:
- WPF UI, runspace nen, hang doi ket qua ve UI.
- Khong tu thay doi cau hinh khi mo.
- Khong goi Start-Job.
- Khong tat Defender, Windows Update, UltraViewer.
- Backup Registry truoc khi sua, co khoi phuc Registry.
- Backup mang, dich vu, firewall de doi chieu.
- IP tinh chi tu dong xu ly card co dung mot IPv4.
- Khong tu xoa IP phu.
- Chua co giao thuc cau hinh nang cao rieng XP-80C.
#>

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------
# 1. KIEM TRA MOI TRUONG
# ---------------------------------------------------------------------

if ($env:OS -ne 'Windows_NT') {
    throw 'Tool chỉ chạy trên Windows.'
}

if ($PSVersionTable.PSEdition -ne 'Desktop' -or
    $PSVersionTable.PSVersion.Major -ne 5) {
    throw 'Hãy chạy bằng Windows PowerShell 5.1, không dùng PowerShell 7.'
}

if ([Environment]::Is64BitOperatingSystem -and
    -not [Environment]::Is64BitProcess) {
    throw @'
Đang chạy PowerShell 32-bit trên Windows 64-bit.
Hãy mở Windows PowerShell 64-bit bằng quyền Administrator.
Tool chưa mở và chưa thay đổi cấu hình.
'@
}

$ccIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
$ccPrincipal = New-Object Security.Principal.WindowsPrincipal($ccIdentity)

if (-not $ccPrincipal.IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator
)) {
    throw 'Hãy mở Windows PowerShell bằng Run as administrator rồi chạy lại.'
}

if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw 'Hãy mở powershell.exe -STA bằng quyền Administrator rồi chạy lại.'
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

$script:CCBuild = '2026.10.24-r1'
$script:CCRoot = Join-Path $env:LOCALAPPDATA 'ToiUuCUKCUK'
$script:CCSession = Join-Path $script:CCRoot (
    '{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'),
    ([guid]::NewGuid().ToString('N').Substring(0, 6))
)

[void][IO.Directory]::CreateDirectory($script:CCSession)

$script:CCQueue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
$script:CCJobs = @{}
$script:CCAdapters = @()
$script:CCApps = @()
$script:CCSelectedInterface = $null
$script:CCClosing = $false

# ---------------------------------------------------------------------
# 2. WORKER
# Moi tac vu co runspace rieng.
# Worker khong nhan control WPF.
# ---------------------------------------------------------------------

$script:CCWorker = @'
param(
    [string]$Task,
    [hashtable]$P,
    $Queue,
    [string]$Folder,
    $Cancel
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

function Emit {
    param([string]$Kind, $Data)
    $Queue.Enqueue([pscustomobject]@{
        Task = $Task
        Kind = $Kind
        Data = $Data
    })
}

function Row {
    param(
        [string]$Action,
        [string]$Status,
        [string]$Reason = ''
    )
    $symbol = 'i'
    if ($Status -eq 'Thành công') { $symbol = '✓' }
    if ($Status -eq 'Thất bại') { $symbol = '✕' }
    if ($Status -match 'khởi động|Chờ|Cần') { $symbol = '!' }

    Emit 'row' ([pscustomobject][ordered]@{
        'Biểu tượng' = $symbol
        'Hành động' = $Action
        'Kết quả' = $Status
        'Lý do / Việc cần làm' = $Reason
    })
}

function Step {
    param([string]$Name, [scriptblock]$Body)
    if ($Cancel.Requested) {
        Row $Name 'Đã dừng' 'Chưa thực hiện.'
        return
    }
    Emit 'log' $Name
    try {
        & $Body
    }
    catch {
        Row $Name 'Thất bại' $_.Exception.Message
    }
}

function Save-Snapshot {
    param([string]$Name, $Value)
    $path = Join-Path $Folder (
        '{0}-{1}-{2}.clixml' -f
        $Name,
        (Get-Date -Format 'HHmmssfff'),
        ([guid]::NewGuid().ToString('N').Substring(0, 5))
    )
    $Value | Export-Clixml -LiteralPath $path -Depth 12
    return $path
}

function Open-RegistryBase {
    param([string]$Hive)

    $view = [Microsoft.Win32.RegistryView]::Registry32
    if ([Environment]::Is64BitOperatingSystem) {
        $view = [Microsoft.Win32.RegistryView]::Registry64
    }

    $h = [Microsoft.Win32.RegistryHive]::LocalMachine
    if ($Hive -eq 'HKCU') {
        $h = [Microsoft.Win32.RegistryHive]::CurrentUser
    }

    return [Microsoft.Win32.RegistryKey]::OpenBaseKey($h, $view)
}

function Set-BackedRegistry {
    param(
        [string]$Hive,
        [string]$Path,
        [string]$Name,
        $Value,
        [string]$Kind = 'DWord'
    )

    $base = Open-RegistryBase $Hive
    $key = $null

    try {
        $key = $base.OpenSubKey($Path, $false)
        $exists = $false
        $oldValue = $null
        $oldKind = $null

        if ($null -ne $key) {
            $exists = @($key.GetValueNames()) -contains $Name
            if ($exists) {
                $oldKind = $key.GetValueKind($Name).ToString()
                $oldValue = $key.GetValue(
                    $Name,
                    $null,
                    [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames
                )
            }
            $key.Dispose()
            $key = $null
        }

        $record = [pscustomobject]@{
            Hive = $Hive
            Path = $Path
            Name = $Name
            Exists = $exists
            Value = $oldValue
            Kind = $oldKind
            Time = [DateTime]::UtcNow
        }

        [void](Save-Snapshot 'registry' $record)

        $key = $base.CreateSubKey($Path)
        $type = [Microsoft.Win32.RegistryValueKind]::$Kind
        $key.SetValue($Name, $Value, $type)

        $read = $key.GetValue($Name)

        if ($Kind -eq 'DWord') {
            if ([int]$read -ne [int]$Value) {
                throw 'Giá trị đọc lại không khớp.'
            }
        }
        elseif ([string]$read -ne [string]$Value) {
            throw 'Giá trị đọc lại không khớp.'
        }
    }
    finally {
        if ($null -ne $key) { $key.Dispose() }
        $base.Dispose()
    }
}

function Policy-Result {
    param([string]$Name)
    Row $Name 'Đã cấu hình, cần khởi động lại' (
        'Đã sao lưu và đọc lại giá trị. Hiệu lực phụ thuộc phiên bản, edition và chính sách quản lý máy.'
    )
}

function Read-Adapters {
    $bestIndex = $null

    try {
        $best = Find-NetRoute -RemoteIPAddress '1.1.1.1' -ErrorAction Stop
        $bestIndex = @(
            $best |
            Where-Object { $null -ne $_.InterfaceIndex } |
            Select-Object -First 1
        )[0].InterfaceIndex
    }
    catch {
        $routes = @(
            Get-NetRoute -AddressFamily IPv4 -DestinationPrefix '0.0.0.0/0' |
            Sort-Object @{
                Expression = {
                    $ipif = Get-NetIPInterface -InterfaceIndex $_.InterfaceIndex `
                        -AddressFamily IPv4 -ErrorAction SilentlyContinue
                    [int]$_.RouteMetric + [int]$ipif.InterfaceMetric
                }
            }
        )
        if ($routes.Count -gt 0) {
            $bestIndex = $routes[0].InterfaceIndex
        }
    }

    $list = @()

    foreach ($a in @(Get-NetAdapter | Sort-Object InterfaceIndex)) {
        $ips = @(
            Get-NetIPAddress -InterfaceIndex $a.InterfaceIndex `
                -AddressFamily IPv4 -ErrorAction SilentlyContinue |
            Where-Object { $_.IPAddress -ne '127.0.0.1' }
        )

        $routes = @(
            Get-NetRoute -InterfaceIndex $a.InterfaceIndex `
                -AddressFamily IPv4 -DestinationPrefix '0.0.0.0/0' `
                -ErrorAction SilentlyContinue |
            Sort-Object RouteMetric
        )

        $dns = @(
            (Get-DnsClientServerAddress -InterfaceIndex $a.InterfaceIndex `
                -AddressFamily IPv4 -ErrorAction SilentlyContinue).ServerAddresses
        )

        $ipif = Get-NetIPInterface -InterfaceIndex $a.InterfaceIndex `
            -AddressFamily IPv4 -ErrorAction SilentlyContinue

        $type = 'LAN / card khác'
        if ([int]$a.NdisPhysicalMedium -eq 9 -or
            [int]$a.NdisPhysicalMedium -eq 1) {
            $type = 'Wi-Fi'
        }
        elseif (-not $a.HardwareInterface) {
            $type = 'VPN / card ảo'
        }

        $gw = ''
        if ($routes.Count -gt 0) { $gw = $routes[0].NextHop }

        $list += [pscustomobject]@{
            Index = [int]$a.InterfaceIndex
            Guid = [string]$a.InterfaceGuid
            Name = $a.Name
            Display = '{0} | {1} | {2}' -f $a.InterfaceIndex, $a.Name, $a.Status
            Type = $type
            Status = [string]$a.Status
            MAC = $a.MacAddress
            IPs = @($ips | Select-Object IPAddress, PrefixLength, PrefixOrigin, SkipAsSource)
            Gateway = $gw
            DNS = $dns
            DHCP = [string]$ipif.Dhcp
            Preferred = ($a.InterfaceIndex -eq $bestIndex)
            Hardware = [bool]$a.HardwareInterface
        }
    }

    return $list
}

function Assert-Adapter {
    param([int]$Index, [string]$Guid)

    $a = Get-NetAdapter -InterfaceIndex $Index -ErrorAction Stop
    if ([string]$a.InterfaceGuid -ne $Guid) {
        throw 'Card mạng không còn khớp với lần đọc trước. Hãy làm mới.'
    }
    return $a
}

function Save-Network {
    param([int]$Index)

    $a = Get-NetAdapter -InterfaceIndex $Index
    $guid = '{' + ([guid]$a.InterfaceGuid).ToString() + '}'
    $reg = 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\' + $guid
    $nameServer = (Get-ItemProperty -LiteralPath $reg `
        -Name NameServer -ErrorAction SilentlyContinue).NameServer

    $data = [pscustomobject]@{
        InterfaceIndex = $Index
        InterfaceGuid = [string]$a.InterfaceGuid
        Name = $a.Name
        IP = @(Get-NetIPAddress -InterfaceIndex $Index -AddressFamily IPv4)
        Interface = Get-NetIPInterface -InterfaceIndex $Index -AddressFamily IPv4
        Routes = @(Get-NetRoute -InterfaceIndex $Index -AddressFamily IPv4)
        DNS = @(
            (Get-DnsClientServerAddress -InterfaceIndex $Index `
                -AddressFamily IPv4).ServerAddresses
        )
        DNSWasAutomatic = [string]::IsNullOrWhiteSpace([string]$nameServer)
    }

    return Save-Snapshot 'network' $data
}

function Parse-IPv4 {
    param([string]$Text)

    if ($Text -notmatch '^\d{1,3}(\.\d{1,3}){3}$') {
        throw ('IPv4 không hợp lệ: {0}' -f $Text)
    }

    $bytes = @()
    foreach ($part in $Text.Split('.')) {
        $n = [int]$part
        if ($n -lt 0 -or $n -gt 255) {
            throw ('IPv4 không hợp lệ: {0}' -f $Text)
        }
        $bytes += [byte]$n
    }

    return [Net.IPAddress]::new([byte[]]$bytes)
}

function IPv4-Number {
    param([string]$Text)

    $ip = Parse-IPv4 $Text
    $b = $ip.GetAddressBytes()

    return [uint64](
        [uint64]$b[0] * 16777216 +
        [uint64]$b[1] * 65536 +
        [uint64]$b[2] * 256 +
        [uint64]$b[3]
    )
}

function Validate-Static {
    param([string]$IP, [int]$Prefix, [string]$Gateway)

    if ($Prefix -lt 1 -or $Prefix -gt 30) {
        throw 'Bản này chỉ hỗ trợ prefix từ 1 đến 30.'
    }

    $n = IPv4-Number $IP
    $g = IPv4-Number $Gateway
    $first = (Parse-IPv4 $IP).GetAddressBytes()[0]
    $gfirst = (Parse-IPv4 $Gateway).GetAddressBytes()[0]

    if ($first -eq 0 -or $first -eq 127 -or $first -ge 224 -or
        $gfirst -eq 0 -or $gfirst -eq 127 -or $gfirst -ge 224) {
        throw 'IP hoặc gateway không phải địa chỉ unicast phù hợp.'
    }

    $size = [uint64][Math]::Pow(2, 32 - $Prefix)
    $network = [uint64]([Math]::Floor($n / $size) * $size)
    $broadcast = $network + $size - 1

    if ($n -eq $network -or $n -eq $broadcast) {
        throw 'IP là địa chỉ network hoặc broadcast.'
    }
    if ($g -le $network -or $g -ge $broadcast -or $g -eq $n) {
        throw 'Gateway phải khác IP và nằm trong cùng subnet.'
    }
}

function Test-Tcp {
    param([string]$IP, [int]$Port, [int]$Timeout = 500)

    $client = New-Object Net.Sockets.TcpClient
    $async = $null

    try {
        $async = $client.BeginConnect($IP, $Port, $null, $null)
        if (-not $async.AsyncWaitHandle.WaitOne($Timeout)) {
            return $false
        }
        $client.EndConnect($async)
        return $client.Connected
    }
    catch {
        return $false
    }
    finally {
        if ($null -ne $async) { $async.AsyncWaitHandle.Close() }
        $client.Close()
    }
}

function Read-AVApps {
    $paths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    $found = @()

    foreach ($x in @(Get-ItemProperty -Path $paths -ErrorAction SilentlyContinue)) {
        if (-not $x.DisplayName -or -not $x.UninstallString) { continue }

        $match = $false

        if ($x.DisplayName -match '^360 (Total Security|安全卫士)' -and
            $x.Publisher -match '360|QIHU') {
            $match = $true
        }
        if ($x.DisplayName -match '^Avast (Free Antivirus|Antivirus|Premium Security|One)' -and
            $x.Publisher -match 'Avast|Gen Digital') {
            $match = $true
        }
        if ($x.DisplayName -match '^AVG (AntiVirus|Internet Security)' -and
            $x.Publisher -match 'AVG|Avast|Gen Digital') {
            $match = $true
        }

        if ($match) {
            $found += [pscustomobject]@{
                Name = $x.DisplayName
                Version = $x.DisplayVersion
                Publisher = $x.Publisher
                Command = $x.UninstallString
                Key = $x.PSPath
            }
        }
    }

    return @($found | Sort-Object Key -Unique)
}

function Remove-ExactCurrentPackage {
    param([string]$Name, [string]$Label)

    $packages = @(Get-AppxPackage -Name $Name)

    if ($packages.Count -eq 0) {
        Row $Label 'Không có trên máy' 'Không có gói trong tài khoản đang chạy tool.'
        return
    }

    foreach ($pkg in $packages) {
        Remove-AppxPackage -Package $pkg.PackageFullName -ErrorAction Stop
    }

    if (@(Get-AppxPackage -Name $Name).Count -gt 0) {
        throw ('Gói vẫn còn sau thao tác: {0}' -f $Name)
    }

    Row $Label 'Thành công' (
        'Đã gỡ và kiểm tra lại trong tài khoản đang chạy tool. Không xử lý tài khoản khác hoặc provisioned package.'
    )
}

try {
    switch ($Task) {
        'network' {
            $adapters = @(Read-Adapters)
            Emit 'network' $adapters

            $wifi = (& netsh.exe wlan show interfaces 2>&1 | Out-String)
            Emit 'wifi' $wifi

            Row 'Đọc cấu hình mạng' 'Thành công' (
                'Đã đọc {0} card. Có IP không đồng nghĩa Internet hoặc CUKCUK hoạt động.' -f $adapters.Count
            )
        }

        'ping' {
            [void](Parse-IPv4 $P.IP)
            $ping = New-Object Net.NetworkInformation.Ping
            $times = @()
            $received = 0

            try {
                for ($i = 1; $i -le 4; $i++) {
                    if ($Cancel.Requested) { break }

                    try {
                        $reply = $ping.Send($P.IP, 1200)
                        if ($reply.Status -eq 'Success') {
                            $received++
                            $times += [double]$reply.RoundtripTime
                            Row ('Ping {0}/4' -f $i) 'Thành công' (
                                '{0}: {1} ms; TTL {2}' -f
                                $P.IP, $reply.RoundtripTime, $reply.Options.Ttl
                            )
                        }
                        else {
                            Row ('Ping {0}/4' -f $i) 'Không phản hồi' (
                                '{0}. Không đủ cơ sở kết luận thiết bị đã tắt.' -f $reply.Status
                            )
                        }
                    }
                    catch {
                        Row ('Ping {0}/4' -f $i) 'Thất bại' $_.Exception.Message
                    }

                    Start-Sleep -Milliseconds 180
                }
            }
            finally {
                $ping.Dispose()
            }

            $summary = 'Nhận {0}/4; mất {1}/4.' -f $received, (4 - $received)
            if ($times.Count -gt 0) {
                $m = $times | Measure-Object -Minimum -Maximum -Average
                $summary += ' Min {0} ms; max {1} ms; TB {2:N1} ms.' -f
                    $m.Minimum, $m.Maximum, $m.Average
            }
            Row 'Tổng kết ping' 'Thông tin' $summary
        }

        'dns' {
            [void](Assert-Adapter $P.Index $P.Guid)
            $snapshot = Save-Network $P.Index
            $servers = @()
            $detail = ''

            if ($P.Mode -eq 'Tự động / DHCP') {
                Set-DnsClientServerAddress -InterfaceIndex $P.Index -ResetServerAddresses
                $read = @(
                    (Get-DnsClientServerAddress -InterfaceIndex $P.Index `
                        -AddressFamily IPv4).ServerAddresses
                )
                Row 'DNS tự động' 'Đã cấu hình' (
                    'Đã yêu cầu lấy DNS tự động. DNS đọc lại: {0}. Backup: {1}' -f
                    ($read -join ', '), $snapshot
                )
                break
            }

            if ($P.Mode -eq 'Nhanh nhất') {
                $routeBefore = @(
                    Get-NetRoute -AddressFamily IPv4 -DestinationPrefix '0.0.0.0/0' |
                    Sort-Object InterfaceIndex, NextHop |
                    Select-Object InterfaceIndex, NextHop, RouteMetric
                ) | ConvertTo-Json -Compress

                $adapterBefore = Get-NetAdapter -InterfaceIndex $P.Index
                $ipBefore = @(
                    Get-NetIPAddress -InterfaceIndex $P.Index -AddressFamily IPv4 |
                    Sort-Object IPAddress |
                    Select-Object IPAddress, PrefixLength
                ) | ConvertTo-Json -Compress

                $providers = @(
                    [pscustomobject]@{
                        Name = 'Cloudflare'
                        Addresses = @('1.1.1.1', '1.0.0.1')
                        Times = New-Object 'System.Collections.Generic.List[double]'
                        Success = 0
                    },
                    [pscustomobject]@{
                        Name = 'Google'
                        Addresses = @('8.8.8.8', '8.8.4.4')
                        Times = New-Object 'System.Collections.Generic.List[double]'
                        Success = 0
                    }
                )

                $domains = @(
                    'www.microsoft.com',
                    'www.cloudflare.com',
                    'www.google.com'
                )

                for ($round = 0; $round -lt 3; $round++) {
                    foreach ($provider in $providers) {
                        foreach ($server in $provider.Addresses) {
                            if ($Cancel.Requested) {
                                throw 'Đã hủy đo. Giữ nguyên DNS.'
                            }

                            $sw = [Diagnostics.Stopwatch]::StartNew()
                            try {
                                $answer = @(
                                    Resolve-DnsName -Name $domains[$round] `
                                        -Server $server -Type A -DnsOnly `
                                        -NoHostsFile -QuickTimeout -ErrorAction Stop
                                )
                                $sw.Stop()

                                if (@($answer | Where-Object { $_.IPAddress }).Count -gt 0) {
                                    $provider.Success++
                                    $provider.Times.Add($sw.Elapsed.TotalMilliseconds)
                                }
                            }
                            catch {
                                $sw.Stop()
                            }
                        }
                    }
                    Emit 'log' ('Đo DNS: xong vòng {0}/3.' -f ($round + 1))
                }

                $routeAfter = @(
                    Get-NetRoute -AddressFamily IPv4 -DestinationPrefix '0.0.0.0/0' |
                    Sort-Object InterfaceIndex, NextHop |
                    Select-Object InterfaceIndex, NextHop, RouteMetric
                ) | ConvertTo-Json -Compress

                $ipAfter = @(
                    Get-NetIPAddress -InterfaceIndex $P.Index -AddressFamily IPv4 |
                    Sort-Object IPAddress |
                    Select-Object IPAddress, PrefixLength
                ) | ConvertTo-Json -Compress

                $adapterAfter = Assert-Adapter $P.Index $P.Guid

                if ($routeBefore -ne $routeAfter -or
                    $ipBefore -ne $ipAfter -or
                    $adapterAfter.Status -ne 'Up' -or
                    $adapterBefore.Status -ne 'Up') {
                    throw 'Đường mạng hoặc card thay đổi trong lúc đo. Giữ nguyên DNS.'
                }

                $ranking = @(
                    foreach ($provider in $providers) {
                        $avg = [double]::PositiveInfinity
                        if ($provider.Times.Count -gt 0) {
                            $avg = ($provider.Times | Measure-Object -Average).Average
                        }
                        [pscustomobject]@{
                            Name = $provider.Name
                            Addresses = $provider.Addresses
                            Success = $provider.Success
                            Average = $avg
                        }
                    }
                )

                foreach ($r in $ranking) {
                    Row ('Đo DNS {0}' -f $r.Name) 'Thông tin' (
                        '{0}/6 mẫu thành công; trung bình {1:N1} ms.' -f
                        $r.Success, $r.Average
                    )
                }

                $eligible = @(
                    $ranking |
                    Where-Object { $_.Success -ge 5 } |
                    Sort-Object @{Expression = 'Success'; Descending = $true},
                        @{Expression = 'Average'; Descending = $false}
                )

                if ($eligible.Count -eq 0) {
                    throw 'Không đủ mẫu đáng tin cậy. DNS cũ được giữ nguyên.'
                }

                $winner = $eligible[0]
                $servers = @($winner.Addresses)
                $detail = '{0}; {1}/6 mẫu; trung bình {2:N1} ms. ' -f
                    $winner.Name, $winner.Success, $winner.Average
            }
            else {
                $servers = @($P.Primary, $P.Secondary) |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

                if ($servers.Count -eq 0) { throw 'Chưa nhập DNS.' }
                foreach ($server in $servers) {
                    [void](Parse-IPv4 $server)
                }
            }

            [void](Assert-Adapter $P.Index $P.Guid)
            Set-DnsClientServerAddress -InterfaceIndex $P.Index -ServerAddresses $servers

            $verify = @(
                (Get-DnsClientServerAddress -InterfaceIndex $P.Index `
                    -AddressFamily IPv4).ServerAddresses
            )

            if (($verify -join ',') -ne ($servers -join ',')) {
                throw 'DNS đọc lại không khớp. Kiểm tra cấu hình và bản sao lưu.'
            }

            Row 'Áp dụng DNS' 'Thành công' (
                '{0}DNS: {1}. Không phải phép đo tốc độ tải Internet.' -f
                $detail, ($verify -join ', ')
            )
        }

        'static' {
            $a = Assert-Adapter $P.Index $P.Guid

            if (-not $a.HardwareInterface) {
                throw 'Không tự đặt IP trên VPN/card ảo. Hãy dùng cấu hình Windows.'
            }

            Validate-Static $P.IP ([int]$P.Prefix) $P.Gateway

            $existing = @(
                Get-NetIPAddress -InterfaceIndex $P.Index -AddressFamily IPv4 |
                Where-Object { $_.IPAddress -ne '127.0.0.1' }
            )

            if ($existing.Count -ne 1) {
                throw 'Card không có đúng một IPv4. Không tự sửa hoặc xóa IP phụ. Hãy dùng cấu hình Windows.'
            }

            $old = $existing[0]
            $snapshot = Save-Network $P.Index

            if ($P.IP -ne $old.IPAddress) {
                throw 'Bản này chỉ chuyển IP hiện tại thành tĩnh. Đổi sang IP khác cần kiểm tra trùng IP và thao tác trong Windows.'
            }

            $routes = @(
                Get-NetRoute -InterfaceIndex $P.Index `
                    -AddressFamily IPv4 -DestinationPrefix '0.0.0.0/0'
            )

            if ($routes.Count -gt 1) {
                throw 'Card có nhiều default route. Không tự sửa. Bản sao lưu đã được tạo.'
            }

            Set-NetIPInterface -InterfaceIndex $P.Index `
                -AddressFamily IPv4 -Dhcp Disabled

            $same = @(
                Get-NetIPAddress -InterfaceIndex $P.Index `
                    -AddressFamily IPv4 -IPAddress $P.IP `
                    -ErrorAction SilentlyContinue
            )

            if ($same.Count -gt 0 -and $same[0].PrefixOrigin -eq 'Manual') {
                Set-NetIPAddress -InterfaceIndex $P.Index -IPAddress $P.IP `
                    -PrefixLength ([int]$P.Prefix)
            }
            else {
                if ($same.Count -gt 0) {
                    Remove-NetIPAddress -InterfaceIndex $P.Index `
                        -IPAddress $P.IP -Confirm:$false
                }
                [void](New-NetIPAddress -InterfaceIndex $P.Index `
                    -IPAddress $P.IP -PrefixLength ([int]$P.Prefix))
            }

            foreach ($route in @(
                Get-NetRoute -InterfaceIndex $P.Index `
                    -AddressFamily IPv4 -DestinationPrefix '0.0.0.0/0' `
                    -ErrorAction SilentlyContinue
            )) {
                if ($route.NextHop -ne $P.Gateway) {
                    $route | Remove-NetRoute -Confirm:$false
                }
            }

            $gatewayRoute = @(
                Get-NetRoute -InterfaceIndex $P.Index -AddressFamily IPv4 `
                    -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue |
                Where-Object { $_.NextHop -eq $P.Gateway }
            )

            if ($gatewayRoute.Count -eq 0) {
                [void](New-NetRoute -InterfaceIndex $P.Index `
                    -DestinationPrefix '0.0.0.0/0' -NextHop $P.Gateway `
                    -RouteMetric 10)
            }

            Start-Sleep -Seconds 2

            $verify = Get-NetIPAddress -InterfaceIndex $P.Index `
                -AddressFamily IPv4 -IPAddress $P.IP
            $verifyIf = Get-NetIPInterface -InterfaceIndex $P.Index -AddressFamily IPv4
            $verifyRoute = @(
                Get-NetRoute -InterfaceIndex $P.Index -AddressFamily IPv4 `
                    -DestinationPrefix '0.0.0.0/0' |
                Where-Object { $_.NextHop -eq $P.Gateway }
            )

            if ($verify.PrefixLength -ne [int]$P.Prefix -or
                $verifyIf.Dhcp -ne 'Disabled' -or
                $verifyRoute.Count -eq 0 -or
                $verify.AddressState -eq 'Duplicate') {
                throw ('Cấu hình đọc lại chưa đạt. Backup: {0}' -f $snapshot)
            }

            Row 'Giữ IP hiện tại làm tĩnh' 'Đã cấu hình' (
                '{0}/{1}; gateway {2}. Áp dụng cho card, KHÔNG theo SSID. Hãy kiểm tra lại kết nối. Backup: {3}' -f
                $P.IP, $P.Prefix, $P.Gateway, $snapshot
            )
        }

        'dhcp' {
            [void](Assert-Adapter $P.Index $P.Guid)

            $ips = @(
                Get-NetIPAddress -InterfaceIndex $P.Index -AddressFamily IPv4 |
                Where-Object { $_.IPAddress -ne '127.0.0.1' }
            )

            if ($ips.Count -gt 1) {
                throw 'Card có nhiều IPv4. Không tự xóa IP phụ. Hãy dùng cấu hình Windows.'
            }

            $snapshot = Save-Network $P.Index

            foreach ($ip in $ips) {
                if ($ip.PrefixOrigin -eq 'Manual') {
                    $ip | Remove-NetIPAddress -Confirm:$false
                }
            }

            foreach ($route in @(
                Get-NetRoute -InterfaceIndex $P.Index -AddressFamily IPv4 `
                    -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue |
                Where-Object { $_.Protocol -eq 'NetMgmt' }
            )) {
                $route | Remove-NetRoute -Confirm:$false
            }

            Set-NetIPInterface -InterfaceIndex $P.Index -AddressFamily IPv4 -Dhcp Enabled
            Set-DnsClientServerAddress -InterfaceIndex $P.Index -ResetServerAddresses

            $verify = Get-NetIPInterface -InterfaceIndex $P.Index -AddressFamily IPv4
            if ($verify.Dhcp -ne 'Enabled') {
                throw 'Chưa xác nhận được DHCP đã bật.'
            }

            Row 'IP/DNS tự động' 'Đã cấu hình' (
                'DHCP đã bật. Có thể cần kết nối lại để nhận lease. Backup: {0}' -f $snapshot
            )
        }

        'wifi5' {
            $a = Assert-Adapter $P.Index $P.Guid
            if ([int]$a.NdisPhysicalMedium -notin @(1, 9)) {
                throw 'Card đang chọn không được xác định là Wi-Fi.'
            }

            $properties = @(
                Get-NetAdapterAdvancedProperty -Name $a.Name -AllProperties |
                Where-Object {
                    $_.DisplayName -match 'Preferred Band|Ưu tiên.*băng|Băng.*ưu tiên'
                }
            )

            if ($properties.Count -ne 1) {
                Row 'Ưu tiên Wi-Fi 5 GHz' 'Không hỗ trợ' (
                    'Không xác định duy nhất thuộc tính ưu tiên băng tần do driver cung cấp.'
                )
                break
            }

            $property = $properties[0]
            $values = @($property.ValidDisplayValues)
            $candidates = @(
                $values |
                Where-Object {
                    $_ -match '5\s*GHz' -and $_ -notmatch '6\s*GHz|2[.,]4'
                }
            )

            if ($candidates.Count -ne 1) {
                Row 'Ưu tiên Wi-Fi 5 GHz' 'Không hỗ trợ' (
                    'Driver không công bố duy nhất giá trị ưu tiên 5 GHz có thể xác minh.'
                )
                break
            }

            [void](Save-Snapshot 'wifi' ([pscustomobject]@{
                InterfaceGuid = [string]$a.InterfaceGuid
                Name = $a.Name
                RegistryKeyword = $property.RegistryKeyword
                RegistryValue = @($property.RegistryValue)
            }))

            Set-NetAdapterAdvancedProperty -Name $a.Name `
                -RegistryKeyword $property.RegistryKeyword `
                -DisplayValue $candidates[0] -NoRestart

            $read = Get-NetAdapterAdvancedProperty -Name $a.Name `
                -RegistryKeyword $property.RegistryKeyword

            if ($read.DisplayValue -ne $candidates[0]) {
                throw 'Giá trị đọc lại không khớp.'
            }

            Row 'Ưu tiên Wi-Fi 5 GHz' 'Đã cấu hình, cần khởi động lại' (
                'Không chủ động ngắt card. Driver có thể cần kết nối lại. Không đảm bảo mạng đang dùng là 5 GHz.'
            )
        }

        'wifiRestore' {
            $a = Assert-Adapter $P.Index $P.Guid
            $records = @(
                Get-ChildItem -LiteralPath $P.Source -Filter 'wifi-*.clixml' |
                Sort-Object Name -Descending
            )

            $restore = $null
            foreach ($file in $records) {
                $r = Import-Clixml -LiteralPath $file.FullName
                if ($r.InterfaceGuid -eq [string]$a.InterfaceGuid) {
                    $restore = $r
                    break
                }
            }

            if ($null -eq $restore) {
                throw 'Không có bản sao lưu Wi-Fi khớp card trong thư mục đã chọn.'
            }

            Set-NetAdapterAdvancedProperty -Name $a.Name `
                -RegistryKeyword $restore.RegistryKeyword `
                -RegistryValue $restore.RegistryValue -NoRestart

            $check = Get-NetAdapterAdvancedProperty -Name $a.Name `
                -RegistryKeyword $restore.RegistryKeyword

            if ((@($check.RegistryValue) -join ',') -ne
                (@($restore.RegistryValue) -join ',')) {
                throw 'Giá trị Wi-Fi khôi phục đọc lại không khớp.'
            }

            Row 'Khôi phục ưu tiên Wi-Fi' 'Đã cấu hình, cần khởi động lại' (
                'Đã đọc lại giá trị cũ. Không chủ động ngắt card.'
            )
        }

        'firewall' {
            $before = @(Get-NetFirewallProfile |
                Select-Object Name, Enabled)
            [void](Save-Snapshot 'firewall' $before)

            $enabled = [bool]$P.Enabled
            Set-NetFirewallProfile -Profile Domain, Private, Public -Enabled $enabled

            $after = @(Get-NetFirewallProfile -PolicyStore ActiveStore)
            foreach ($profile in $after) {
                $ok = ([string]$profile.Enabled -eq [string]$enabled)
                if ($ok) {
                    Row ('Firewall {0}' -f $profile.Name) 'Thành công' (
                        'Đọc lại ActiveStore: {0}' -f $profile.Enabled
                    )
                }
                else {
                    Row ('Firewall {0}' -f $profile.Name) 'Thất bại' (
                        'ActiveStore không khớp yêu cầu; có thể bị chính sách quản lý ghi đè.'
                    )
                }
            }
        }

        'avscan' {
            $apps = @(Read-AVApps)
            Emit 'apps' $apps

            if ($apps.Count -eq 0) {
                Row 'Quét antivirus' 'Không có trên máy' (
                    'Không phát hiện ứng dụng khớp danh sách nhận diện giới hạn của bản này.'
                )
            }
            else {
                foreach ($app in $apps) {
                    Row $app.Name 'Đã phát hiện' (
                        'Phiên bản {0}; hãng {1}. Chưa gỡ.' -f $app.Version, $app.Publisher
                    )
                }
            }

            try {
                $mp = Get-MpComputerStatus
                Row 'Microsoft Defender' 'Thông tin' (
                    'AntivirusEnabled={0}; RealTimeProtectionEnabled={1}' -f
                    $mp.AntivirusEnabled, $mp.RealTimeProtectionEnabled
                )
            }
            catch {
                Row 'Microsoft Defender' 'Chưa xác định' $_.Exception.Message
            }
        }

        'avuninstall' {
            $app = @(
                Read-AVApps |
                Where-Object { $_.Key -eq $P.Key }
            )

            if ($app.Count -ne 1) {
                throw 'Ứng dụng đã thay đổi. Hãy quét lại.'
            }

            $command = [Environment]::ExpandEnvironmentVariables($app[0].Command)
            $exe = ''
            $arguments = ''

            if ($command -match '^\s*"([^"]+)"\s*(.*)$') {
                $exe = $matches[1]
                $arguments = $matches[2]
            }
            elseif ($command -match '^\s*(.+?\.exe)\s*(.*)$') {
                $exe = $matches[1]
                $arguments = $matches[2]
            }
            else {
                throw 'Không phân tích an toàn được lệnh gỡ. Hãy mở Installed apps.'
            }

            if ([IO.Path]::GetFileName($exe) -match '^msiexec\.exe$') {
                $exe = Join-Path $env:WINDIR 'System32\msiexec.exe'
                $arguments = $arguments -replace '(?i)/I(?=\s*\{)', '/X'
            }
            elseif (-not [IO.Path]::IsPathRooted($exe)) {
                throw 'Trình gỡ không có đường dẫn tuyệt đối. Hãy gỡ qua Installed apps.'
            }

            if (-not (Test-Path -LiteralPath $exe -PathType Leaf)) {
                throw ('Không tìm thấy trình gỡ: {0}' -f $exe)
            }

            if ([string]::IsNullOrWhiteSpace($arguments)) {
                [void](Start-Process -FilePath $exe -PassThru)
            }
            else {
                [void](Start-Process -FilePath $exe -ArgumentList $arguments -PassThru)
            }

            Row $app[0].Name 'Cần người dùng hoàn tất' (
                'Đã mở trình gỡ đăng ký với Windows. Chưa xác nhận đã gỡ. Hoàn tất, khởi động lại nếu được yêu cầu rồi quét lại.'
            )
        }

        'setup' {
            foreach ($a in @(Read-Adapters | Where-Object { $_.Status -eq 'Up' })) {
                Row ('Mạng: {0}' -f $a.Name) 'Thông tin' (
                    '{0}; IP {1}; gateway {2}; MAC {3}' -f
                    $a.Type,
                    (($a.IPs | ForEach-Object {
                        '{0}/{1}' -f $_.IPAddress, $_.PrefixLength
                    }) -join ', '),
                    $a.Gateway,
                    $a.MAC
                )
            }

            Step 'RAM' {
                $os = Get-CimInstance Win32_OperatingSystem
                Row 'RAM' 'Thông tin' (
                    'Tổng {0:N1} GB; còn khả dụng {1:N1} GB.' -f
                    ($os.TotalVisibleMemorySize / 1MB),
                    ($os.FreePhysicalMemory / 1MB)
                )
            }

            Step 'Ổ hệ thống' {
                $disk = Get-CimInstance Win32_LogicalDisk `
                    -Filter ("DeviceID='{0}'" -f $env:SystemDrive)
                Row 'Ổ hệ thống' 'Thông tin' (
                    '{0}; tổng {1:N1} GB; trống {2:N1} GB.' -f
                    $disk.DeviceID, ($disk.Size / 1GB), ($disk.FreeSpace / 1GB)
                )
            }

            Step 'Print Spooler' {
                $spooler = Get-CimInstance Win32_Service -Filter "Name='Spooler'"
                Row 'Print Spooler' 'Thông tin' (
                    'Trạng thái {0}; khởi động {1}.' -f
                    $spooler.State, $spooler.StartMode
                )
            }

            Step 'Máy in Windows' {
                $printers = @(Get-Printer)
                if ($printers.Count -eq 0) {
                    Row 'Máy in Windows' 'Không có trên máy' 'Không có máy in trong phạm vi đọc được.'
                }
                foreach ($printer in $printers) {
                    Row $printer.Name 'Thông tin' (
                        'Driver {0}; cổng {1}; trạng thái {2}.' -f
                        $printer.DriverName, $printer.PortName, $printer.PrinterStatus
                    )
                }
            }

            Step 'UltraViewer' {
                $services = @(Get-CimInstance Win32_Service |
                    Where-Object {
                        $_.Name -match 'UltraViewer' -or
                        $_.DisplayName -match 'UltraViewer'
                    })
                $processes = @(Get-Process -ErrorAction SilentlyContinue |
                    Where-Object { $_.ProcessName -match '^UltraViewer' })

                if ($services.Count -eq 0 -and $processes.Count -eq 0) {
                    Row 'UltraViewer' 'Chưa xác định' (
                        'Không thấy tiến trình/dịch vụ đang đọc được. Không khẳng định chưa cài.'
                    )
                }
                else {
                    foreach ($service in $services) {
                        Row $service.DisplayName 'Thông tin' (
                            '{0}; khởi động {1}.' -f $service.State, $service.StartMode
                        )
                    }
                    Row 'Tiến trình UltraViewer' 'Thông tin' (
                        '{0} tiến trình. Không kiểm chứng được phiên hỗ trợ từ xa.' -f $processes.Count
                    )
                }
            }
        }

        'printerList' {
            $ports = @(Get-PrinterPort)
            foreach ($printer in @(Get-Printer)) {
                $port = @($ports | Where-Object { $_.Name -eq $printer.PortName })
                $ip = ''
                if ($port.Count -gt 0) { $ip = $port[0].PrinterHostAddress }

                Emit 'printer' ([pscustomobject][ordered]@{
                    'Nguồn' = 'Windows'
                    'Tên / Model' = $printer.Name
                    'IP' = $ip
                    'MAC' = ''
                    'Cổng' = $printer.PortName
                    'Trạng thái' = [string]$printer.PrinterStatus
                    'Máy in Windows' = $printer.Name
                })
            }

            Row 'Máy in Windows' 'Thành công' (
                'Tên máy in lấy từ Windows, không phải xác minh model từ firmware.'
            )
        }

        'scan' {
            $targets = New-Object 'System.Collections.Generic.HashSet[string]'

            foreach ($port in @(Get-PrinterPort -ErrorAction SilentlyContinue)) {
                if ($port.PrinterHostAddress) {
                    try {
                        [void](Parse-IPv4 $port.PrinterHostAddress)
                        [void]$targets.Add($port.PrinterHostAddress)
                    }
                    catch {}
                }
            }

            foreach ($neighbor in @(
                Get-NetNeighbor -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.State -notin @('Unreachable', 'Incomplete') -and
                    $_.IPAddress -notmatch '^(224\.|239\.|255\.)'
                }
            )) {
                [void]$targets.Add($neighbor.IPAddress)
            }

            if (-not [string]::IsNullOrWhiteSpace($P.Range)) {
                $parts = $P.Range.Trim().Split('/')
                if ($parts.Count -eq 1) {
                    [void](Parse-IPv4 $parts[0])
                    [void]$targets.Add($parts[0])
                }
                elseif ($parts.Count -eq 2) {
                    $prefix = 0
                    if (-not [int]::TryParse($parts[1], [ref]$prefix) -or
                        $prefix -lt 24 -or $prefix -gt 30) {
                        throw 'Quét CIDR chỉ hỗ trợ /24 đến /30 để giới hạn phạm vi.'
                    }

                    $n = IPv4-Number $parts[0]
                    $size = [uint64][Math]::Pow(2, 32 - $prefix)
                    $net = [uint64]([Math]::Floor($n / $size) * $size)

                    for ($offset = 1; $offset -lt ($size - 1); $offset++) {
                        $v = $net + $offset
                        $text = '{0}.{1}.{2}.{3}' -f
                            [int][Math]::Floor($v / 16777216),
                            [int]([Math]::Floor($v / 65536) % 256),
                            [int]([Math]::Floor($v / 256) % 256),
                            [int]($v % 256)
                        [void]$targets.Add($text)
                    }
                }
                else {
                    throw 'Nhập một IPv4 hoặc CIDR như 192.168.1.0/24.'
                }
            }

            if ($targets.Count -gt 512) {
                throw 'Phạm vi vượt 512 IP. Hãy thu hẹp hoặc kiểm tra từng mạng.'
            }

            $done = 0
            foreach ($target in @($targets)) {
                if ($Cancel.Requested) { break }
                $done++

                foreach ($portNumber in @(9100, 515, 631)) {
                    if ($Cancel.Requested) { break }

                    if (Test-Tcp $target $portNumber 250) {
                        $neighbor = @(
                            Get-NetNeighbor -IPAddress $target `
                                -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                            Select-Object -First 1
                        )
                        $mac = ''
                        if ($neighbor.Count -gt 0) {
                            $mac = $neighbor[0].LinkLayerAddress
                        }

                        Emit 'printer' ([pscustomobject][ordered]@{
                            'Nguồn' = 'TCP / Windows / ARP'
                            'Tên / Model' = 'Chưa xác định'
                            'IP' = $target
                            'MAC' = $mac
                            'Cổng' = [string]$portNumber
                            'Trạng thái' = 'Cổng mở; nghi là máy in'
                            'Máy in Windows' = ''
                        })
                    }
                }

                Emit 'progress' ([pscustomobject]@{
                    Done = $done
                    Total = $targets.Count
                })
            }

            if ($Cancel.Requested) {
                Row 'Dò máy in' 'Đã dừng' 'Giữ các kết quả đã tìm được.'
            }
            else {
                Row 'Dò máy in' 'Hoàn thành' (
                    'Đã kiểm tra {0} IP. Cổng mở không xác minh model. Không có route hoặc không nằm trong phạm vi có thể không được phát hiện.' -f $done
                )
            }
        }

        'port' {
            [void](Parse-IPv4 $P.IP)
            $port = [int]$P.Port
            if ($port -lt 1 -or $port -gt 65535) { throw 'Cổng không hợp lệ.' }

            if (Test-Tcp $P.IP $port 1500) {
                Row 'Kiểm tra cổng' 'Thành công' (
                    '{0}:{1} chấp nhận TCP. Chưa xác minh thiết bị/model hoặc khả năng in.' -f $P.IP, $port
                )
            }
            else {
                Row 'Kiểm tra cổng' 'Không phản hồi' (
                    '{0}:{1}. Có thể do route, firewall, thiết bị hoặc cổng khác.' -f $P.IP, $port
                )
            }
        }

        'print' {
            [void](Parse-IPv4 $P.IP)
            $port = [int]$P.Port
            if ($port -lt 1 -or $port -gt 65535) { throw 'Cổng không hợp lệ.' }

            $client = New-Object Net.Sockets.TcpClient
            $ar = $null
            try {
                $ar = $client.BeginConnect($P.IP, $port, $null, $null)
                if (-not $ar.AsyncWaitHandle.WaitOne(2000)) {
                    throw 'Hết thời gian kết nối.'
                }
                $client.EndConnect($ar)
                $client.SendTimeout = 2500

                $code = 'CC-' + (Get-Date -Format 'HHmmss')
                $text = @"

================================
       CUKCUK - NHAN DIEN
================================
IP TOOL GUI TOI: $($P.IP)
THOI GIAN: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
MA DOI CHIEU: $code

Day khong phai self-test firmware.
Khach bao lai may vua ra giay.
================================




"@
                # Chi gui ASCII va xuong dong.
                # Khong gui lenh cat/buzzer/cau hinh chua xac minh.
                $bytes = [Text.Encoding]::ASCII.GetBytes($text.Replace("`r`n", "`n"))
                $stream = $client.GetStream()
                $stream.Write($bytes, 0, $bytes.Length)
                $stream.Flush()

                Row 'In phiếu nhận diện' 'Cần người dùng xác nhận' (
                    'Đã gửi {0} byte tới {1}:{2}; mã {3}. Chưa xác nhận giấy đã in.' -f
                    $bytes.Length, $P.IP, $port, $code
                )
            }
            finally {
                if ($null -ne $ar) { $ar.AsyncWaitHandle.Close() }
                $client.Close()
            }
        }

        'restoreRegistry' {
            $files = @(
                Get-ChildItem -LiteralPath $P.Source -Filter 'registry-*.clixml' |
                Sort-Object Name -Descending
            )

            if ($files.Count -eq 0) {
                throw 'Thư mục không có bản sao lưu Registry.'
            }

            foreach ($file in $files) {
                $r = Import-Clixml -LiteralPath $file.FullName
                $base = Open-RegistryBase $r.Hive
                $key = $null

                try {
                    $key = $base.CreateSubKey($r.Path)
                    if ($r.Exists) {
                        $kind = [Microsoft.Win32.RegistryValueKind]::$($r.Kind)
                        $key.SetValue($r.Name, $r.Value, $kind)

                        $actual = $key.GetValue(
                            $r.Name, $null,
                            [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames
                        )

                        if (($actual | ConvertTo-Json -Compress) -ne
                            ($r.Value | ConvertTo-Json -Compress)) {
                            throw 'Giá trị khôi phục đọc lại không khớp.'
                        }
                    }
                    else {
                        $key.DeleteValue($r.Name, $false)
                        if (@($key.GetValueNames()) -contains $r.Name) {
                            throw 'Giá trị vẫn tồn tại sau khôi phục.'
                        }
                    }

                    Row ('Khôi phục {0}' -f $r.Name) 'Đã cấu hình, cần khởi động lại' (
                        'Đã đọc lại giá trị. Không xóa key rỗng để tránh ảnh hưởng giá trị khác.'
                    )
                }
                catch {
                    Row ('Khôi phục {0}' -f $r.Name) 'Thất bại' $_.Exception.Message
                }
                finally {
                    if ($null -ne $key) { $key.Dispose() }
                    $base.Dispose()
                }
            }
        }

        'optimize' {
            $os = Get-CimInstance Win32_OperatingSystem
            $build = [int]$os.BuildNumber
            $isClient = ([int]$os.ProductType -eq 1)

            if (-not $isClient -or $build -lt 10240) {
                throw 'Tối ưu này chỉ dành cho Windows 10/11 client.'
            }

            # Tao diem khoi phuc truoc cac thay doi.
            Step 'Tạo điểm khôi phục' {
                $description = 'CUKCUK ' + (Get-Date -Format 'yyyyMMdd-HHmmss')
                Enable-ComputerRestore -Drive ($env:SystemDrive + '\')
                Checkpoint-Computer -Description $description `
                    -RestorePointType MODIFY_SETTINGS -ErrorAction Stop

                $points = @(
                    Get-ComputerRestorePoint |
                    Where-Object { $_.Description -eq $description }
                )

                if ($points.Count -eq 0) {
                    throw 'Không xác nhận được điểm khôi phục mới. Có thể vướng giới hạn tần suất hoặc System Protection.'
                }
                Row 'Tạo điểm khôi phục' 'Thành công' $description
            }

            Step 'Tắt lịch sử hoạt động' {
                foreach ($name in @(
                    'EnableActivityFeed', 'PublishUserActivities', 'UploadUserActivities'
                )) {
                    Set-BackedRegistry HKLM `
                        'SOFTWARE\Policies\Microsoft\Windows\System' $name 0
                }
                Policy-Result 'Tắt lịch sử hoạt động'
            }

            Step 'Giảm nội dung và ứng dụng gợi ý' {
                Set-BackedRegistry HKLM `
                    'SOFTWARE\Policies\Microsoft\Windows\CloudContent' `
                    'DisableWindowsConsumerFeatures' 1
                Policy-Result 'Giảm nội dung và ứng dụng gợi ý'
            }

            Step 'Tắt Delivery Optimization ngang hàng' {
                Set-BackedRegistry HKLM `
                    'SOFTWARE\Policies\Microsoft\Windows\DeliveryOptimization' `
                    'DODownloadMode' 0
                Policy-Result 'Tắt Delivery Optimization ngang hàng'
            }

            Step 'End task trên taskbar' {
                if ($build -lt 22631) {
                    Row 'End task trên taskbar' 'Không hỗ trợ' (
                        'Bản này chỉ áp dụng từ Windows 11 build 22631. Không báo hỗ trợ trên Windows 10.'
                    )
                }
                else {
                    Set-BackedRegistry HKCU `
                        'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\TaskbarDeveloperSettings' `
                        'TaskbarEndTask' 1
                    Policy-Result 'End task trên taskbar'
                }
            }

            Step 'Tắt nhận diện loại thư mục tự động' {
                Set-BackedRegistry HKCU `
                    'Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\Bags\AllFolders\Shell' `
                    'FolderType' 'NotSpecified' String
                Policy-Result 'Tắt nhận diện loại thư mục tự động'
            }

            Step 'Tắt theo dõi vị trí' {
                Set-BackedRegistry HKLM `
                    'SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors' `
                    'DisableLocation' 1
                Policy-Result 'Tắt theo dõi vị trí'
            }

            Row 'Kết quả Microsoft Store trong tìm kiếm' 'Không hỗ trợ' (
                'Chưa xác minh chính xác mục WinUtil mong muốn. Không thay bằng một policy tìm kiếm khác.'
            )

            Step 'Chuyển dịch vụ đã chọn sang Manual' {
                $names = @('MapsBroker', 'WerSvc', 'SysMain', 'TrkWks')
                $snapshot = @(
                    Get-CimInstance Win32_Service |
                    Where-Object { $_.Name -in $names } |
                    Select-Object Name, StartMode, State
                )
                [void](Save-Snapshot 'services' $snapshot)

                foreach ($name in $names) {
                    try {
                        $service = Get-Service -Name $name -ErrorAction SilentlyContinue
                        if ($null -eq $service) {
                            Row ('Dịch vụ {0}' -f $name) 'Không có trên máy' ''
                            continue
                        }

                        Set-Service -Name $name -StartupType Manual
                        $verify = Get-CimInstance Win32_Service `
                            -Filter ("Name='{0}'" -f $name)

                        if ($verify.StartMode -ne 'Manual') {
                            throw 'Kiểu khởi động đọc lại không phải Manual.'
                        }

                        Row ('Dịch vụ {0}' -f $name) 'Thành công' (
                            'Đã chuyển Manual. Không chủ động dừng dịch vụ. Không thay WSearch.'
                        )
                    }
                    catch {
                        Row ('Dịch vụ {0}' -f $name) 'Thất bại' $_.Exception.Message
                    }
                }
            }

            Row 'Start Menu bố cục trước đây' 'Không hỗ trợ' (
                'Không ép Registry không tương thích. Chưa có cách đáng tin cậy cho mọi build.'
            )

            Step 'Giảm Telemetry' {
                # Gia tri 1: diagnostic data required, khong hua tat sach.
                Set-BackedRegistry HKLM `
                    'SOFTWARE\Policies\Microsoft\Windows\DataCollection' `
                    'AllowTelemetry' 1
                Policy-Result 'Giảm Telemetry về mức dữ liệu bắt buộc'
            }

            Step 'Dọn TEMP cũ hơn 7 ngày' {
                $roots = @(
                    [IO.Path]::GetFullPath($env:TEMP),
                    (Join-Path $env:WINDIR 'Temp')
                ) | Select-Object -Unique

                $cutoff = (Get-Date).AddDays(-7)
                $deleted = 0
                $skipped = 0
                $enumerationErrors = @()

                foreach ($root in $roots) {
                    if (-not (Test-Path -LiteralPath $root)) { continue }

                    $rootItem = Get-Item -LiteralPath $root
                    if ($rootItem.Attributes -band [IO.FileAttributes]::ReparsePoint) {
                        $skipped++
                        continue
                    }

                    # Tu duyet thu muc de khong di theo junction/symlink.
                    $stack = New-Object 'System.Collections.Generic.Stack[string]'
                    $stack.Push($root)

                    while ($stack.Count -gt 0) {
                        if ($Cancel.Requested) { break }
                        $directory = $stack.Pop()

                        $children = @(
                            Get-ChildItem -LiteralPath $directory -Force `
                                -ErrorAction SilentlyContinue `
                                -ErrorVariable +enumerationErrors
                        )

                        foreach ($item in $children) {
                            if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) {
                                $skipped++
                                continue
                            }

                            if ($item.PSIsContainer) {
                                $stack.Push($item.FullName)
                                continue
                            }

                            if ($item.LastWriteTime -ge $cutoff) { continue }

                            try {
                                Remove-Item -LiteralPath $item.FullName `
                                    -Force -ErrorAction Stop
                                $deleted++
                            }
                            catch {
                                $skipped++
                            }
                        }
                    }
                }

                Row 'Dọn TEMP cũ hơn 7 ngày' 'Hoàn thành' (
                    'Đã xóa {0} tệp; bỏ qua/lỗi {1}; lỗi đọc thư mục {2}. Không xóa thư mục dữ liệu người dùng.' -f
                    $deleted, $skipped, $enumerationErrors.Count
                )
            }

            Step 'Tắt Widgets' {
                if ($build -lt 22000) {
                    Row 'Tắt Widgets' 'Không hỗ trợ' 'Không áp dụng Widgets Windows 11 trên Windows 10.'
                }
                else {
                    Set-BackedRegistry HKLM `
                        'SOFTWARE\Policies\Microsoft\Dsh' 'AllowNewsAndInterests' 0
                    Policy-Result 'Tắt Widgets'

                    $widgets = @(Get-AppxPackage -Name 'MicrosoftWindows.Client.WebExperience')
                    Row 'Gói Widgets' 'Thông tin' (
                        'Tìm thấy {0} gói trong tài khoản hiện tại. Bản này tắt bằng policy, không gỡ gói Web Experience.' -f $widgets.Count
                    )
                }
            }

            Step 'Vô hiệu hóa WPBT' {
                Set-BackedRegistry HKLM `
                    'SYSTEM\CurrentControlSet\Control\Session Manager' `
                    'DisableWpbtExecution' 1
                Policy-Result 'Vô hiệu hóa WPBT'
            }

            Step 'Gỡ Microsoft OneDrive' {
                $uninstallEntries = @(
                    Get-ItemProperty -Path @(
                        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
                        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
                        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
                    ) -ErrorAction SilentlyContinue |
                    Where-Object {
                        $_.DisplayName -eq 'Microsoft OneDrive' -and
                        $_.Publisher -match 'Microsoft'
                    }
                )

                $candidates = @(
                    (Join-Path $env:WINDIR 'System32\OneDriveSetup.exe'),
                    (Join-Path $env:WINDIR 'SysWOW64\OneDriveSetup.exe')
                )

                $setup = $null
                foreach ($entry in $uninstallEntries) {
                    $command = [Environment]::ExpandEnvironmentVariables(
                        [string]$entry.UninstallString
                    )
                    if ($command -match '^\s*"([^"]*OneDriveSetup\.exe)"') {
                        if (Test-Path -LiteralPath $matches[1]) {
                            $setup = $matches[1]
                            break
                        }
                    }
                }

                if (-not $setup) {
                    $setup = $candidates |
                        Where-Object { Test-Path -LiteralPath $_ } |
                        Select-Object -First 1
                }

                if (-not $setup) {
                    if ($uninstallEntries.Count -eq 0) {
                        Row 'Gỡ Microsoft OneDrive' 'Không có trên máy' (
                            'Không thấy ứng dụng đăng ký hoặc bộ gỡ trong phạm vi kiểm tra.'
                        )
                    }
                    else {
                        Row 'Gỡ Microsoft OneDrive' 'Cần người dùng xác nhận' (
                            'Có ứng dụng đăng ký nhưng không tìm được bộ gỡ phù hợp. Dùng Installed apps.'
                        )
                    }
                }
                else {
                    $sig = Get-AuthenticodeSignature -LiteralPath $setup
                    if ($sig.Status -ne 'Valid' -or
                        $sig.SignerCertificate.Subject -notmatch 'Microsoft') {
                        throw 'Bộ gỡ không có chữ ký Microsoft hợp lệ.'
                    }

                    $process = Start-Process -FilePath $setup `
                        -ArgumentList '/uninstall' -PassThru
                    $process.WaitForExit()

                    $left = @(
                        Get-ItemProperty -Path @(
                            'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
                            'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
                            'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
                        ) -ErrorAction SilentlyContinue |
                        Where-Object {
                            $_.DisplayName -eq 'Microsoft OneDrive' -and
                            $_.Publisher -match 'Microsoft'
                        }
                    )

                    if ($left.Count -eq 0 -and $process.ExitCode -eq 0) {
                        Row 'Gỡ Microsoft OneDrive' 'Thành công' (
                            'Bộ gỡ kết thúc mã 0; không còn mục đăng ký trong phạm vi kiểm tra. Không xóa dữ liệu OneDrive.'
                        )
                    }
                    else {
                        Row 'Gỡ Microsoft OneDrive' 'Cần kiểm tra lại' (
                            'Mã thoát {0}; còn {1} mục đăng ký. Không khẳng định đã gỡ.' -f
                            $process.ExitCode, $left.Count
                        )
                    }
                }
            }

            Step 'Menu chuột phải kiểu cũ' {
                if ($build -lt 22000) {
                    Row 'Menu chuột phải kiểu cũ' 'Không áp dụng' (
                        'Windows 10 không cần thay đổi này.'
                    )
                }
                else {
                    Set-BackedRegistry HKCU `
                        'Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32' `
                        '' '' String
                    Policy-Result 'Menu chuột phải kiểu cũ'
                }
            }

            Step 'Gỡ ứng dụng Microsoft Copilot' {
                Remove-ExactCurrentPackage 'Microsoft.Copilot' 'Ứng dụng Microsoft Copilot'
            }

            Step 'Tắt Copilot Windows kiểu cũ' {
                if ($build -lt 22000) {
                    Row 'Copilot Windows kiểu cũ' 'Không áp dụng' (
                        'Không áp dụng policy Windows 11 trên máy này.'
                    )
                }
                else {
                    Set-BackedRegistry HKCU `
                        'Software\Policies\Microsoft\Windows\WindowsCopilot' `
                        'TurnOffWindowsCopilot' 1
                    Policy-Result 'Tắt Copilot Windows kiểu cũ'
                }
            }

            Step 'Recall' {
                if ($build -lt 26100) {
                    Row 'Recall' 'Không hỗ trợ' 'Không thuộc build mục tiêu có Recall của bản này.'
                }
                else {
                    $feature = @(
                        Get-WindowsOptionalFeature -Online |
                        Where-Object { $_.FeatureName -eq 'Recall' }
                    )

                    if ($feature.Count -eq 0) {
                        Row 'Recall' 'Không có trên máy' (
                            'Không có optional feature tên Recall trong phạm vi kiểm tra.'
                        )
                    }
                    else {
                        Set-BackedRegistry HKLM `
                            'SOFTWARE\Policies\Microsoft\Windows\WindowsAI' `
                            'DisableAIDataAnalysis' 1

                        if ($feature[0].State -eq 'Enabled') {
                            $result = Disable-WindowsOptionalFeature -Online `
                                -FeatureName Recall -NoRestart

                            $verify = Get-WindowsOptionalFeature -Online -FeatureName Recall
                            if ($verify.State -notin @('Disabled', 'DisablePending')) {
                                throw ('Recall đọc lại: {0}' -f $verify.State)
                            }

                            Row 'Recall' 'Đã cấu hình, cần khởi động lại' (
                                'Đã yêu cầu tắt optional feature; trạng thái {0}. Không tuyên bố gỡ toàn bộ AI.' -f $verify.State
                            )
                        }
                        else {
                            Row 'Recall' 'Đã cấu hình, cần khởi động lại' (
                                'Feature đang ở trạng thái {0}; đã cấu hình không lưu snapshot.' -f $feature[0].State
                            )
                        }
                    }
                }
            }

            Step 'Bảo vệ Print Spooler' {
                $spoolerBefore = Get-CimInstance Win32_Service -Filter "Name='Spooler'"
                [void](Save-Snapshot 'spooler' $spoolerBefore)

                Set-Service -Name Spooler -StartupType Automatic
                $service = Get-Service -Name Spooler
                if ($service.Status -ne 'Running') {
                    Start-Service -Name Spooler
                }

                $verify = Get-CimInstance Win32_Service -Filter "Name='Spooler'"
                if ($verify.StartMode -ne 'Auto' -or $verify.State -ne 'Running') {
                    throw ('Spooler: {0}; {1}' -f $verify.StartMode, $verify.State)
                }

                Row 'Bảo vệ Print Spooler' 'Thành công' (
                    'Đã xác nhận Automatic và Running. Không restart khi đang chạy.'
                )
            }

            Row 'Phạm vi tối ưu' 'Thông tin' (
                'Không thay UltraViewer, Defender, Firewall, Windows Update, BitLocker, IPv6, Hibernate hoặc Edge.'
            )
        }

        default {
            throw ('Tác vụ chưa được định nghĩa: {0}' -f $Task)
        }
    }
}
catch {
    Row $Task 'Thất bại' $_.Exception.Message
}
finally {
    Emit 'finished' $null
}
'@

# ---------------------------------------------------------------------
# 3. GIAO DIEN WPF
# ---------------------------------------------------------------------

[xml]$script:CCXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Tối Ưu CUKCUK"
        Width="1160" Height="800"
        MinWidth="960" MinHeight="690"
        WindowStartupLocation="CenterScreen"
        Background="#0B1120"
        Foreground="#E5EAF4"
        FontFamily="Segoe UI"
        FontSize="13"
        UseLayoutRounding="True">
    <Window.Resources>
        <SolidColorBrush x:Key="Card" Color="#141E32"/>
        <SolidColorBrush x:Key="Muted" Color="#9BAAC2"/>

        <Style TargetType="TextBlock">
            <Setter Property="TextWrapping" Value="Wrap"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
        </Style>

        <Style TargetType="Button">
            <Setter Property="Foreground" Value="#F3F6FF"/>
            <Setter Property="Background" Value="#22324F"/>
            <Setter Property="BorderBrush" Value="#3C5276"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="13,10"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="MinHeight" Value="40"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="HorizontalContentAlignment" Value="Center"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="B"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="9"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter
                                HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}"
                                VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="B" Property="Background" Value="#315581"/>
                                <Setter TargetName="B" Property="BorderBrush" Value="#58C4FF"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="B" Property="Background" Value="#1678B7"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="B" Property="Opacity" Value="0.40"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#0C1425"/>
            <Setter Property="Foreground" Value="#F1F5FC"/>
            <Setter Property="BorderBrush" Value="#35476A"/>
            <Setter Property="Padding" Value="9,7"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="MinHeight" Value="34"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="CaretBrush" Value="White"/>
        </Style>

        <Style TargetType="ComboBox">
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Padding" Value="8"/>
            <Setter Property="MinHeight" Value="34"/>
            <Setter Property="Foreground" Value="#111827"/>
            <Setter Property="Background" Value="#E5EAF4"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>

        <Style TargetType="TabControl">
            <Setter Property="Background" Value="#0B1120"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="8"/>
        </Style>

        <Style TargetType="TabItem">
            <Setter Property="Foreground" Value="#AEBED6"/>
            <Setter Property="Margin" Value="4"/>
            <Setter Property="Padding" Value="15,12"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="TabBorder"
                                CornerRadius="9"
                                Background="#152039"
                                BorderBrush="#2A3A58"
                                BorderThickness="1"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter ContentSource="Header"
                                              VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="TabBorder" Property="Background" Value="#16466A"/>
                                <Setter TargetName="TabBorder" Property="BorderBrush" Value="#3DC7F4"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="TabBorder" Property="BorderBrush" Value="#6B9BD2"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="#10192B"/>
            <Setter Property="Foreground" Value="#E6EDF7"/>
            <Setter Property="RowBackground" Value="#142037"/>
            <Setter Property="AlternatingRowBackground" Value="#192840"/>
            <Setter Property="BorderBrush" Value="#304
