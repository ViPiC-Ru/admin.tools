<#
  .SYNOPSIS
  Добавляет и контролирует временный маршрут в приделах метрики.

  .DESCRIPTION
  Скрипт проверяет список IPv4 и IPv6 маршрутов с заданой метрикой.
  Если необходимый маршрут отсутствует, то добовляет его и удаляет
  устаревшие маршруты.

  .PARAMETER Destination
  Доменное имя или IP адрес хоста или сети до которого нужно добавить маршрут.

  .PARAMETER Gateway
  Доменное имя или IP адрес шлюза через который нужно добавить маршрут.

  .PARAMETER Check
  Доменное имя или IP адрес узла который должен быть доступен.

  .PARAMETER Metric
  Метрика маршрута в приделах который производиться работа с маршрутами.

  .INPUTS
  Вы не можете передавать в скрипт объекты по коневееру.

  .OUTPUTS
  Скрипт не генерит возвращаемые объекты.

  .NOTES
  Версия: 0.2.1
  Автор: @ViPiC
#>

[CmdletBinding()]
Param (
    [Parameter (Mandatory = $true)]
    [string]$Destination,

    [Parameter (Mandatory = $true)]
    [string]$Gateway,

    [string]$Check,

    [Parameter (Mandatory = $true)]
    [ValidateRange(1, 255)]
    [int]$Metric
);

$MaskDelim = "/";
$maxCIDRv4 = 32; $maxCIDRv6 = 128;
$isPASSv4 = $true; $isPASSv6 = $true;
# Получаем информацию об ip адресах шлюза
$GwIPv4s = @(); $GwIPv6s = @();
if ($Gateway -and ($isPASSv4 -or $isPASSv6)) {
    $List = @($Gateway);
    if ([IPAddress]::TryParse($List[0], [ref][IPAddress]::Loopback)) {
        if (($List[0].IndexOf(".") -ge 0) -and $isPASSv4) { $GwIPv4s = @($List[0]); };
        if (($List[0].IndexOf(":") -ge 0) -and $isPASSv6) { $GwIPv6s = @($List[0]); };
    }
    else {
        $Resolve = (Resolve-DnsName -Name $List[0] -ErrorAction "Ignore" | Where-Object "Section" -EQ "Answer");
        if ($Resolve -and $isPASSv4) { foreach ($GwIPv4 in $Resolve.IP4Address) { $GwIPv4s += $GwIPv4; }; };
        if ($Resolve -and $isPASSv6) { foreach ($GwIPv6 in $Resolve.IP6Address) { $GwIPv6s += $GwIPv6; }; };
    };
    $isPASSv4 = ($GwIPv4s.Count -gt 0);
    $isPASSv6 = ($GwIPv6s.Count -gt 0);
};
# Получаем информацию об ip адресах проверяемого узла
$ChkIPv4s = @(); $ChkIPv6s = @();
if ($Check -and ($isPASSv4 -or $isPASSv6)) {
    $List = @($Check);
    if ([IPAddress]::TryParse($List[0], [ref][IPAddress]::Loopback)) {
        if (($List[0].IndexOf(".") -ge 0) -and $isPASSv4) { $ChkIPv4s = @($List[0]); };
        if (($List[0].IndexOf(":") -ge 0) -and $isPASSv6) { $ChkIPv6s = @($List[0]); };
    }
    else {
        $Resolve = (Resolve-DnsName -Name $List[0] -ErrorAction "Ignore" | Where-Object "Section" -EQ "Answer");
        if ($Resolve -and $isPASSv4) { foreach ($ChkIPv4 in $Resolve.IP4Address) { $ChkIPv4s += $ChkIPv4; }; };
        if ($Resolve -and $isPASSv6) { foreach ($ChkIPv6 in $Resolve.IP6Address) { $ChkIPv6s += $ChkIPv6; }; };
    };
    $isPASSv4 = (($ChkIPv4s.Count -gt 0) -and (Test-Connection -ComputerName $ChkIPv4s -Quiet -ErrorAction "Ignore"));
    $isPASSv6 = (($ChkIPv6s.Count -gt 0) -and (Test-Connection -ComputerName $ChkIPv6s -Quiet -ErrorAction "Ignore"));
};
# Получаем информацию об ip адресах хоста назначения
$DstIPv4s = @(); $DstIPv6s = @();
if ($Destination -and ($isPASSv4 -or $isPASSv6)) {
    $List = $Destination.Split($MaskDelim);
    if ([IPAddress]::TryParse($List[0], [ref][IPAddress]::Loopback)) {
        if (($List[0].IndexOf(".") -ge 0) -and $isPASSv4) { $DstIPv4s = @($List[0]); };
        if (($List[0].IndexOf(":") -ge 0) -and $isPASSv6) { $DstIPv6s = @($List[0]); };
    }
    else {
        $Resolve = (Resolve-DnsName -Name $List[0] -ErrorAction "Ignore" | Where-Object "Section" -EQ "Answer");
        if ($Resolve -and $isPASSv4) { foreach ($DstIPv4 in $Resolve.IP4Address) { $DstIPv4s += $DstIPv4; }; };
        if ($Resolve -and $isPASSv6) { foreach ($DstIPv6 in $Resolve.IP6Address) { $DstIPv6s += $DstIPv6; }; };
    };
    $isPASSv4 = ($DstIPv4s.Count -gt 0);
    $isPASSv6 = ($DstIPv6s.Count -gt 0);
};
# Прнобразовываем ip адреса хоста назначения в адреса подсетей
$DstNETv4s = @(); $DstNETv6s = @();
if ($Destination -and ($isPASSv4 -or $isPASSv6)) {
    $List = $Destination.Split($MaskDelim);
    $CIDRv4 = if ($List.Count -eq 2) { [Math]::Max(0, [Math]::Min($List[1] * 1, $maxCIDRv4)) } else { $maxCIDRv4 };
    foreach ($DstIPv4 in $DstIPv4s) {
        $IPv4 = [IPAddress]::Parse($DstIPv4);
        $MASKv4 = [IPAddress]([Math]::Pow(2, $maxCIDRv4) - 1 -bxor [Math]::Pow(2, ($maxCIDRv4 - $CIDRv4)) - 1);
        $NETv4 = [IPAddress]($IPv4.Address -band $MASKv4.Address);
        $DstNETv4 = $NETv4.IPAddressToString + $MaskDelim + $CIDRv4;
        if (-not($DstNETv4 -in $DstNETv4s)) { $DstNETv4s += $DstNETv4; };
    };
    $CIDRv6 = if ($List.Count -eq 2) { [Math]::Max(0, [Math]::Min($List[1] * 1, $maxCIDRv6)) } else { $maxCIDRv6 };
    foreach ($DstIPv6 in $DstIPv6s) {
        $IPv6 = [IPAddress]::Parse($DstIPv6);
        $NETv6 = $IPv6;
        $DstNETv6 = $NETv6.IPAddressToString + $MaskDelim + $CIDRv6;
        if (-not($DstNETv6 -in $DstNETv6s)) { $DstNETv6s += $DstNETv6; };
    };
};
# Определяем интерфейсы для шлюзов
$IfIDv4s = @(); $IfIDv6s = @();
if ($Gateway -and ($isPASSv4 -or $isPASSv6)) {
    foreach ($GwIPv4 in $GwIPv4s) {
        $Find = Find-NetRoute -RemoteIPAddress $GwIPv4 -ErrorAction "Ignore";
        $IfIDv4 = if ($Find) { $Find[0].InterfaceIndex } else { $null };
        $IfIDv4s += $IfIDv4;
    };
    foreach ($GwIPv6 in $GwIPv6s) {
        $Find = Find-NetRoute -RemoteIPAddress $GwIPv6 -ErrorAction "Ignore";
        $IfIDv6 = if ($Find) { $Find[0].InterfaceIndex } else { $null };
        $IfIDv6s += $IfIDv6;
    };
};
# Удаляем устаревшие маршруты
$NetRoutes = (Get-NetRoute -RouteMetric $Metric -PolicyStore "ActiveStore" -ErrorAction "Ignore" | Where-Object "NextHop" -NotIn "::","0.0.0.0");
if (-not($NetRoutes)) { $NetRoutes = @(); };
foreach ($NetRoute in $NetRoutes) {
    $isPASSv4 = $false;
    for ($i = 0; ($i -lt $GwIPv4s.Count) -and -not($isPASSv4); $i++) {
        if (($NetRoute.NextHop -eq $GwIPv4s[$i]) -and ($NetRoute.InterfaceIndex -eq $IfIDv4s[$i])) {
            for ($j = 0; ($j -lt $DstNETv4s.Count) -and -not($isPASSv4); $j++) {
                if ($NetRoute.DestinationPrefix -eq $DstNETv4s[$j]) { $isPASSv4 = $true; };
            };
        };
    };
    $isPASSv6 = $false;
    for ($i = 0; ($i -lt $GwIPv6s.Count) -and -not($isPASSv6); $i++) {
        if (($NetRoute.NextHop -eq $GwIPv6s[$i]) -and ($NetRoute.InterfaceIndex -eq $IfIDv6s[$i])) {
            for ($j = 0; ($j -lt $DstNETv6s.Count) -and -not($isPASSv6); $j++) {
                if ($NetRoute.DestinationPrefix -eq $DstNETv6s[$j]) { $isPASSv6 = $true; };
            };
        };
    };
    if (-not($isPASSv4 -or $isPASSv6)) { Remove-NetRoute -InputObject $NetRoute -Confirm:$false; };
};
# Добавляем новые маршруты
for ($i = 0; $i -lt $GwIPv4s.Count; $i++) {
    for ($j = 0; $j -lt $DstNETv4s.Count; $j++) {
        $isPASSv4 = $IfIDv4s[$i];
        foreach ($NetRoute in $NetRoutes) {
            if (($NetRoute.NextHop -eq $GwIPv4s[$i]) -and ($NetRoute.InterfaceIndex -eq $IfIDv4s[$i])) {
                if ($NetRoute.DestinationPrefix -eq $DstNETv4s[$j]) { $isPASSv4 = $false; };
            };
        };
        if ($isPASSv4) { $NetRoute = New-NetRoute -DestinationPrefix $DstNETv4s[$j] -NextHop $GwIPv4s[$i] -InterfaceIndex $IfIDv4s[$i] -RouteMetric $Metric -PolicyStore "ActiveStore"; };
    };
};
for ($i = 0; $i -lt $GwIPv6s.Count; $i++) {
    for ($j = 0; $j -lt $DstNETv6s.Count; $j++) {
        $isPASSv6 = $IfIDv6s[$i];
        foreach ($NetRoute in $NetRoutes) {
            if (($NetRoute.NextHop -eq $GwIPv6s[$i]) -and ($NetRoute.InterfaceIndex -eq $IfIDv6s[$i])) {
                if ($NetRoute.DestinationPrefix -eq $DstNETv6s[$j]) { $isPASSv6 = $false; };
            };
        };
        if ($isPASSv6) { $NetRoute = New-NetRoute -DestinationPrefix $DstNETv6s[$j] -NextHop $GwIPv6s[$i] -InterfaceIndex $IfIDv6s[$i] -RouteMetric $Metric -PolicyStore "ActiveStore"; };
    };
};