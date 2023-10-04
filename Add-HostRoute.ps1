<#
  .SYNOPSIS
  Добавляет и контролирует временный маршрут в приделах метрики.

  .DESCRIPTION
  Скрипт проверяет список IPv4 и IPv6 маршрутов в приделах заданной
  метрики. Если необходимый маршрут отсутствует, то добавляет его и удаляет
  устаревшие маршруты. Чтобы маршрут добавился, необходимо наличие
  IP адресов нужной версии у всех хостов, задействованных в процессе.

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
  Версия: 0.2.3
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
$isPassIPv4 = $true; $isPassIPv6 = $true;
# Получаем информацию об ip адресах шлюза
$GatewaysIPv4 = @(); $GatewaysIPv6 = @();
if ($Gateway -and ($isPassIPv4 -or $isPassIPv6)) {
    $List = @($Gateway);
    if ([IPAddress]::TryParse($List[0], [ref][IPAddress]::Loopback)) {
        if (($List[0].IndexOf(".") -ge 0) -and $isPassIPv4) { $GatewaysIPv4 = @($List[0]); };
        if (($List[0].IndexOf(":") -ge 0) -and $isPassIPv6) { $GatewaysIPv6 = @($List[0]); };
    }
    else {
        $Resolve = (Resolve-DnsName -Name $List[0] -ErrorAction "Ignore" | Where-Object "Section" -EQ "Answer");
        if ($Resolve -and $isPassIPv4) { foreach ($GatewayIPv4 in $Resolve.IP4Address) { $GatewaysIPv4 += $GatewayIPv4; }; };
        if ($Resolve -and $isPassIPv6) { foreach ($GatewayIPv6 in $Resolve.IP6Address) { $GatewaysIPv6 += $GatewayIPv6; }; };
    };
    $isPassIPv4 = ($GatewaysIPv4.Count -gt 0);
    $isPassIPv6 = ($GatewaysIPv6.Count -gt 0);
};
# Получаем информацию об ip адресах проверяемого узла и проверяем доступность
$ChecksIPv4 = @(); $ChecksIPv6 = @();
if ($Check -and ($isPassIPv4 -or $isPassIPv6)) {
    $List = @($Check);
    if ([IPAddress]::TryParse($List[0], [ref][IPAddress]::Loopback)) {
        if (($List[0].IndexOf(".") -ge 0) -and $isPassIPv4) { $ChecksIPv4 = @($List[0]); };
        if (($List[0].IndexOf(":") -ge 0) -and $isPassIPv6) { $ChecksIPv6 = @($List[0]); };
    }
    else {
        $Resolve = (Resolve-DnsName -Name $List[0] -ErrorAction "Ignore" | Where-Object "Section" -EQ "Answer");
        if ($Resolve -and $isPassIPv4) { foreach ($CheckIPv4 in $Resolve.IP4Address) { $ChecksIPv4 += $CheckIPv4; }; };
        if ($Resolve -and $isPassIPv6) { foreach ($CheckIPv6 in $Resolve.IP6Address) { $ChecksIPv6 += $CheckIPv6; }; };
    };
    $isPassIPv4 = (($ChecksIPv4.Count -gt 0) -and (Test-Connection -ComputerName $ChecksIPv4 -Quiet -ErrorAction "Ignore"));
    $isPassIPv6 = (($ChecksIPv6.Count -gt 0) -and (Test-Connection -ComputerName $ChecksIPv6 -Quiet -ErrorAction "Ignore"));
};
# Получаем информацию об ip адресах хоста назначения
$DestinationsIPv4 = @(); $DestinationsIPv6 = @();
if ($Destination -and ($isPassIPv4 -or $isPassIPv6)) {
    $List = $Destination.Split($MaskDelim);
    if ([IPAddress]::TryParse($List[0], [ref][IPAddress]::Loopback)) {
        if (($List[0].IndexOf(".") -ge 0) -and $isPassIPv4) { $DestinationsIPv4 = @($List[0]); };
        if (($List[0].IndexOf(":") -ge 0) -and $isPassIPv6) { $DestinationsIPv6 = @($List[0]); };
    }
    else {
        $Resolve = (Resolve-DnsName -Name $List[0] -ErrorAction "Ignore" | Where-Object "Section" -EQ "Answer");
        if ($Resolve -and $isPassIPv4) { foreach ($DestinationIPv4 in $Resolve.IP4Address) { $DestinationsIPv4 += $DestinationIPv4; }; };
        if ($Resolve -and $isPassIPv6) { foreach ($DestinationIPv6 in $Resolve.IP6Address) { $DestinationsIPv6 += $DestinationIPv6; }; };
    };
    $isPassIPv4 = ($DestinationsIPv4.Count -gt 0);
    $isPassIPv6 = ($DestinationsIPv6.Count -gt 0);
};
# Прнобразовываем ip адреса хоста назначения в адреса подсетей
$DestinationsNETv4 = @(); $DestinationsNETv6 = @();
if ($Destination -and ($isPassIPv4 -or $isPassIPv6)) {
    $List = $Destination.Split($MaskDelim);
    $CIDRv4 = if ($List.Count -eq 2) { [Math]::Max(0, [Math]::Min($List[1] * 1, $maxCIDRv4)) } else { $maxCIDRv4 };
    foreach ($DestinationIPv4 in $DestinationsIPv4) {
        $IPv4 = [IPAddress]::Parse($DestinationIPv4);
        $MASKv4 = [IPAddress]([Math]::Pow(2, $maxCIDRv4) - 1 -bxor [Math]::Pow(2, ($maxCIDRv4 - $CIDRv4)) - 1);
        $NETv4 = [IPAddress]($IPv4.Address -band $MASKv4.Address);
        $DestinationNETv4 = $NETv4.IPAddressToString + $MaskDelim + $CIDRv4;
        if (-not($DestinationNETv4 -in $DestinationsNETv4)) { $DestinationsNETv4 += $DestinationNETv4; };
    };
    $CIDRv6 = if ($List.Count -eq 2) { [Math]::Max(0, [Math]::Min($List[1] * 1, $maxCIDRv6)) } else { $maxCIDRv6 };
    foreach ($DestinationIPv6 in $DestinationsIPv6) {
        $IPv6 = [IPAddress]::Parse($DestinationIPv6);
        $NETv6 = $IPv6;
        $DestinationNETv6 = $NETv6.IPAddressToString + $MaskDelim + $CIDRv6;
        if (-not($DestinationNETv6 -in $DestinationsNETv6)) { $DestinationsNETv6 += $DestinationNETv6; };
    };
};
# Определяем интерфейсы для шлюзов
$InterfacesIDXv4 = @(); $InterfacesIDXv6 = @();
if ($Gateway -and ($isPassIPv4 -or $isPassIPv6)) {
    foreach ($GatewayIPv4 in $GatewaysIPv4) {
        $Find = Find-NetRoute -RemoteIPAddress $GatewayIPv4 -ErrorAction "Ignore";
        $InterfaceIDXv4 = if ($Find) { $Find[0].InterfaceIndex } else { $null };
        $InterfacesIDXv4 += $InterfaceIDXv4;
    };
    foreach ($GatewayIPv6 in $GatewaysIPv6) {
        $Find = Find-NetRoute -RemoteIPAddress $GatewayIPv6 -ErrorAction "Ignore";
        $InterfaceIDXv6 = if ($Find) { $Find[0].InterfaceIndex } else { $null };
        $InterfacesIDXv6 += $InterfaceIDXv6;
    };
};
# Удаляем устаревшие маршруты
$NetRoutes = (Get-NetRoute -RouteMetric $Metric -PolicyStore "ActiveStore" -ErrorAction "Ignore" | Where-Object "NextHop" -NotIn "::", "0.0.0.0");
if (-not($NetRoutes)) { $NetRoutes = @(); };
foreach ($NetRoute in $NetRoutes) {
    $isPassIPv4 = $false;
    for ($i = 0; ($i -lt $GatewaysIPv4.Count) -and -not($isPassIPv4); $i++) {
        if (($NetRoute.NextHop -eq $GatewaysIPv4[$i]) -and ($NetRoute.InterfaceIndex -eq $InterfacesIDXv4[$i])) {
            for ($j = 0; ($j -lt $DestinationsNETv4.Count) -and -not($isPassIPv4); $j++) {
                if ($NetRoute.DestinationPrefix -eq $DestinationsNETv4[$j]) { $isPassIPv4 = $true; };
            };
        };
    };
    $isPassIPv6 = $false;
    for ($i = 0; ($i -lt $GatewaysIPv6.Count) -and -not($isPassIPv6); $i++) {
        if (($NetRoute.NextHop -eq $GatewaysIPv6[$i]) -and ($NetRoute.InterfaceIndex -eq $InterfacesIDXv6[$i])) {
            for ($j = 0; ($j -lt $DestinationsNETv6.Count) -and -not($isPassIPv6); $j++) {
                if ($NetRoute.DestinationPrefix -eq $DestinationsNETv6[$j]) { $isPassIPv6 = $true; };
            };
        };
    };
    if (-not($isPassIPv4 -or $isPassIPv6)) { Remove-NetRoute -InputObject $NetRoute -Confirm:$false; };
};
# Добавляем новые маршруты
for ($i = 0; $i -lt $GatewaysIPv4.Count; $i++) {
    for ($j = 0; $j -lt $DestinationsNETv4.Count; $j++) {
        $isPassIPv4 = $InterfacesIDXv4[$i];
        foreach ($NetRoute in $NetRoutes) {
            if (($NetRoute.NextHop -eq $GatewaysIPv4[$i]) -and ($NetRoute.InterfaceIndex -eq $InterfacesIDXv4[$i])) {
                if ($NetRoute.DestinationPrefix -eq $DestinationsNETv4[$j]) { $isPassIPv4 = $false; };
            };
        };
        if ($isPassIPv4) { $NetRoute = New-NetRoute -DestinationPrefix $DestinationsNETv4[$j] -NextHop $GatewaysIPv4[$i] -InterfaceIndex $InterfacesIDXv4[$i] -RouteMetric $Metric -PolicyStore "ActiveStore"; };
    };
};
for ($i = 0; $i -lt $GatewaysIPv6.Count; $i++) {
    for ($j = 0; $j -lt $DestinationsNETv6.Count; $j++) {
        $isPassIPv6 = $InterfacesIDXv6[$i];
        foreach ($NetRoute in $NetRoutes) {
            if (($NetRoute.NextHop -eq $GatewaysIPv6[$i]) -and ($NetRoute.InterfaceIndex -eq $InterfacesIDXv6[$i])) {
                if ($NetRoute.DestinationPrefix -eq $DestinationsNETv6[$j]) { $isPassIPv6 = $false; };
            };
        };
        if ($isPassIPv6) { $NetRoute = New-NetRoute -DestinationPrefix $DestinationsNETv6[$j] -NextHop $GatewaysIPv6[$i] -InterfaceIndex $InterfacesIDXv6[$i] -RouteMetric $Metric -PolicyStore "ActiveStore"; };
    };
};