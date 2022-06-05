<#
  .SYNOPSIS
  Добавляет постоянный статический маршрут к адресу по домену.

  .DESCRIPTION
  Скрипт проверяет список IPv4 маршрутов по 32 маске с заданой метрикой.
  Если необходимый маршрут отсутствует, то добовляет его и удаляет
  устаревшие маршруты.

  .PARAMETER HostName
  Доменное имя хоста до которого нужно добавить маршрут.

  .PARAMETER GateName
  Доменное имя шлюза через который нужно добавить маршрут.

  .PARAMETER RouteMetric
  Метрика маршрута, которая может выступать как дополнительный фильтр
  при формирования списка маршрутов для проверки.

  .INPUTS
  Вы не можете передавать в скрипт объекты по коневееру.

  .OUTPUTS
  Скрипт не генерит возвращаемые объекты.

  .NOTES
  Версия: 0.1.0
  Автор: @ViPiC
#>

[CmdletBinding()]
Param (
    [Parameter (Mandatory = $true)]
    [string]$HostName,

    [Parameter (Mandatory = $true)]
    [string]$GateName,

    [int]$RouteMetric = 1
);

# Получаем вспомогательные данные
$NetPrefix = "/32";
$HostIPv4 = (Resolve-DnsName -Name $HostName -ErrorAction "Stop").IP4Address | Select-Object -First 1;
$GateIPv4 = (Resolve-DnsName -Name $GateName -ErrorAction "Stop").IP4Address | Select-Object -First 1;
$InterfaceIndex = (Find-NetRoute -RemoteIPAddress $HostIPv4 | Select-Object -First 1).InterfaceIndex;

# Удаляем устаревшие маршруты
$IsRouteFound = $false;
foreach ($NetRoute in (Get-NetRoute -RouteMetric $RouteMetric -DestinationPrefix ("*" + $NetPrefix) -PolicyStore "ActiveStore" -ErrorAction "Ignore")) {
    $IsThisRoute = $true;
    $IsThisRoute = $IsThisRoute -and ($NetRoute.NextHop -eq $GateIPv4);
    $IsThisRoute = $IsThisRoute -and ($NetRoute.InterfaceIndex -eq $InterfaceIndex);
    $IsThisRoute = $IsThisRoute -and ($NetRoute.DestinationPrefix -eq ($HostIPv4 + $NetPrefix));
    if ($IsThisRoute) {
        $IsRouteFound = $true;
    }
    else {
        Remove-NetRoute -InputObject $NetRoute -Confirm:$false;
    };
};

# Добавляем новый маршрут
if (-not $IsRouteFound) {
    New-NetRoute -DestinationPrefix ($HostIPv4 + $NetPrefix) -NextHop $GateIPv4 -InterfaceIndex $InterfaceIndex -RouteMetric $RouteMetric -PolicyStore "ActiveStore" | Out-Null;
};