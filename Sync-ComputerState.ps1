<#
  .SYNOPSIS
  Включение или выключение целевого компьютера.

  .DESCRIPTION
  Скрипт включает широковещательным магическим пакетом целевой компьютер
  или выключает его, в зависимрсти от состояния доступности исходнного
  компьютера. Так же возможна дополнительная проверка другого участника
  для исключения ложного срабатывания при сбои сети.

  .PARAMETER Source
  Исходный компьютер для получения состояния.

  .PARAMETER Target
  Целевой компьютер для установки состояния.

  .PARAMETER Check
  Дополнительный компьютер для проверки.

  .PARAMETER MAC
  MAC адрес целевого компьютера для отправки мачического пакета.

  .PARAMETER Try
  Количество попыток проверки доступности каждого узла, а так же
  количество отправляемых магических пакетов.

  .PARAMETER Port
  Порт получателя магтческих пакетов.

  .INPUTS
  Вы не можете передавать в скрипт объекты по коневееру.

  .OUTPUTS
  Скрипт не генерит возвращаемые объекты.

  .NOTES
  Версия: 0.1.1
  Автор: @ViPiC
#>

[CmdletBinding()]
Param (
    [string]$Source,
    [string]$Target,
    [string]$Check,
    [string]$MAC,

    [ValidateRange(1, 99)]
    [int]$Try = 3,

    [ValidateRange(1, 65535)]
    [int]$Port = 9
);

$isCheckOnline = $true; $isSourceOnline = $true; $isTargetOnline = $true;
# Получаем информацию об ip адресах дополнительного компьютера и проверяем доступность
$ChecksIPv4 = @(); $ChecksIPv6 = @();
if ($Check) {
    $List = @($Check);
    if ([IPAddress]::TryParse($List[0], [ref][IPAddress]::Loopback)) {
        if (($List[0].IndexOf(".") -ge 0)) { $ChecksIPv4 = @($List[0]); };
        if (($List[0].IndexOf(":") -ge 0)) { $ChecksIPv6 = @($List[0]); };
    }
    else {
        $Resolve = (Resolve-DnsName -Name $List[0] -ErrorAction "Ignore" | Where-Object "Section" -EQ "Answer");
        if ($Resolve) { foreach ($CheckIPv4 in $Resolve.IP4Address) { $ChecksIPv4 += $CheckIPv4; }; };
        if ($Resolve) { foreach ($CheckIPv6 in $Resolve.IP6Address) { $ChecksIPv6 += $CheckIPv6; }; };
    };
    $isPassIPv4 = (($ChecksIPv4.Count -gt 0) -and (Test-Connection -ComputerName $ChecksIPv4 -Count $Try -Quiet -ErrorAction "Ignore"));
    $isPassIPv6 = (($ChecksIPv6.Count -gt 0) -and (Test-Connection -ComputerName $ChecksIPv6 -Count $Try -Quiet -ErrorAction "Ignore"));
    $isCheckOnline = ($isPassIPv4 -or $isPassIPv6);
};
# Получаем информацию об ip адресах исходного компьютера и проверяем доступность
$SourcesIPv4 = @(); $SourcesIPv6 = @();
if ($Source -and $isCheckOnline) {
    $List = @($Source);
    if ([IPAddress]::TryParse($List[0], [ref][IPAddress]::Loopback)) {
        if (($List[0].IndexOf(".") -ge 0)) { $SourcesIPv4 = @($List[0]); };
        if (($List[0].IndexOf(":") -ge 0)) { $SourcesIPv4 = @($List[0]); };
    }
    else {
        $Resolve = (Resolve-DnsName -Name $List[0] -ErrorAction "Ignore" | Where-Object "Section" -EQ "Answer");
        if ($Resolve) { foreach ($SourceIPv4 in $Resolve.IP4Address) { $SourcesIPv4 += $SourceIPv4; }; };
        if ($Resolve) { foreach ($SourceIPv6 in $Resolve.IP6Address) { $SourcesIPv6 += $SourceIPv6; }; };
    };
    $isPassIPv4 = (($SourcesIPv4.Count -gt 0) -and (Test-Connection -ComputerName $SourcesIPv4 -Count $Try -Quiet -ErrorAction "Ignore"));
    $isPassIPv6 = (($SourcesIPv6.Count -gt 0) -and (Test-Connection -ComputerName $SourcesIPv6 -Count $Try -Quiet -ErrorAction "Ignore"));
    $isSourceOnline = ($isPassIPv4 -or $isPassIPv6);
};
# Получаем информацию об ip адресах целевого компьютера и проверяем доступность
$TargetsIPv4 = @(); $TargetsIPv6 = @();
if ($Target -and $isCheckOnline) {
    $List = @($Target);
    if ([IPAddress]::TryParse($List[0], [ref][IPAddress]::Loopback)) {
        if (($List[0].IndexOf(".") -ge 0)) { $TargetsIPv4 = @($List[0]); };
        if (($List[0].IndexOf(":") -ge 0)) { $TargetsIPv4 = @($List[0]); };
    }
    else {
        $Resolve = (Resolve-DnsName -Name $List[0] -ErrorAction "Ignore" | Where-Object "Section" -EQ "Answer");
        if ($Resolve) { foreach ($TargetIPv4 in $Resolve.IP4Address) { $TargetsIPv4 += $TargetIPv4; }; };
        if ($Resolve) { foreach ($TargetIPv6 in $Resolve.IP6Address) { $TargetsIPv6 += $TargetIPv6; }; };
    };
    $isPassIPv4 = (($TargetsIPv4.Count -gt 0) -and (Test-Connection -ComputerName $TargetsIPv4 -Count $Try -Quiet -ErrorAction "Ignore"));
    $isPassIPv6 = (($TargetsIPv6.Count -gt 0) -and (Test-Connection -ComputerName $TargetsIPv6 -Count $Try -Quiet -ErrorAction "Ignore"));
    $isTargetOnline = ($isPassIPv4 -or $isPassIPv6);
};
# Выполняем действие над целевым компьютером
if ($isCheckOnline) {
    # Включаем целевой компьютер
    if (-not($isTargetOnline) -and $isSourceOnline -and $MAC) {
        $BroadcastProxy = [System.Net.IPAddress]::Broadcast;
        $chainSync = [byte[]](, 0xFF * 6); # цепочка синфронизации
        $MacAddress = $MAC -split "-" | ForEach-Object { [byte]("0x" + $PSItem) };
        $MagicPacket = $chainSync + $MacAddress * 16;
        $UdpClient = New-Object "System.Net.Sockets.UdpClient";
        $UdpClient.Connect($BroadcastProxy, $Port);
        for ($i = 0; $i -lt $Try; $i++) {
            $UdpClient.Send($MagicPacket, $MagicPacket.Length) | Out-Null;
        };
        $UdpClient.Close();
    };
    # Выключаем целевой компьютер
    if (-not($isSourceOnline) -and $isTargetOnline) {
        if ($Target) { Stop-Computer -ComputerName $Target -Force; }
        else { Stop-Computer -Force; };
    };
};
