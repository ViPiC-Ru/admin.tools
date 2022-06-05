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
  Версия: 0.1.0
  Автор: @ViPiC
#>

[CmdletBinding()]
Param (
    [string]$Source = "",
    [string]$Target = "",
    [string]$Check = "",
    [string]$MAC = "",

    [ValidateRange(1, 99)]
    [int]$Try = 3,

    [ValidateRange(1, 65535)]
    [int]$Port = 9
);

if ((-not $Check) -or (Test-Connection -ComputerName $Check -Quiet -Count $Try)) {
    if ((-not $Source) -or (Test-Connection -ComputerName $Source -Quiet -Count $Try)) {
        if ((-not $Target) -or (Test-Connection -ComputerName $Target -Quiet -Count $Try)) {
        }
        else {
            # Включаем целевой компьютер
            if ($MAC) {
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
        };
    }
    else {
        if ((-not $Target) -or (Test-Connection -ComputerName $Target -Quiet -Count $Try)) {
            # Выключаем целевой компьютер
            if ($Target) {
                Stop-Computer -ComputerName $Target -Force;
            }
            else {
                Stop-Computer -Force;
            };
        };
    };
};