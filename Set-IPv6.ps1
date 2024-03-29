﻿<#
  .SYNOPSIS
  Включает или отключает IPv6 на всех интерфейсах.

  .DESCRIPTION
  Скрипт выполняет команду включения или отключения
  IPv6 на всех интерфейсах.

  .PARAMETER Mode
  В какое состояние нужно перевести IPv6.

  .PARAMETER Check
  Предварительно проверить дополнительное условие на интерфейсе.

  .INPUTS
  Вы не можете передавать в скрипт объекты по коневееру.

  .OUTPUTS
  Скрипт не генерит возвращаемые объекты.

  .NOTES
  Версия: 0.1.2
  Автор: @ViPiC
#>

[CmdletBinding()]
Param (
    [Parameter (Mandatory = $true)]
    [ValidateSet("Enable", "Disable")]
    [string]$Mode,

    [ValidateSet("IPv4")]
    [string]$Check
);

$NetAdapterBindings = switch ($Check) {
  "IPv4" { Get-NetAdapterBinding -ComponentID "ms_tcpip" | Where-Object -Property "Enabled" | Get-NetAdapterBinding -ComponentID "ms_tcpip6"; }
  default { Get-NetAdapterBinding -ComponentID "ms_tcpip6"; }
};
switch ($Mode) {
    "Enable" { $NetAdapterBindings | Enable-NetAdapterBinding -ComponentID "ms_tcpip6" -PassThru; }
    "Disable" { $NetAdapterBindings | Disable-NetAdapterBinding -ComponentID "ms_tcpip6" -PassThru; }
};