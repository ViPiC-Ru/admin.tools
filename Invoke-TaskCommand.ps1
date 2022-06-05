<#
  .SYNOPSIS
  Выполняет команду через задачу на удаленных компьютерах.

  .DESCRIPTION
  Скрипт выполняет команду на удалённых компьютерах в контексте
  заданной учётной записи через планировщик заданий.

  .PARAMETER Command
  Командная строка которую нужно выполнить.

  .PARAMETER ComputerName
  Список компьютеров на которых нужно выполнить команду.

  .PARAMETER TaskName
  Имя временно создаваемой задачи в планировщике заданий.

  .PARAMETER AccountSid
  Идентификатор безопастности учётной записи в контексте которой
  нужно выполнить заданную команду.

  .PARAMETER ResultPath
  Путь к папке в которую нужно переместить результат работы скрипта
  из рабочий папке с удалённого компьютера.

  .INPUTS
  Вы можете передавать в скрипт список компьютеров по коневееру.

  .OUTPUTS
  В случае успеха скрипт возвращает результирующую папку или
  массив булевых значений успешности выполнения при отсутствие параметра
  с результирующей папкой.

  .NOTES
  Версия: 0.1.1
  Автор: @ViPiC
#>

[CmdletBinding()]
Param (
    [Parameter (Mandatory = $true)]
    [string]$Command,

    [Parameter (ValueFromPipeline = $true)]
    [string[]]$ComputerName = @($env:COMPUTERNAME),

    [string]$TaskName = "Temporary Task",

    [string]$AccountSid = "S-1-5-18",

    [string]$ResultPath
);

PROCESS {
    $Result = @();
    $IsCommandInvoke = $null;
    # Создаём папку для выгрузки результата
    if ($ResultPath) {
        if (Test-Path -Path $ResultPath) {
            Remove-Item -Path $ResultPath -Force -Recurse;
        };
        $ResultDirectory = New-Item -ItemType "Directory" -Path $ResultPath;
    };
    # Последовательно выполняем команду на удалённых компьютерах
    foreach ($ComputerNameItem in $ComputerName) {
        # Создаём сессию на удалённом компьютере для выполнения команды
        $RemoteSession = New-PSSession -ComputerName $ComputerNameItem -ErrorAction "Ignore";
        if ($RemoteSession) {
            # Выполняем удалённую команду и получаем рабочую папку
            $WorkingDirectory = Invoke-Command -Session $RemoteSession -ArgumentList $TaskName, $AccountSid, $Command, $ResultPath -ScriptBlock {
                # Принимаем параметры из родительского контекста
                Param ($TaskName, $AccountSid, $Command);
                # Создаём рабочую папку для выполнения команды
                $ParentDirectoryPath = [System.IO.Path]::GetTempPath();
                $WorkingDirectoryPath = Join-Path $ParentDirectoryPath $TaskName;
                if (Test-Path -Path $WorkingDirectoryPath) {
                    Remove-Item -Path $WorkingDirectoryPath -Force -Recurse;
                };
                $WorkingDirectory = New-Item -ItemType "Directory" -Path $WorkingDirectoryPath;
                # Добавляем в рабочую папку скрипт обёрку из JScript для скрытия окна консоли
                $WrapperName = "wrapper.js";
                $WrapperPath = Join-Path $WorkingDirectory $WrapperName;
                $WrapperScript = '(function(a,d){var e=[];d=new ActiveXObject("WScript.Shell");for(var c=0,f=a.arguments.length;c<f;c++){var b=a.arguments.item(c);-1!=b.indexOf(" ")&&(b="\""+b+"\"");e.push(b)}a=e.join(" ");if(a.length)try{d.run(a,0,!0)}catch(g){}})(WSH);';
                Set-Content -Path $WrapperPath -Value $WrapperScript -Encoding "ASCII";
                # Преобразуем SID в название учётной записи или группы
                $AccountIdentifier = New-Object System.Security.Principal.SecurityIdentifier($AccountSid);
                $AccountName = $AccountIdentifier.Translate([System.Security.Principal.NTAccount]).Value;
                # Выдаём полные права для учётной записи или группы на рабочую папку
                $WorkingDirectoryAccess = Get-Acl -Path $WorkingDirectory;
                $WorkingDirectoryAccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($AccountIdentifier, "FullControl", 3, "None", "Allow");
                $WorkingDirectoryAccess.SetAccessRule($WorkingDirectoryAccessRule);
                Set-Acl -Path $WorkingDirectory -AclObject $WorkingDirectoryAccess;
                # Создаём задачу в планировщике заданий
                $Command = $Command.Replace('$WorkingDirectory', $WorkingDirectoryPath);
                $TaskAction = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "/b $WrapperName $Command" -WorkingDirectory $WorkingDirectory;
                $TaskSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -DontStopOnIdleEnd -ExecutionTimeLimit (New-TimeSpan -Minutes 5);
                if ($AccountIdentifier.IsAccountSid()) {
                    $TaskPrincipal = New-ScheduledTaskPrincipal -UserId $AccountName;
                }
                else {
                    $TaskPrincipal = New-ScheduledTaskPrincipal -GroupId $AccountName;
                };
                Stop-ScheduledTask -TaskName $TaskName -ErrorAction "Ignore";
                Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction "Ignore";
                $Task = Register-ScheduledTask -TaskName $TaskName -Principal $TaskPrincipal -Action $TaskAction -Settings $TaskSettings;
                # Запускаем созданную задачу и дожидаемся завершения
                $IsCommandInvoke = $false;
                Start-ScheduledTask -TaskName $TaskName;
                $Task = Get-ScheduledTask -TaskName $TaskName;
                while ($Task.State -eq "Running") {
                    $IsCommandInvoke = $true;
                    $Task = Get-ScheduledTask -TaskName $TaskName;
                    Start-Sleep -Milliseconds 500;
                };
                # Удаляем созданную задачу
                Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false;
                # Удаляем файл скрипт обёртки
                Remove-Item -Path $WrapperPath;
                # Возвращаем результат выполнения
                if ($IsCommandInvoke) {
                    $WorkingDirectory;
                };
            };
            # Копируем рабочую папку с удалённого компьютера на локальный
            if ($ResultPath) {
                if ($WorkingDirectory) {
                    $ParentDirectoryPath = [System.IO.Path]::GetTempPath();
                    $TempDirectoryName = [System.IO.Path]::GetRandomFileName();
                    $TempDirectoryPath = Join-Path $ParentDirectoryPath $TempDirectoryName;
                    Copy-Item -Path $WorkingDirectory -FromSession $RemoteSession -Destination $TempDirectoryPath -Recurse;
                    $ResulItems = Get-ChildItem -Path $TempDirectoryPath;
                    foreach ($ResulItem in $ResulItems) {
                        $ResulItemPath = Join-Path $ResultDirectory $ResulItem.Name;
                        if (Test-Path -Path $ResulItemPath) {
                            Remove-Item -Path $ResulItemPath -Force -Recurse;
                        };
                        Copy-Item -Path $ResulItem.FullName -Destination $ResulItemPath -Force;
                    };
                    Remove-Item -Path $TempDirectoryPath -Recurse -Force;
                    $IsCommandInvoke = $IsCommandInvoke -or $true;
                }
                else {
                    $IsCommandInvoke = $IsCommandInvoke -or $false;
                };
            };
            # Удаляем рабочую папку на удалённом компьютере
            Invoke-Command -Session $RemoteSession -ScriptBlock { Remove-Item -Path $WorkingDirectory -Force -Recurse; };
            Remove-PSSession -Session $RemoteSession;
        }
        else {
            $WorkingDirectory = $null;
            if ($ResultPath) {
                $IsCommandInvoke = $IsCommandInvoke -or $false;
            };
        };
        # Готовим возвращаемый результат
        if ($ResultPath) {
            if ($WorkingDirectory) {
                $Result = $ResultDirectory;
            };
        }
        else {
            if ($WorkingDirectory) {
                $Result += $true;
            }
            else {
                $Result += $false;
            };
        };
    };
    # Работаем с результирующей папкой
    if ($ResultPath -and -not($IsCommandInvoke)) {
        Remove-Item -Path $ResultPath -Force -Recurse;
    };
    # Возвращаем результат
    return $Result;
};
