<#
  .SYNOPSIS
  Выполняет команду через задачу на удаленных компьютерах.

  .DESCRIPTION
  Скрипт выполняет команду на удалённых компьютерах в контексте
  заданной учётной записи через планировщик заданий.

  .PARAMETER Command
  Командная строка которую нужно выполнить, поддерживается
  шаблон {host} и {input}, {output} для указания полных путей к папкам
  ввода и вывода на исполняемом компьютере.

  .PARAMETER ComputerName
  Список компьютеров на которых нужно выполнить команду.

  .PARAMETER TaskName
  Имя временно создаваемой задачи в планировщике заданий.

  .PARAMETER AccountSID
  Идентификатор безопастности учётной записи в контексте которой
  нужно выполнить заданную команду.

  .PARAMETER InputPath
  Путь к папке которую нужно скопировать на удалённый компьютер
  для выполнения командной строки в контексте этой папки.

  .PARAMETER OutputPath
  Путь к папке в которую нужно переместить результат работы скрипта
  из папки вывода с удалённого компьютера.

  .PARAMETER IncludeWQL
  Список WQL запросов для удалённого компьютера, все из которых
  должны вернуть хотябы один элимент для оставления компьюера в списке.
  Двойные кавычки заменяются на одинарные.
  
  .PARAMETER ExcludeWQL
  Список WQL запросов для удалённого компьютера, хотябы один из которых
  должн вернуть хотябы один элимент для исключения компьютера из списка.
  Двойные кавычки заменяются на одинарные.

  .INPUTS
  Вы можете передавать в скрипт список компьютеров по коневееру.

  .OUTPUTS
  В случае успеха скрипт возвращает результирующую папку или
  массив объектов с результатами выполнения команды при отсутствие параметра
  с папкой вывода.

  .NOTES
  Версия: 0.3.3
  Автор: @ViPiC
#>

#Requires -Modules ThreadJob

[CmdletBinding()]
Param (
    [Parameter (Mandatory = $true)]
    [string]$Command,

    [Parameter (ValueFromPipeline = $true)]
    [string[]]$ComputerName = @($env:COMPUTERNAME),

    [string]$TaskName = "Temporary Task",

    [string]$AccountSID = "S-1-5-18",

    [string]$InputPath,

    [string]$OutputPath,

    [string[]]$IncludeWQL = @(),

    [string[]]$ExcludeWQL = @()
);

PROCESS {
    $Result = @(); # Возвращаемый результат
    $JobTimes = @{}; # Хеш таблица временных меток стартов выполнения заданий
    $JobResults = @{}; # Хеш таблица результатов для сбойных заданий
    $Jobs = [System.Collections.ArrayList]@(); # Список фоновых заданий
    $JobLimit = 16; # Количество параллельных задач для выполнения
    $JobTimeOut = 70; # Максимальное время выполнения задания в минутах
    $TaskTimeOut = 60; # Максимальное время выполнения задачи в минутах
    $CustomError = 0; # Внутренняя ошибка при выполнении скрипта
    $WrapperName = "wrapper.js"; # Имя файла скрипта обёртки для скрытия консоли
    $InputName = "input"; # Имя удалённой директории для ввода
    $HostName = "host"; # Имя удалённого компьютера
    $OutputName = "output"; # Имя удалённой директории для вывода
    $SuccessRunCount = 0; # Счётчик успешных запусков команд
    # Получаем директорию ввода
    if (-not($CustomError) -and $InputPath) {
        $LocalInputDirectory = Get-Item -Path $InputPath;
        if (-not($LocalInputDirectory)) { $CustomError = 1; };
    };
    # Получаем или создаём директорию вывода
    if (-not($CustomError) -and $OutputPath) {
        if (Test-Path -Path $OutputPath) {
            $LocalOutputDirectory = Get-Item -Path $OutputPath;
        }
        else {
            $LocalOutputDirectory = New-Item -ItemType "Directory" -Path $OutputPath;
        };
        if (-not($LocalOutputDirectory)) { $CustomError = 2; };
    };
    # Последовательно выполняем команду на удалённых компьютерах
    if (-not($CustomError)) {
        $Index = 0; # Сбрасываем индекс для перебора элиментов
        while ((-not $Index) -or $Jobs.Count) {
            if (($Jobs.Count -lt $JobLimit) -and ($Index -lt $ComputerName.Count)) {
                # Заполняем пул заданий до указанного лимита
                $ComputerNameItem = $ComputerName[$Index]; # Получаем очередной элимент
                $JobTime = Get-Date;
                $Job = Start-ThreadJob -ArgumentList $ComputerNameItem, $IncludeWQL, $ExcludeWQL, $InputPath, $OutputPath, $LocalInputDirectory, $LocalOutputDirectory, $AccountSID, $WrapperName, $HostName, $InputName, $OutputName, $TaskName, $TaskTimeOut, $Command -ThrottleLimit $JobLimit -ScriptBlock {
                    # Принимаем параметры из родительского контекста
                    Param ($ComputerNameItem, $IncludeWQL, $ExcludeWQL, $InputPath, $OutputPath, $LocalInputDirectory, $LocalOutputDirectory, $AccountSID, $WrapperName, $HostName, $InputName, $OutputName, $TaskName, $TaskTimeOut, $Command);
                    # Непосредственно само задание
                    $CustomError = 0; # Сбрасываем внутренную ошибка выполнения скрипта
                    $IsCommandRun = $false; # Был ли непосредственный запуск команды
                    $IsСheckPass = $null; # Прошёл ли удалённый компьютер проверку
                    $RemoteSession = $null; # Сбрасываем сессию на удалённом компьютере
                    $RemoteTempDirectory = $null; # Сбрасываем удалённую временную директорию
                    $TaskResult = $null; # Сбрасываем строку с кодом возврата выполнения команды
                    # Создаём сессию на удалённом компьютере
                    if (-not($CustomError)) {
                        $RemoteSession = New-PSSession -ComputerName $ComputerNameItem -ErrorAction "Ignore";
                        if (-not($RemoteSession)) { $CustomError = 3; };
                    };
                    # Проверяем соответствие удалённого компьютера фильтрам
                    if (-not($CustomError)) {
                        $IsСheckPass = Invoke-Command -Session $RemoteSession -ArgumentList $IncludeWQL, $ExcludeWQL -ScriptBlock {
                            # Принимаем параметры из родительского контекста
                            Param ($IncludeWQL, $ExcludeWQL);
                            # Выполняем проверки
                            $IsСheckPass = $true;
                            # Проверяем фильтры для исключения
                            for ($Index = 0; $Index -lt $ExcludeWQL.Count -and $IsСheckPass; $Index++) {
                                $WQL = $ExcludeWQL[$Index].Replace('"', "'");
                                $CimResponse = Get-CimInstance -Query $WQL -ErrorAction "Ignore";
                                if ($CimResponse) { $IsСheckPass = $false; };
                            };
                            # Проверяем фильтры для включения
                            for ($Index = 0; $Index -lt $IncludeWQL.Count -and $IsСheckPass; $Index++) {
                                $WQL = $IncludeWQL[$Index].Replace('"', "'");
                                $CimResponse = Get-CimInstance -Query $WQL -ErrorAction "Ignore";
                                if (-not($CimResponse)) { $IsСheckPass = $false; };
                            };
                            # Возвращаем результат
                            $IsСheckPass;
                        };
                    };
                    # Создаём временную директорию на удалённом компьютере
                    if (-not($CustomError) -and $IsСheckPass) {
                        $RemoteTempDirectory = Invoke-Command -Session $RemoteSession -ArgumentList $WrapperName -ScriptBlock {
                            # Принимаем параметры из родительского контекста
                            Param ($WrapperName);
                            # Генерируем имя для временной папки
                            $Stream = [System.IO.MemoryStream]::New();
                            $Writer = [System.IO.StreamWriter]::New($Stream);
                            $Writer.Write($TaskName);
                            $Writer.Flush();
                            $Stream.Position = 0;
                            $TempName = (Get-FileHash -InputStream $Stream -Algorithm "MD5").Hash.SubString(22);
                            # Создаём временную папку для выполнения команды
                            $TempPath = Join-Path ([System.IO.Path]::GetTempPath()) $TempName;
                            if (Test-Path -Path $TempPath) { Remove-Item -Path $TempPath -Force -Recurse; };
                            $RemoteTempDirectory = New-Item -ItemType "Directory" -Path $TempPath;
                            # Добавляем во временную папку скрипт обёрку из JScript для скрытия окна консоли
                            $WrapperPath = Join-Path $RemoteTempDirectory.FullName $WrapperName;
                            $WrapperScript = '(function(b,d){var e=[],f=0;d=new ActiveXObject("WScript.Shell");for(var c=0,g=b.arguments.length;c<g;c++){var a=b.arguments.item(c);-1!=a.indexOf(" ")&&(a="\""+a+"\"");e.push(a)}a=e.join(" ");if(a.length)try{f=d.run(a,0,!0)}catch(h){}b.quit(f)})(WSH);';
                            Set-Content -Path $WrapperPath -Value $WrapperScript -Encoding "ASCII";
                            # Возвращаем результат
                            $RemoteTempDirectory;
                        };
                        if (-not($RemoteTempDirectory)) { $CustomError = 4; };
                    };
                    # Создаём директорию ввода на удалённом компьютере
                    if (-not($CustomError) -and $IsСheckPass) {
                        $RemoteInputDirectory = Invoke-Command -Session $RemoteSession -ArgumentList $RemoteTempDirectory, $InputName -ScriptBlock {
                            # Принимаем параметры из родительского контекста
                            Param ($RemoteTempDirectory, $InputName);
                            # Создаём дерикторию ввода
                            $RemoteInputDirectory = New-Item -ItemType "Directory" -Path (Join-Path $RemoteTempDirectory.FullName $InputName);
                            # Возвращаем результат
                            $RemoteInputDirectory;
                        };
                        if (-not($RemoteInputDirectory)) { $CustomError = 5; };
                    };
                    # Создаём директорию вывода на удалённом компьютере
                    if (-not($CustomError) -and $IsСheckPass) {
                        $RemoteOutputDirectory = Invoke-Command -Session $RemoteSession -ArgumentList $RemoteTempDirectory, $OutputName -ScriptBlock {
                            # Принимаем параметры из родительского контекста
                            Param ($RemoteTempDirectory, $OutputName);
                            # Создаём дерикторию вывода
                            $RemoteOutputDirectory = New-Item -ItemType "Directory" -Path (Join-Path $RemoteTempDirectory.FullName $OutputName);
                            # Возвращаем результат
                            $RemoteOutputDirectory;
                        };
                        if (-not($RemoteOutputDirectory)) { $CustomError = 6; };
                    };
                    # Копируем содержимое директории ввода на удалённый компьютер
                    if (-not($CustomError) -and $IsСheckPass -and $InputPath) {
                        $Items = Get-ChildItem -Path $LocalInputDirectory.FullName;
                        foreach ($Item in $Items) {
                            $ItemPath = Join-Path $RemoteInputDirectory.FullName $Item.Name;
                            Copy-Item -Path $Item.FullName -Destination $ItemPath -ToSession $RemoteSession -Force -Recurse;
                        };
                    };
                    # Назначаем права и выполняем команду на удалённом компьютере через задачу
                    if (-not($CustomError) -and $IsСheckPass) {
                        $InvokeResult = Invoke-Command -Session $RemoteSession -ArgumentList $ComputerNameItem, $RemoteTempDirectory, $RemoteInputDirectory, $RemoteOutputDirectory, $AccountSID, $WrapperName, $HostName, $InputName, $OutputName, $TaskName, $TaskTimeOut, $Command -ScriptBlock {
                            # Принимаем параметры из родительского контекста
                            Param ($ComputerNameItem, $RemoteTempDirectory, $RemoteInputDirectory, $RemoteOutputDirectory, $AccountSID, $WrapperName, $HostName, $InputName, $OutputName, $TaskName, $TaskTimeOut, $Command);
                            # Преобразуем SID в название учётной записи или группы
                            $AccountIdentifier = New-Object System.Security.Principal.SecurityIdentifier($AccountSID);
                            $AccountName = $AccountIdentifier.Translate([System.Security.Principal.NTAccount]).Value;
                            # Выдаём полные права для учётной записи или группы на рабочую папку
                            $DirectoryAccess = Get-Acl -Path $RemoteTempDirectory;
                            $DirectoryAccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($AccountIdentifier, "FullControl", 3, "None", "Allow");
                            $DirectoryAccess.SetAccessRule($DirectoryAccessRule);
                            Set-Acl -Path $RemoteTempDirectory -AclObject $DirectoryAccess;
                            # Создаём задачу в планировщике заданий
                            $FixedCommand = $Command.Replace("{$HostName}", $ComputerNameItem).Replace("{$InputName}", $RemoteInputDirectory.FullName).Replace("{$OutputName}", $RemoteOutputDirectory.FullName);
                            $TaskAction = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "/b ..\$WrapperName $FixedCommand" -WorkingDirectory $RemoteInputDirectory;
                            $TaskSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -DontStopOnIdleEnd -ExecutionTimeLimit (New-TimeSpan -Minutes $TaskTimeOut);
                            $TaskPrincipal = if ($AccountIdentifier.IsAccountSid()) { New-ScheduledTaskPrincipal -UserId $AccountName }else { New-ScheduledTaskPrincipal -GroupId $AccountName };
                            Stop-ScheduledTask -TaskName $TaskName -ErrorAction "Ignore";
                            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction "Ignore";
                            $Task = Register-ScheduledTask -TaskName $TaskName -Principal $TaskPrincipal -Action $TaskAction -Settings $TaskSettings;
                            # Запускаем созданную задачу и дожидаемся завершения
                            $IsCommandRun = $false;
                            Start-ScheduledTask -TaskName $TaskName;
                            $Task = Get-ScheduledTask -TaskName $TaskName;
                            while ($Task.State -eq "Running") {
                                $IsCommandRun = $true;
                                $Task = Get-ScheduledTask -TaskName $TaskName;
                                Start-Sleep -Milliseconds 500;
                            };
                            # Получаем информацию и удаляем созданную задачу
                            $TaskInfo = Get-ScheduledTaskInfo -TaskName $TaskName;
                            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false;
                            if ($IsCommandRun) { $TaskResult = $TaskInfo.LastTaskResult; };
                            # Возвращаем результат
                            $InvokeResult = @{"TaskResult" = $TaskResult; "IsCommandRun" = $IsCommandRun; };
                            $InvokeResult;
                        };
                        $IsCommandRun = $InvokeResult.IsCommandRun;
                        if ($IsCommandRun) {
                            $TaskResult = $InvokeResult.TaskResult;
                        };
                    };
                    # Копируем содержимое директории вывода с удалённого компьютера
                    if (-not($CustomError) -and $IsСheckPass -and $OutputPath -and $IsCommandRun) {
                        $Items = Invoke-Command -Session $RemoteSession -ArgumentList $RemoteOutputDirectory -ScriptBlock {
                            # Принимаем параметры из родительского контекста
                            Param ($RemoteOutputDirectory);
                            # Получаем список элиментов в папке
                            $Items = Get-ChildItem -Path $RemoteOutputDirectory.FullName;
                            # Возвращаем результат
                            $Items;
                        };
                        foreach ($Item in $Items) {
                            $ItemPath = Join-Path $LocalOutputDirectory.FullName $Item.Name;
                            Copy-Item -Path $Item.FullName -Destination $ItemPath -FromSession $RemoteSession -Force -Recurse;
                        };
                    };
                    # Удаляем временную директорию на удалённый компьютер
                    if ($RemoteTempDirectory) {
                        Invoke-Command -Session $RemoteSession -ArgumentList $RemoteTempDirectory -ScriptBlock {
                            # Принимаем параметры из родительского контекста
                            Param ($RemoteTempDirectory);
                            # Удаляем временную директорию
                            Remove-Item -Path $RemoteTempDirectory.FullName -Force -Recurse;
                        };
                    };
                    # Завершаем сессию на удалённом компьютере
                    if ($RemoteSession) {
                        Remove-PSSession -Session $RemoteSession;
                    };
                    # Возвращаем результат
                    $JobResult = @{"ComputerNameItem" = $ComputerNameItem; "TaskResult" = $TaskResult; "IsСheckPass" = $IsСheckPass; "IsCommandRun" = $IsCommandRun; };
                    $JobResult;
                };
                $Jobs.Add($Job) | Out-Null;
                $JobResults.Add($Job.Id, @{"ComputerNameItem" = $ComputerNameItem; "TaskResult" = $null; "IsСheckPass" = $null; "IsCommandRun" = $false; });
                $JobTimes.Add($Job.Id, $JobTime);
                $Index++;
            }
            else {
                # Получаем результат выполненого задания из пула
                Wait-Job -Job $Jobs -Any -Timeout 3 | Out-Null;
                foreach ($Job in $Jobs) {
                    $NowTime = Get-Date;
                    $JobTime = $JobTimes.Item($Job.Id);
                    if ((New-TimeSpan -Start $JobTime -End $NowTime).TotalMinutes -gt $JobTimeOut) {
                        Stop-Job -Job $Job;
                    };
                };
                $States = @("Completed", "Failed", "Stopped", "Suspended", "Disconnected");
                $JobsComplete = $Jobs | Where-Object -Property "State" -in -Value $States;
                foreach ($Job in $JobsComplete) {
                    $JobResult = Receive-Job -Job $Job;
                    if (-not $JobResult) {
                        $JobResult = $JobResults.Item($Job.Id);
                    };
                    Remove-Job -Job $Job;
                    $Jobs.Remove($Job);
                    # Присваеваем значение переменным
                    $ComputerNameItem = $JobResult.ComputerNameItem;
                    $TaskResult = $JobResult.TaskResult;
                    $IsСheckPass = $JobResult.IsСheckPass;
                    $IsCommandRun = $JobResult.IsCommandRun;
                    # Обрабатываем результат выполнения задания
                    if (-not($TaskResult) -and $IsCommandRun) {
                        $SuccessRunCount++;
                    };
                    if (-not($OutputPath) -and $IsСheckPass) {
                        $Result += [PSCustomObject]@{
                            "ComputerName" = $ComputerNameItem;
                            "TaskResult"   = $TaskResult;
                        };
                    };
                };
            };
        };
    };
    # Работаем с директорией вывода и результатом
    if ($OutputPath) {
        if ($SuccessRunCount) {
            $Result = $LocalOutputDirectory;
        }
        else {
            $Result = $null;
        };
    };
    # Возвращаем результат
    $Result;
};
