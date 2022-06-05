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
  
  .PARAMETER ExcludeWQL
  Список WQL запросов для удалённого компьютера, хотябы один из которых
  должн вернуть хотябы один элимент для исключения компьютера из списка.

  .INPUTS
  Вы можете передавать в скрипт список компьютеров по коневееру.

  .OUTPUTS
  В случае успеха скрипт возвращает результирующую папку или
  массив объектов с результатами выполнения команды при отсутствие параметра
  с папкой вывода.

  .NOTES
  Версия: 0.2.0
  Автор: @ViPiC
#>

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
    $CustomError = 0; # Внутренняя ошибка при выполнении скрипта
    $TaskTimeOut = 15; # Максимальное время выполнения задачи в минутах
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
    # Создаём директорию вывода
    if (-not($CustomError) -and $OutputPath) {
        if (Test-Path -Path $OutputPath) {
            foreach ($Item in (Get-ChildItem -Path $OutputPath)) {
                Remove-Item -Path $Item.FullName -Force -Recurse;
            };
            $LocalOutputDirectory = Get-Item -Path $OutputPath;
        }
        else {
            $LocalOutputDirectory = New-Item -ItemType "Directory" -Path $OutputPath;
        };
        if (-not($LocalOutputDirectory)) { $CustomError = 2; };
    };
    # Последовательно выполняем команду на удалённых компьютерах
    if (-not($CustomError)) {
        foreach ($ComputerNameItem in $ComputerName) {
            $IsCommandRun = $false; # Был ли непосредственный запуск команды
            $IsСheckPass = $null; # Прошёл ли удалённый компьютер проверку
            $RemoteSession = $null; # Сбрасываем сессию на удалённом компьютере
            $RemoteTempDirectory = $null; # Сбрасываем удалённую временную директорию
            $LocalTempDirectory = $null; # Сбрасываем локальную временную директорию
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
                        $WQL = $ExcludeWQL[$Index]; # Получаем очередной запрос
                        $WmiResponse = Get-WmiObject -Query $WQL -ErrorAction "Ignore";
                        if ($WmiResponse) { $IsСheckPass = $false; };
                    };
                    # Проверяем фильтры для включения
                    for ($Index = 0; $Index -lt $IncludeWQL.Count -and $IsСheckPass; $Index++) {
                        $WQL = $IncludeWQL[$Index]; # Получаем очередной запрос
                        $WmiResponse = Get-WmiObject -Query $WQL -ErrorAction "Ignore";
                        if (-not($WmiResponse)) { $IsСheckPass = $false; };
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
                    # Создаём временную папку для выполнения команды
                    $TempPath = Join-Path ([System.IO.Path]::GetTempPath()) ([System.IO.Path]::GetRandomFileName());
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
            # Копируем или создаём директорию ввода на удалённый компьютере
            if (-not($CustomError) -and $IsСheckPass) {
                if ($InputPath) { Copy-Item -Path $LocalInputDirectory.FullName -ToSession $RemoteSession -Destination $RemoteTempDirectory.FullName -Recurse; };
                $RemoteInputDirectory = Invoke-Command -Session $RemoteSession -ArgumentList $LocalInputDirectory, $RemoteTempDirectory, $InputName, $InputPath -ScriptBlock {
                    # Принимаем параметры из родительского контекста
                    Param ($LocalInputDirectory, $RemoteTempDirectory, $InputName, $InputPath);
                    # Переименовываем директорию ввода
                    if ($InputPath) {
                        $RemoteInputDirectory = Rename-Item -Path (Join-Path $RemoteTempDirectory.FullName $LocalInputDirectory.Name) -NewName $InputName -PassThru;
                    }
                    else {
                        $RemoteInputDirectory = New-Item -ItemType "Directory" -Path (Join-Path $RemoteTempDirectory.FullName $InputName);
                    };
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
                    # Переименовываем директорию ввода
                    $RemoteOutputDirectory = New-Item -ItemType "Directory" -Path (Join-Path $RemoteTempDirectory.FullName $OutputName);
                    # Возвращаем результат
                    $RemoteOutputDirectory;
                };
                if (-not($RemoteOutputDirectory)) { $CustomError = 6; };
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
                    $InvokeResult = @{"TaskResult" = $TaskResult; "IsCommandRun" = $IsCommandRun };
                    $InvokeResult;
                };
                $IsCommandRun = $InvokeResult.IsCommandRun;
                if ($IsCommandRun) {
                    $TaskResult = $InvokeResult.TaskResult;
                    if (-not($TaskResult)) { $SuccessRunCount++; };
                };
            };
            # Копируем директорию вывода с удалённого компьютера
            if (-not($CustomError) -and $IsСheckPass -and $OutputPath -and $IsCommandRun) {
                # Создаём временную директорию для получения результата
                $TempPath = Join-Path ([System.IO.Path]::GetTempPath()) ([System.IO.Path]::GetRandomFileName());
                if (Test-Path -Path $TempPath) { Remove-Item -Path $TempPath -Force -Recurse; };
                $LocalTempDirectory = New-Item -ItemType "Directory" -Path $TempPath;
                # Копируем полученный результат с удалённого компьютера через временную папку
                Copy-Item -Path $RemoteOutputDirectory.FullName -FromSession $RemoteSession -Destination $LocalTempDirectory.FullName -Recurse;
                foreach ($Directory in (Get-ChildItem -Path $LocalTempDirectory.FullName)) {
                    foreach ($Item in (Get-ChildItem -Path $Directory.FullName)) {
                        $ItemPath = Join-Path $LocalOutputDirectory.FullName $Item.Name;
                        if (Test-Path -Path $ItemPath) { Remove-Item -Path $ItemPath -Force -Recurse; };
                        Copy-Item -Path $Item.FullName -Destination $ItemPath -Force;
                    };
                };
            };
            # Удаляем временную директорию на локальном компьютер
            if ($LocalTempDirectory) {
                Remove-Item -Path $LocalTempDirectory.FullName -Force -Recurse;
            };
            # Удаляем временную директорию на удалённый компьютер
            if ($RemoteTempDirectory) {
                Invoke-Command -Session $RemoteSession -ScriptBlock {
                    Remove-Item -Path $RemoteTempDirectory.FullName -Force -Recurse;
                };
            };
            # Завершаем сессию на удалённом компьютере
            if ($RemoteSession) {
                Remove-PSSession -Session $RemoteSession;
            };
            # Формируем результ
            if (-not($OutputPath) -and $IsСheckPass) {
                $Result += [PSCustomObject]@{
                    "ComputerName" = $ComputerNameItem;
                    "TaskResult"   = $TaskResult;
                };
            };
            # Сбрасываем внутренную ошибка выполнения скрипта
            $CustomError = 0; 
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
