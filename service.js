/* 1.1.0 служба для постоянного периодического выполнения переданных команд

сборка: %WINDIR%\Microsoft.NET\Framework64\%VERSION%\jsc.exe /target:winexe service.js
создание: New-Service -Name "ID" -DisplayName "Имя" -Description "Описание" -BinaryPathName "service.exe [<timeout>] [<workdir>] ""<command>"" <command> ..."
удаление: sc.exe delete "ID"

*/

import System;
import System.IO;
import System.ServiceProcess;
import System.Threading;
import System.Diagnostics;
import System.Management;

class MyService extends ServiceBase {
    private var thread: Thread;
    private var myWorker: MyWorker;
    protected override function OnStart(args: String[]) {
        myWorker = new MyWorker();
        thread = new Thread(ThreadStart(myWorker.Run));
        thread.Start();
    }
    protected override function OnStop() {
        thread.Abort();
        myWorker.End();
    }
}

class MyWorker {
    private var workDirectory: String;
    private var process: Process;
    private var timeout: Number;
    private var commands = [];

    static function KillProcessTree(pid: Number) {
        var item, list, searcher, query;

        query = "SELECT ProcessID FROM Win32_Process WHERE ParentProcessID=" + pid;
        searcher = new ManagementObjectSearcher(query);// получаем список процессов
        for (list = new Enumerator(searcher.Get()); !list.atEnd(); list.moveNext()) {
            item = list.item();// получаем очередной элимент коллекции
            MyWorker.KillProcessTree(item["ProcessID"]);
        };
        try {// пробуем завершить процесс
            Process.GetProcessById(pid).Kill();
        } catch (error) { };
    }

    static function SplitCommand(command: String) {
        return command.match(/(?:[^\s"]|"(?:\\.|[^"\\])*")+/g);
    }

    public function MyWorker() {
        var arg, args, command;

        commands = [];// инициализируем список
        timeout = 60;// таймаут по умолчанию
        args = Environment.GetCommandLineArgs();
        // получаем параметры командной строки
        for (var i = 0, iLen = args.length; i < iLen; i++) {
            arg = args[i];// получаем параметр
            switch (i) {// поддерживаемые параметры
                case 0:// рабочая директория по умолчанию
                    workDirectory = Path.GetDirectoryName(arg);
                    break;
                case 1:// таймаут между командами
                    if (!isNaN(arg)) {
                        timeout = Math.max(arg, 0);
                        break;
                    };
                case 2:// рабочая директория
                    if (Directory.Exists(arg)) {
                        workDirectory = arg;
                        break;
                    };
                default:// команды
                    command = arg.split("'").join('"');
                    commands.push(command);
                    break;
            };
        };
    }

    public function Run() {
        var args, command;

        // готовим окружение и бесконечно выполняем команды
        Directory.SetCurrentDirectory(workDirectory);
        while (commands.length) {// если список команд не пуст
            for (var i = 0, iLen = commands.length; i < iLen; i++) {
                command = commands[i];// получаем команду
                args = MyWorker.SplitCommand(command);
                process = Process.Start(args.shift(), args);
                process.WaitForExit();
            };
            if (timeout) Thread.Sleep(timeout * 1000);
        };
    }

    public function End() {
        MyWorker.KillProcessTree(process.Id);
    }
}

ServiceBase.Run(new MyService());