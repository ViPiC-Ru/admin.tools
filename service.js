/* 1.0.0 служба для постоянного периодического выполнения переданных команд

сборка: %WINDIR%\Microsoft.NET\Framework64\%VERSION%\jsc.exe /target:winexe service.js
создание: New-Service -Name "ID" -DisplayName "Имя" -Description "Описание" -BinaryPathName "service.exe [<timeout>] ""<command>"" <command> ..."
удаление: sc.exe delete "ID"

*/

import System;
import System.IO;
import System.ServiceProcess;
import System.Threading;

class MyService extends ServiceBase {
    private var thread: Thread;
    protected override function OnStart(args: String[]) {
        thread = new Thread(ThreadStart(MyThread.Work));
        thread.Start();
    }
    protected override function OnStop() {
        thread.Abort();
    }
}

class MyThread {
    static function Work() {
        var arg, args, shell, command, commands = [], workDirectory = "", timeout = 60;

        args = Environment.GetCommandLineArgs();
        shell = new ActiveXObject("WScript.Shell");
        // получаем параметры из командной строки
        for (var i = 0, iLen = args.length; i < iLen; i++) {
            arg = args[i];
            switch (i) {
                case 0:
                    workDirectory = Path.GetDirectoryName(arg);
                    break;
                case 1:
                    if (!isNaN(arg)) {
                        timeout = Math.max(arg, 0);
                        break;
                    };
                default:
                    command = arg.split("'").join('"');
                    commands.push(command);
                    break;
            };
        };
        // готовим окружение и бесконечно выполняем команды
        Directory.SetCurrentDirectory(workDirectory);
        while (commands.length) {
            for (var i = 0, iLen = commands.length; i < iLen; i++) {
                command = commands[i];
                try {// пробуем выполниить
                    shell.Run(command, 0, true);
                } catch (error) { };
            };
            Thread.Sleep(timeout * 1000);
        };
    }
}

ServiceBase.Run(new MyService());