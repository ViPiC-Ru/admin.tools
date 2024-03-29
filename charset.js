/* 0.1.5 изменяет кодировку текстовых файлов

cscript charset.js <action> <source> [[... <source>] <destination>] <input> <output>

<action>        - действие над исходным файлом
    copy        - скопировать исходный файл
    move        - переместить исходный файл
<source>        - путь для исходного текстового файла
<destination>   - путь для конечного текстового файла
<input>         - кодировка исходного файла
<output>        - кодировка конечного файла

*/

(function (wsh, undefined) {// замыкаем что бы не сорить глобалы
    var fso, srcStream, dstStream, binStream, value, index, action,
        source, destination, input, output, doMove, flag, error = 0;

    // создаём необходимые объекты
    if (!error) {// если нету ошибок
        fso = new ActiveXObject('Scripting.FileSystemObject');
        srcStream = new ActiveXObject('ADODB.Stream');
        dstStream = new ActiveXObject('ADODB.Stream');
        binStream = new ActiveXObject('ADODB.Stream');
    };
    // получаем действие над файлом
    if (!error) {// если нету ошибок
        if (0 < wsh.arguments.length) {// если передан параметр
            action = wsh.arguments(0).toLowerCase();
            switch (action) {// поддерживаемые действия
                case 'copy': doMove = false; break;
                case 'move': doMove = true; break;
                default: error = 2;
            };
        } else error = 1;
    };
    // получаем путь для конечного файла
    if (!error) {// если нету ошибок
        if (3 < wsh.arguments.length) {// если передан параметр
            index = wsh.arguments.length - 3;
            value = wsh.arguments(index);
            destination = fso.getAbsolutePathName(value);
        } else error = 2;
    };
    // получаем кодировку исходного файла
    if (!error) {// если нету ошибок
        if (3 < wsh.arguments.length) {// если передан параметр
            index = wsh.arguments.length - 2;
            input = wsh.arguments(index).toLowerCase();
        } else error = 3;
    };
    // получаем кодировку конечного файла
    if (!error) {// если нету ошибок
        if (3 < wsh.arguments.length) {// если передан параметр
            index = wsh.arguments.length - 1;
            output = wsh.arguments(index).toLowerCase();
        } else error = 4;
    };
    // инициализируем необходимые объекты
    if (!error) {// если нету ошибок
        try {// пробуем задать кодировки
            // входящий поток данных
            srcStream.type = 2;// adTypeText
            srcStream.mode = 3;// adModeReadWrite
            srcStream.charset = input;
            srcStream.open();
            // исходящий поток данных
            dstStream.type = 2;// adTypeText
            dstStream.mode = 3;// adModeReadWrite
            dstStream.charset = output;
            dstStream.open();
            // бинарный поток данных
            binStream.type = 1;// adTypeBinary
            binStream.mode = 3;// adModeReadWrite
            binStream.open();
        } catch (e) { error = 5; };
    };
    // обрабатываем последовательность исходных файлов
    for (var i = 0, iLen = wsh.arguments.length; !error && (!i || i < iLen - 4); i++) {
        value = wsh.arguments(i + 1);// получаем очередное знчение
        source = fso.getAbsolutePathName(value);
        // читаем и конвертируем данные из файла
        if (!error) {// если нету ошибок
            try {// пробуем прочитать и сконвертировать данные
                srcStream.loadFromFile(source);
                if (i) dstStream.writeText('', 1);// adWriteLine
                srcStream.copyTo(dstStream);
            } catch (e) { error = 6; };
        };
    };
    // записываем данные в конечный файл
    if (!error) {// если нету ошибок
        try {// пробуем записать данные в файл
            srcStream.close();// закрываем входящий поток данных
            if ('utf-8' == output) {// если нужно убрать BOM в utf-8
                dstStream.position = 3;// пропускаем BOM
                dstStream.copyTo(binStream);
                binStream.saveToFile(destination, 2);// adSaveCreateOverWrite
            } else dstStream.saveToFile(destination, 2);// adSaveCreateOverWrite
            dstStream.close();// закрываем исходящий поток данных
            binStream.close();// закрываем бинарный поток данных
        } catch (e) { error = 7; };
    };
    // обрабатываем последовательность исходных файлов
    for (var i = 0, iLen = wsh.arguments.length; !error && doMove && i < iLen - 4; i++) {
        value = wsh.arguments(i + 1);// получаем очередное знчение
        source = fso.getAbsolutePathName(value);
        flag = source.toLowerCase() != destination.toLowerCase();
        // удаляем исходные файлы при необходимости
        if (!error && flag && fso.fileExists(source)) {// если файл существует
            try {// пробуем сделать операцию над файлом
                fso.deleteFile(source, true);
            } catch (e) { error = 8; };
        };
    };
    // завершаем сценарий кодом
    wsh.quit(error);
})(WSH);