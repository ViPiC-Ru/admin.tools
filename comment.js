/* 0.1.3 копирует первый комментарий в файл

cscript comment.js <source> <destination> <charset> [<short>]

<source>        - путь для исходного текстового файла
<destination>   - путь для конечного текстового файла
<charset>       - кодировка обоих текстовых файлов 
<short>         - преобразовать в короткий комментарий

*/

(function (wsh, undefined) {// замыкаем что бы не сорить глобалы
    var fso, srcStream, dstStream, binStream, value, source, destination,
        iMB, iME, iSB, iSE, dMB = '/*', dME = '*/', dSB = '//', dSR = '\r',
        dSN = '\n', short = null, error = 0;

    // создаём необходимые объекты
    if (!error) {// если нету ошибок
        fso = new ActiveXObject('Scripting.FileSystemObject');
        srcStream = new ActiveXObject('ADODB.Stream');
        dstStream = new ActiveXObject('ADODB.Stream');
        binStream = new ActiveXObject('ADODB.Stream');
    };
    // получаем путь для исходного файла
    if (!error) {// если нету ошибок
        if (0 < wsh.arguments.length) {// если передан параметр
            value = wsh.arguments(0);
            source = fso.getAbsolutePathName(value);
        } else error = 1;
    };
    // получаем путь для конечного файла
    if (!error) {// если нету ошибок
        if (1 < wsh.arguments.length) {// если передан параметр
            value = wsh.arguments(1);
            destination = fso.getAbsolutePathName(value);
        } else error = 2;
    };
    // получаем кодировку обоих файлов
    if (!error) {// если нету ошибок
        if (2 < wsh.arguments.length) {// если передан параметр
            value = wsh.arguments(2);
            charset = value.toLowerCase();
        } else error = 3;
    };
    // получаем необходимость преобразования комментария
    if (!error) {// если нету ошибок
        if (3 < wsh.arguments.length) {// если передан параметр
            value = wsh.arguments(3);
            value = value.toLowerCase();
            short = 'true' == value;
        };
    };
    // инициализируем необходимые объекты
    if (!error) {// если нету ошибок
        try {// пробуем задать кодировки
            // входящий поток данных
            srcStream.type = 2;// adTypeText
            srcStream.mode = 3;// adModeReadWrite
            srcStream.charset = charset;
            srcStream.open();
            // исходящий поток данных
            dstStream.type = 2;// adTypeText
            dstStream.mode = 3;// adModeReadWrite
            dstStream.charset = charset;
            dstStream.open();
            // бинарный поток данных
            binStream.type = 1;// adTypeBinary
            binStream.mode = 3;// adModeReadWrite
            binStream.open();
        } catch (e) { error = 4; };
    };
    // читаем данные из исходного файла
    if (!error) {// если нету ошибок
        try {// пробуем прочитать данные
            srcStream.loadFromFile(source);
            value = srcStream.readText();
        } catch (e) { error = 5; };
    };
    // получаем комментарий из данных
    if (!error) {// если нету ошибок
        iMB = value.indexOf(dMB); // index Multiline Begin
        iME = ~iMB ? value.indexOf(dME, iMB + dMB.length) : -1;
        iSB = value.indexOf(dSB); // index Singleline Begin
        iSE = ~iSB ? value.indexOf(dSR + dSN, iSB + dSB.length) : -1;
        if (~iSB && !~iSE) iSE = value.indexOf(dSN, iSB + dSB.length);
        if (~iSB && !~iSE) iSE = value.indexOf(dSR, iSB + dSB.length);
        if (~iSB && !~iSE) iSE = value.length - iSB;
        if (~iMB && ~iME && (!~iSE || iSB > iMB)) {// если многострочный
            if (short) {// если необходимо преобразовать в короткий комментарий
                value = value.substr(iMB + dMB.length, iME - iMB - dMB.length);
                iSE = value.indexOf(dSR + dSN);// ищем конец для короткого комментария
                if (!~iSE) iSE = value.indexOf(dSN);// ищем первый альтернативный конц
                if (!~iSE) iSE = value.indexOf(dSR);// ищем второй альтернативный конц
                if (!~iSE) iSE = value.length;// выбираем конец всей строки
                value = dSB + value.substr(0, iSE);// сокращаем длину комментария
                dstStream.writeText(value, 1);// adWriteLine
            } else {// если преобразования не требуются
                value = value.substr(iMB, iME - iMB + dME.length);
                dstStream.writeText(value, 1);// adWriteLine
                dstStream.writeText('', 1);// adWriteLine
            };
        } else if (~iSB && ~iSE && (!~iME || iMB > iSB)) {// если однострочный
            value = value.substr(iSB, iSE - iSB);
            dstStream.writeText(value, 1);// adWriteLine
        } else error = 6;
    };
    // читаем данные из конечного файла
    if (!error) {// если нету ошибок
        try {// пробуем прочитать данные
            srcStream.loadFromFile(destination);
            srcStream.copyTo(dstStream);
        } catch (e) { error = 7; };
    };
    // записываем данные в конечный файл
    if (!error) {// если нету ошибок
        try {// пробуем записать данные в файл
            srcStream.close();// закрываем входящий поток данных
            if ('utf-8' == charset) {// если нужно убрать BOM в utf-8
                dstStream.position = 3;// пропускаем BOM
                dstStream.copyTo(binStream);
                binStream.saveToFile(destination, 2);// adSaveCreateOverWrite
            } else dstStream.saveToFile(destination, 2);// adSaveCreateOverWrite
            dstStream.close();// закрываем исходящий поток данных
            binStream.close();// закрываем бинарный поток данных
        } catch (e) { error = 8; };
    };
    // завершаем сценарий кодом
    wsh.quit(error);
})(WSH);