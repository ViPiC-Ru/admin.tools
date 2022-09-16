<#
  .SYNOPSIS
  Получает объект с данными в сертификате.

  .DESCRIPTION
  Скрипт получает данные из сертификата и возвращает их в виде объекта.

  .PARAMETER Certificate
  Список сертификатор из которых нужно получить данные.

  .PARAMETER DataKey
  Ключ для возврата одного значения из объекта с данными
  или для добавления первого ключа.

  .PARAMETER DataValue
  Значение для добавления первого ключа.

  .PARAMETER Expanded
  Возвращать расширенные свойства из сертификата.

  .PARAMETER FixExpire
  Исправлять дату окончания сертификата для долгих сертификатов.

  .INPUTS
  Вы можете передавать в скрипт список сертификатов по коневееру.

  .OUTPUTS
  Возвращает список объектов с данными по каждому сертификату или
  значение для заданного ключа для последнего сертификата.

  .NOTES
  Версия: 0.1.3
  Автор: @ViPiC
#>

[CmdletBinding()]
Param (
    [Parameter (Mandatory = $true, ValueFromPipeline = $true)]
    [array]$Certificate,

    [string]$DataKey,

    [string]$DataValue,

    [switch]$Expanded,

    [switch]$FixExpire
);

PROCESS {

    function ConvertTo-HashTable {
        # Конвертирует строки в хеш таблицу
        Param (
            [String]$StringData,
            $ExtensionData
        )

        $HashTable = @{};
        if ($StringData) {
            $DelimNewLine = [Environment]::NewLine;
            $StringLines = $StringData.Split($DelimNewLine);
            foreach ($StringLine in $StringLines) {
                $StringLineData = ConvertFrom-StringData -StringData $StringLine;
                foreach ($Key in $StringLineData.Keys) {
                    $Value = $StringLineData[$Key];
                    if ($Value.Length -ge 2 -and $Value[0] -eq '"' -and $Value[$Value.Length - 1] -eq '"') {
                        $Value = $Value.Substring(1, $Value.Length - 2);
                    };
                    if ($HashTable[$Key]) {
                        $HashTable[$Key] = switch ($Key) {
                            "DC" { $Value + "." + $HashTable[$Key] }
                            "OU" { $HashTable[$Key] + "\" + $Value }
                            default { $HashTable[$Key] + " " + $Value }
                        };
                    }
                    else {
                        $HashTable[$Key] = $Value;
                    };
                };
            };
        }
        elseif ($ExtensionData) {
            foreach ($Extension in $ExtensionData) {
                $Value = $Extension.Format($true);
                if (-not($Value)) { $Value = $true; };
                $HashTable[$Extension.Oid.Value] = $Value;
            };
        };
        return $HashTable;
    };

    function Repair-Name {
        # Исправляет имена в значениях
        Param (
            [String]$Name = "",
            [Switch]$Expanded
        )

        $Name = $Name.Replace("Общество с ограниченной ответственностью", "ООО");
        $Name = $Name.Replace("ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ", "ООО");
        $Name = $Name.Replace("_", " ");
        $Name = $Name.Replace('"', "");
        if ($Expanded) {
            $Name = $Name.Replace(",", "");
            $Name = $Name.Replace("*.", "");
            $Name = $Name.TrimStart("0");
        };
        return $Name;
    };

    function Repair-NotEmpty {
        # Исправляет и возвращает не пустое значение
        Param (
            $Object,
            $Replacement
        )

        $Value = $null;
        foreach ($Item in $Replacement.GetEnumerator()) {
            if ($Object[$Item.Name]) { $Value = $Item.Value; };
        }
        return $Value;
    };

    function Repair-Date {
        # Исправляет дату если она привышает максимальную
        Param (
            $BeforeDate,
            $AfterDate,
            [Int]$MaxDays = 0
        )
        if ($BeforeDate -and $AfterDate -and $MaxDays) {
            if (($AfterDate - $BeforeDate).Days -gt $MaxDays) {
                $AfterDate = $BeforeDate.AddYears(1);
            };
        };
        return $AfterDate;
    };

    function Join-Items {
        # Объединяет элименты массива
        Param (
            [Array]$Items = @(),
            [Switch]$NoEmpty
        )

        $Delim = " ";
        $Result = "";
        
        foreach ($Item in $Items) {
            $IsNeedAdd = (-not $NoEmpty) -or $Item;
            if ($IsNeedAdd) {
                if ($Result) {
                    $Result += $Delim + $Item;
                }
                else {
                    $Result += $Item;
                };
            };
        };
        return $Result;
    };

    function Get-FirstItem {
        # Получает первый элимент
        Param (
            [Array]$Items = @(),
            [Switch]$NoEmpty
        )

        $First = $null;
        $IsFound = $false;
        foreach ($Item in $Items) {
            if (-not $IsFound) {
                $IsFound = (-not $NoEmpty) -or $Item;
                if ($IsFound) {
                    $First = $Item;
                };
            };
        };
        return $First;
    };

    #################################################################################################################

    $Result = @();
    # Последовательно обрабатываем переданные сертификаты
    foreach ($CertificateItem in $Certificate) {
        if ($FixExpire) { $MaxExpireDays = 500; } else { $MaxExpireDays = 0; };
        $IssuerData = ConvertTo-HashTable -StringData $CertificateItem.IssuerName.Format($true);
        $SubjectData = ConvertTo-HashTable -StringData $CertificateItem.SubjectName.Format($true);
        $ExtensionData = ConvertTo-HashTable -ExtensionData $CertificateItem.Extensions;
        $CertificateData = [PSCustomObject]@{};
        # Информация об сертификате
        if ($DataKey -and $DataValue) {
            Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name $DataKey -Value $DataValue;
        };
        Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name "CER-THUMBPRINT" -Value $CertificateItem.Thumbprint;
        Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name "CER-CREATE" -Value $CertificateItem.NotBefore.ToString("dd.MM.yyyy HH:mm:ss");
        Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name "CER-EXPIRE" -Value (Repair-Date -BeforeDate $CertificateItem.NotBefore -AfterDate $CertificateItem.NotAfter -MaxDays $MaxExpireDays).ToString("dd.MM.yyyy HH:mm:ss");
        if ($Expanded) {
            Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name "CER-NAME" -Value ((Repair-Date -BeforeDate $CertificateItem.NotBefore -AfterDate $CertificateItem.NotAfter -MaxDays $MaxExpireDays).ToString("yyyy.MM.dd") + " #" + $CertificateItem.Thumbprint.Substring($CertificateItem.Thumbprint.Length - 4) + " - " + (Get-FirstItem (Repair-Name (Join-Items $SubjectData["SN"], $SubjectData["G"] -NoEmpty)), (Repair-Name $SubjectData["CN"] -Expanded) -NoEmpty));
        };
        # Информация об издателе
        Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name "CER-ISSUER" -Value (Repair-Name $IssuerData["CN"]);
        Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name "CER-ISSUER-DOMAIN" -Value $IssuerData["DC"];
        Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name "CER-ISSUER-DEPARTMENT" -Value $IssuerData["OU"];
        Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name "CER-ISSUER-LOCALITY" -Value $IssuerData["L"];
        Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name "CER-ISSUER-EMAIL" -Value $IssuerData["E"];
        Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name "CER-ISSUER-INN" -Value (Repair-Name $IssuerData["ИНН"] -Expanded);
        Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name "CER-ISSUER-OGRN" -Value (Repair-Name $IssuerData["ОГРН"] -Expanded);
        if ($Expanded) {
            Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name "CER-ISSUER-STATE" -Value $IssuerData["S"];
            Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name "CER-ISSUER-COUNTRY" -Value $IssuerData["C"];
        };
        # Информация об субъекте
        Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name "CER-SUBJECT" -Value (Repair-Name $SubjectData["CN"]);
        Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name "CER-SUBJECT-DOMAIN" -Value $SubjectData["DC"];
        Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name "CER-SUBJECT-DEPARTMENT" -Value $SubjectData["OU"];
        Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name "CER-SUBJECT-LOCALITY" -Value $SubjectData["L"];
        Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name "CER-SUBJECT-EMAIL" -Value $SubjectData["E"];
        Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name "CER-SUBJECT-EMPLOYEE" -Value (Repair-Name (Join-Items $SubjectData["SN"], $SubjectData["G"] -NoEmpty));
        Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name "CER-SUBJECT-TITLE" -Value $SubjectData["T"];
        Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name "CER-SUBJECT-INN" -Value (Repair-Name $SubjectData["ИНН ЮЛ"] -Expanded);
        Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name "CER-SUBJECT-OGRN" -Value (Repair-Name $SubjectData["ОГРН"] -Expanded);
        if ($Expanded) {
            Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name "CER-SUBJECT-SNILS" -Value (Repair-Name $SubjectData["СНИЛС"] -Expanded);
            Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name "CER-SUBJECT-STATE" -Value $SubjectData["S"];
            Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name "CER-SUBJECT-COUNTRY" -Value $SubjectData["C"];
        };
        # Информация об встроенных лицензиях
        Add-Member -InputObject $CertificateData -MemberType "NoteProperty" -Name "CER-EMBEDDED-LICENSE" -Value (Repair-NotEmpty $ExtensionData @{"1.2.643.2.2.49.2" = "КриптоПро CSP" });
        # Добавляем в результат
        if ((-not $DataValue) -and $DataKey) {
            $Result = $CertificateData.$DataKey;
        }
        else {
            $Result += $CertificateData;
        };
    };
    # Возвращаем результат
    return $Result;
};