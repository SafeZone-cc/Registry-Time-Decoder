# Registry time decoder

Консольный VBS скрипт.

Позволяет увидеть в привычном формате DD.MM.YYYY hh:mm:ss даты, которые указаны в реестре в бинарном формате или в виде 16-ричного значения.

Можно указывать на выбор:

* бинарную строку
* 16-ричное число
* путь к параметру реестра

Поддерживаемые форматы:
* Unix-Time (4 байта)
* FILETIME (8 байт)
* SYSTEMTIME (16 байт)

Фейс:

![regdateconv](https://user-images.githubusercontent.com/19956568/42976492-7cbdb674-8bca-11e8-9e7f-1040b71d8d17.png)

**Примечание**

Скрипт поддерживает задание конвертируемой строки через аргументы, например:
`cscript путь\RegTimeDecoder.vbs 00,80,8c,a3,c5,94,c6,01`
Скрипт может попросить повышения привилегий, если ему их не хватит, чтобы прочитать параметр реестра.

**Буфер обмена**

Чтобы вставить содержимое буфера обмена в окно консоли CSCRIPT, нажмите правой кнопкой мыши по заголовку окна => Свойства => Общие => Поставьте галочку на "Выделение мышью" => ОК.
Теперь вы сможете вставлять буфер правой кнопкой мышки.

![clip](https://user-images.githubusercontent.com/19956568/42976495-807ee0ee-8bca-11e8-84ca-74481fb7cf25.png)
