# Excel - save, open file

Пример сохранения и открытия файла .xlsx на двух библиотеках

### В примере используются:
#### ClosedXML.Excel - открытие, сохранение.
#### LinqToExcel - открытие.

### [ClosedXML.Excel](https://github.com/odi1n/Excel-Save-Open/tree/master#closedxmlexcel---%D0%BE%D1%82%D0%BA%D1%80%D1%8B%D1%82%D0%B8%D0%B5-%D1%81%D0%BE%D1%85%D1%80%D0%B0%D0%BD%D0%B5%D0%BD%D0%B8%D0%B5)
Имеется:
* [Сохранение файла](https://github.com/odi1n/Excel-Save-Open/blob/d7499043fd6225d0752b5d91bdf0c29261b4589a/Test%20Excel/Program.cs#L157)
* [Открытие файла](https://github.com/odi1n/Excel-Save-Open/blob/d7499043fd6225d0752b5d91bdf0c29261b4589a/Test%20Excel/Program.cs#L101)

Очень удобная в плане работы с самим Excel файлом, можно указывать как формулы так и все остальное. Много вариантов работы с формированием файла.
Плоха тем что при сохранение, открытие файла идет большая нагрузка.

### [LinqToExcel](https://github.com/odi1n/Excel-Save-Open/tree/master#linqtoexcel---%D0%BE%D1%82%D0%BA%D1%80%D1%8B%D1%82%D0%B8%D0%B5)
Имеется:
* [Открытие файла](https://github.com/odi1n/Excel-Save-Open/blob/d7499043fd6225d0752b5d91bdf0c29261b4589a/Test%20Excel/Program.cs#L79)

Хороша тем что можно сразу открыть в модель. Нет нагрузки при открытие файла


Имеется несколько вариантов открытие и сохранения данных.
