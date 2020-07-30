# Excel - save, open file .xlsx

Пример сохранения и открытия файла .xlsx на двух библиотеках

### В примере используются:
#### [ClosedXML.Excel](https://github.com/odi1n/Excel-Save-Open#closedxmlexcel) - открытие, сохранение.
#### [LinqToExcel](https://github.com/odi1n/Excel-Save-Open#closedxmlexcel) - открытие.

### ClosedXML.Excel
Имеется:
* [Сохранение файла](https://github.com/odi1n/Excel-Save-Open/blob/d7499043fd6225d0752b5d91bdf0c29261b4589a/Test%20Excel/Program.cs#L157)
* [Открытие файла](https://github.com/odi1n/Excel-Save-Open/blob/d7499043fd6225d0752b5d91bdf0c29261b4589a/Test%20Excel/Program.cs#L101)

Очень удобная в плане работы с самим Excel файлом, можно указывать как формулы так и все остальное. Много вариантов работы с формированием файла.
Плоха тем что при сохранение, открытие файла идет большая нагрузка.

### LinqToExcel
Имеется:
* [Открытие файла](https://github.com/odi1n/Excel-Save-Open/blob/d7499043fd6225d0752b5d91bdf0c29261b4589a/Test%20Excel/Program.cs#L79)

Хороша тем что можно сразу открыть в модель. Нет нагрузки при открытие файла

Имеется по несколько вариантов открытия и сохранения файлов.
