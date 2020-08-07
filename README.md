# Excel - save, open file .xlsx

Пример сохранения/открытия файла .xlsx на нескольких библиотеках.

## EPPLus
* [Открытие файла](https://github.com/odi1n/Excel-Save-Open/blob/ab0eef01856e886f4f43f2e45c7ee32888bdb693/Test%20Excel/Program.cs#L165), [class импорт excel в list](https://github.com/odi1n/Excel-Save-Open/blob/ab0eef01856e886f4f43f2e45c7ee32888bdb693/Test%20Excel/Program.cs#L199)
* [Сохранение файла](https://github.com/odi1n/Excel-Save-Open/blob/ab0eef01856e886f4f43f2e45c7ee32888bdb693/Test%20Excel/Program.cs#L177)

## LinqToExcel
* [Открытие файла](https://github.com/odi1n/Excel-Save-Open/blob/ab0eef01856e886f4f43f2e45c7ee32888bdb693/Test%20Excel/Program.cs#L54)

## ClosedXML.Excel
* [Открытие файла](https://github.com/odi1n/Excel-Save-Open/blob/ab0eef01856e886f4f43f2e45c7ee32888bdb693/Test%20Excel/Program.cs#L70)
* [Открытие файла 2](https://github.com/odi1n/Excel-Save-Open/blob/ab0eef01856e886f4f43f2e45c7ee32888bdb693/Test%20Excel/Program.cs#L95)
* [Сохранение файла](https://github.com/odi1n/Excel-Save-Open/blob/ab0eef01856e886f4f43f2e45c7ee32888bdb693/Test%20Excel/Program.cs#L120)
* [Сохранение файла 2](https://github.com/odi1n/Excel-Save-Open/blob/ab0eef01856e886f4f43f2e45c7ee32888bdb693/Test%20Excel/Program.cs#L134)

## Выводы
Мои цели:
1. Сохранить linq в excel.
2. Открыть excel в linq.
3. Количество строк 300к+.
4. Не было нагрузки выше 200-500мб.

По результатам личных наблюдений и тестов я расположил их в следующем порядке:

|Место|Библиотека|
|:---:|:---:|
|1|EPPLus|
|2|LinqToExcel|
|3|ClosedXML.Excel|

1. EPPLus - по моему мнению оказалась лучшей. Хорошо работает с памятью грузило примерно 200-400мб во время работы, после того как открыла/сохранила файл очищает память и приложение где-то занимает 40-100мб. Имеется много вариантов работы с форматированием файла.
2. LinqToExcel -  имеет только возможность загрузки файла, сохранения файла тут нет, понравилась тем что при открытии память особо и не занимает, было примерно 120-150мб во время открытия, после 40-100мб.
3. ClosedXML.Excel - единственный плюс в ней как по мне это то что имеется большое количество форматирования файла, во время открытия/сохранения файла, память долетала до 1-1.5к гб, после чего так и оставалсь, то есть не было очистики
