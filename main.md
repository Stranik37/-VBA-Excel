<h1>Шпаргалка по VBA excel</h1>
<h2>Начало работы</h2>
1. Чтобы войти в VBA нужно либо найти Visual Basic во вкладке разработчика, либо нажать сочетание ALT + F11
2. Нажать 2 раза по excel файлу -> insert -> module
3. Чтобы начать зааписывать макрос, нам нужно:
  3.1 Sub и End Sub — это ключевые слова в языке программирования VBA (Visual Basic for Applications), которые обозначают начало и конец макроса.
  Пример: 
        Sub testSub()
              MsgBox "Hello world!"
        End Sub
  3.2 testSub() - начало макроса
  3.3 MsgBox - команда, чтобы вывести сообщение
4.Чтобы зпустить макрос нудно нжать на Run Sub
<h2>Переменные</h2>
1. Типы переменных:
  1.1 Basic Variable - содержит одно знчение заданного наами типа днных (1,а,?)
  1.2 Variant - содержит одно значение. Тип данных автоматически определяется VBA.
  1.3 Object Variable - содержит в себе один объект заданного нами типа
  1.4 Array - содержит в себе множество значений или объектов
2. Объявление переменных:
  2.1 Чтобы объявить переменную нужно написать команду Dim и задать формат через AS.
  Например:
          Sub testSub()
              'BASIC VARIABLE
               Dim someText AS String
               Dim someNumber AS Long
               Dim someDate AS Date

              'OBJECT VARIABLE
              Dim someWorkbook AS Workbook
              Dim someWorksheet AS Worksheet
              Dim someRange AS Range
          End Sub
  2.2 ' - это апостраф, с помощью него можно сделать коментарии.
  2.3 Чтобы объявить переменную нужно: имя пересменной = значение переменной
  Например: someText = "Текстовая информация"; someNumber = 100; someDate = "19.01.2022"
  2.4 Чтобы объявить переменную, которая ссылается на объект, нужно задать ключевое слово Set
  Оператор Set в VBA (Visual Basic для приложений) — это ключевое слово, которое назначает ссылку на объект переменной или свойству.
  Например:
          Set someWorkbook = ThisWorkbook
          Set someWorksheet = Worksheets("Лист1")
          Set someRange = Range("C3")
<h2>Типы данных в VBA</h2>
исловые типы

- Byte — 1 байт, целые числа от 0 до 255
- Integer — 2 байта, целые числа от -32,768 до 32,767
- Long — 4 байта, целые числа от -2,147,483,648 до 2,147,483,647
- LongLong — 8 байт, целые числа (только для 64-битных версий VBA)
- Single — 4 байта, числа с плавающей точкой одинарной точности
- Double — 8 байт, числа с плавающей точкой двойной точности
- Currency — 8 байт, фиксированная десятичная точка (до 4 знаков после запятой), подходит для финансовых расчётов

Логический тип

- Boolean — 2 байта, принимает значения True или False

Строковые типы

- String — строковый тип, может быть фиксированной длины или переменной длины

Дата и время

- Date — 8 байт, хранит даты и время в формате даты VBA

Объектные типы

- Object — универсальный тип для ссылок на объекты

Специальные типы

- Variant — универсальный тип, содержит любые данные, кроме типизированных массивов; занимает минимум 16 байт
- Decimal — точный десятичный тип, используемый внутри Variant, не является отдельным типом переменной
- Array — структура данных для хранения упорядоченного набора значений одного типа

<h2>Workbooks/Worksheets/Range/Cells</h2>
1.Workbooks
1.1 Workbooks - это объект и к нему можно применять различные методы, чтобы применить метод, надо в конце поставить точку.
Workbooks("Книга1").Name - обращение к имени книги по названию
Woorkbooks(2).Name - обращения к имени книги по порядку открытия
ThisWorkbook.Name - обращение к имени книги в которой пишется макрос
ActiveWorkbook.Name - обращение к имени книги по которой последний раз кликнули мышкой
2.Worksheets
2.1 Чтобы обратиться к Worksheets, нужно сперва обратиться к Workbooks, так как в VBA есть свой порядок обращения от большего к меньшему
2.2 Это также является объектом и мы можем применять к нему различные методы и свойства с помощью точки в конце.
Workbooks(1).Worksheets("Лист1").Name - обращение к имени листа по названию листа
Workbooks(1).Worksheets(1).Name - обращение к имени листа по порядку создания листа
Workbooks(1).ActiveSheet.Name - обращение к имени листа по последнему листу, по которому мы кликнули мышкой
3.Range
3.1 Можно указывать, как 1 ячейку, так и диапазон ячеек. Также как и с остальными, можно через точку указывать методы и свойства.
Range("A1").address - обращение к адресу ячейки(по активной книге и листу)
4.Cells
4.1 Так как объет Cells находится на 1 урове с Range, то чтобы к нему обратиться, сначала нужно указать рабочую книгу и рабочий лист, а только потом указывать Cells
4.2 В cells значение ячейки указывается так: Cells(x,y), где х - номер строки; у - номер столбца
thisworkbook.worksheets(1).cells(4,3) = "Test" - написание текста в ячейку
5.Совместное использование Range и Cells
5.1 Можно указывать диапазон ячеек в Range, используя Cells
thisworkbooks.Worksheets(1).range(cells(1,1),cells(10,5)) = 2  - замена всех значение в диапазоне ячеек на цифру 2
<h2>With/Offset</h2>
1. With - нужна чтобы сокращать код
Без with:
thisworkbook.worksheets(1).range("A1") = 1
thisworkbook.worksheets(1).range("A2") = 2
thisworkbook.worksheets(1).range("A3") = 3

C with:
With thisworkbook.workssheets(1)
    .Range("A1") = 1
    .Range("A2") = 2
    .Range("A3") = 3
End with
1.1 Также можно обращаться к свойствам внутри ячеек
With thisworkbook.workssheets(1).Range("A1")
    .Font.Bold = True - делает текст жирным
    .Font.Color = vbRed - меняет цвет текста на красный
    .Value = 3 - записывает в ячейку цифру 3
End with
2.Offset - нужна чтобы сдвигать значения ячеек
Activesheet.Range("A1").Offset(1,0)
1 - Сдвиг по строке
0 - Сдвиг по столбцу
<h2>Циклы For/For each</h2>
1. Цикл For
Dim i As Long
For i = 1 To 10
    MsgBox i
Next i
Данный код выводит сообщения цифр от 1 до 10.
2. Цикл For Each
Dim cellChecked As Range
For each cellChecked in range("A1:A10")
    MsgBox cellChecked.Value
Next cellChecked
Данный код перебирает все элементы в заданном диапазоне и выводит их значения
<h2>IF</h2>
1.Sub lessonSub()
    If Range("B3") = 3 Then  - Если в ячейке B3 значение 3 тогда выводим на экран "Три"
          MsgBox "Три"
    End If
End Sub
2. С циклом
Sub lessonSub()
  Dim cellChecked As Range

  For Each cellChecked In Range("B3:B7")
      If cellChecked = 5 Then
          MsgBox "Пять"
      Else If cellChecked = 4 Then
          MsgBox "Четыре"
      Else:
          MsgBox "?"
      Next cellChecked
End Sub
<h2>Цикл Do Loop (While/Untill</h2>
1. Sub learningDoLoop_While_1()
Dim checker As Strang
checker = "ОК"
Do While checker = "OK"
    checker = InputBox("Пишите 'ОК' для повторения цикла!")  - пока checker = OK, то вводите значение, если другое, то цикл закончится
Loop
End Sub

2. Sub learningDoLoop_While_2()
Dim checker As Strang
checker = "ОК"
Do 
    checker = InputBox("Пишите 'ОК' для повторения цикла!")  - тоже самое, но выполнится хотя бы 1 рааз
Loop While checker = "OK"
End Sub

3.Sub learningDoLoop_Until_3()
Dim checker As Strang

Do Until checker = "ОК"
    checker = InputBox("Не пишите 'ОК' для повторения цикла!")  - делай до тех пор пока checker не будет равен OK
Loop
End Sub
<h2>Поиск последней используемой строки и столбца</h2>
1. Поиск последней строки
Sub identifyingLastRow()
Dim lastRow As Long
lastRow = Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row - считает все кол-во строк и ссылается на последнюю используемую
MsgBox lastRow
End Sub

2.Sub identifyingLastColumn()
Dim lastColumn As Long
lastColumn = Worksheets(1).Cells(11, Columns.Count).End(xlToLeft).Column - считает все кол-во столбцов и ссылается на последний используемый
MsgBox lastColumn
End Sub


