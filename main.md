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




          
