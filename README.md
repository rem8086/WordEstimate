WordEstimate
============

WordEstimate
Программа служит для вывода данных из набора файлов смет в формате Word в excel таблицы.
Программа запускается из командной строки
Первый параметр - входная директория 
Второй параметр - директория для выходного файла
Выходной файл содержит три заполненных листа
1. Общие данные по каждой смете. Выводится по строке с итоговой информацией по каждой смете: название, общие стоимости по ЭМ, ОЗП, МАТ, ЗПМ, а также наклдадные расходы и сметную прибыль
2. Данные по всем строкам всех входящих смет, с объемами и ценами
3. Данные по всем строкам материалов смет (строки сметы с нулевыми ценами на все, кроме М)
Так как формат входящих смет может различаться (существует несколько форм word для представления сметных данных), в программе присутствует конфигурационный файл с набором параметров и файл excel шаблонов.

Конфигурационный файл содержит:
порядковые номера таблиц, из которых берутся данные
название файла с excel шаблонами и названия листов отдельных шаблонов
шаблоны названий ключевых полей, которые служат якорями и на основании которых выбираются данные 

Файл шаблонов состоит из четырех листов, на каждом из которых находится именованный диапазон, в соответствии с которым происходит поиск данных.
1. String. Блок, по которому ищутся отдельные строки внутри сметы.
Имеет ключевое поле Number, которое ищется в смете в соответствии с шаблоном StringNumberPattern из конфигурационного файла. После нахождения в файле сметы поля Number из сметы считываются все остальные именованные поля из соответствующих ячеек: Name, Volume, Cost, Pay, Machine, MachinePay. Помимо этого, если в смете присутутсвует блок накладных расходов/сметной прибыли для некоторых строк, то вводится поле StringCondition, вычислемое по шаблону StringConditionPattern, и при совпадении ячейки сметы с шаблоном, также  заполняются все значения после этого поля (Overheads, Profit, TotalCost).
2. Resume. Блок строки итогов сметы. Ключевое поле для поиска строки Name, сверяется со значением параметра ResumeStringPattern. При совпадении заполняются значения из ячеек (Volume, Cost, Pay, Machine, MachinePay).
3. Cost. Прочие строки, информация из которых используется (такие как "Накладные расходы", "Стоимость оборудования" и т.д.). Ключевое поле Name, из файла сметы считывается стоимость строки (Cost).
4. Header. Таблица заголовка сметы. Из данного блока выбирается информация о коде сметы и ее названии (Number, Name). Так как блок заголовка единственный и находится в начале сметы, то в качестве ключевого поля используется FirstCell, поиск котороого происходит не по шаблону, а вместо этого указывается первая ячейка таблицы.