# Calendar_Sheet

Пишу скрипт для гугл таблицы чтобы вести учёт тренировок в спортивном зале

1) События созданные в гугул календаре - по желанию пользователя загружаются в гугл таблицу.
2) События сортируются по времени (час) и распределяются на колонки. 
3) У каждого события есть слова маркеры, которые помогают вести отдельный учет их количества.
4) Каждый день события суммируются и выводятся в отдельное поле.
5) Если подлючена таблица администратора с учётом посещаемости клиентов, то скрипт сверяется и данными обеих таблиц и в случае отличия значений - окрашивает ячейку в красный цвет.
6) По завершении месяца скрипт формирует новую строку желтого цвета и вписывает в её ячейки формулы суммирующие столбцы с ежедневными значениями.
7) Когда суммирующий скрипт заканчивается - он спрашивает пользователя о том необходимо ли перенести значения в соседний лист "Аналитический).
8) Данные за месяц переносятся в отдельную таблицу.
