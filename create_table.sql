use Test;

drop table if exists Sales;
create table Sales (
    dt date not null,
    article varchar(10) not null,
    kg numeric(6, 3) not null,
    y as year(dt) persisted, -- см. таблицу ниже
    m as month(dt) persisted
);
-- Хранимые расчетные поля повышают производительность запроса,
-- несмотря на повышение logical reads:
-- --------------------------------------------
--           | logical reads | elapsed time, ms
-- --------------------------------------------
-- без полей |          1558 |              83 
-- с полями  |          1069 |             131
-- 
-- Но, разумеется, если испытываем трудности с хранением, то стоит
-- отказаться от этой затеи, я сейчас руководствовался тем, что в 
-- приоритете скорость выполнения аналитического запроса.

alter table Sales
add constraint CHK_Sales_article_is_numeric
check (isnumeric(article) = 1);

alter table Sales
add constraint CHK_Sales_kg_not_negative
check (kg >= 0);

-- Индекс позволяет избавиться от сортировок, однако на деле прироста
-- в производительности на тестовых данных не наблюдается, поэтому 
-- закомментировал код ниже (возможно пригодится в будущем):
-- create nonclustered index IX_Sales_article_y_m_dt
-- on Sales(article, y, m, dt) include (kg);
