use Test;
go

create or alter procedure GetDataFromSales (
    @db as date = null, -- дата начала периода
    @de as date = null -- дата окончания периода
) as begin
    set nocount on;
    -- агрегируем данные по дням, считаем, сколько было продаж в среднем за каждый день
    -- года и месяца 
    with daily_sales as (
        select dt, y, m, article,
               sum(kg) day_total,
               avg(sum(kg)) over(partition by article, y) year_avg,
               avg(sum(kg)) over(partition by article, y, m) month_avg
        from Sales
        group by dt, y, m, article
    ),
    -- ограничиваем период заданными границами и считаем долю каждого артикула в выборке
    period_sales as (
        select *,
               sum(day_total) over(partition by article) / sum(day_total) over() article_percent,
               row_number() over(partition by article, y, m order by dt) rn
        from daily_sales
        where dt >= @db
          and dt <= @de
    )
    -- фильтруем лишние строки, переименовываем столбцы, форматируем числовые данные
    select y "Год",
           m "Месяц",
           article "Арктикул",
           year_avg "Средние продажи за год", 
           month_avg "Средние продажи за месяц",
           round(article_percent, 2) "Доля продаж артикула за выбранный период"
    from period_sales
    where rn = 1;
end;
