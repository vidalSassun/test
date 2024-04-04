use Test;
go

create or alter procedure UploadDataToSales (
    @xml as xml = null
) as begin
    set nocount on;
    -- парсим xml
    with input_data as (
        select t.c.value('@dt', 'datetime') dt,
               t.c.value('@article', 'varchar(10)') article,
               t.c.value('@kg', 'numeric(6, 3)') kg
        from @xml.nodes('//row') as t(c)
    )
    insert into Sales
    select *
    from input_data
    where kg >= 0 and isnumeric(article) = 1; -- очищаем от некорректных строк
end;