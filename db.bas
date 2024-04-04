Attribute VB_Name = "db"
'----------------------------------------------------------------------
' ћодуль дл€ работы с базами данных.
'----------------------------------------------------------------------

Option Explicit

'----------------------------------------------------------------------
' ¬ыполнение запроса SQL и сохранение результатов на лист.
'----------------------------------------------------------------------
Sub get_data()

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    Dim db As String, de As String
    db = Format(Worksheets("управление").Range("C3").value, "YYYY-MM-DD")
    de = Format(Worksheets("управление").Range("C4").value, "YYYY-MM-DD")
    'настраиваем подключение к базе данных
    Set conn = CreateObject("ADODB.Connection")
    
    'формируем строку подключени€ к базе данных
    conn.ConnectionString = "DSN=docker_mssql;" & _
                            "Database=Test;" & _
                            "Uid=sa;" & _
                            "Pwd=Hh#!098765_"
    'открываем подключение
    conn.Open
    
    'выполн€ем команды на сервере SQL
    Set cmd = CreateObject("ADODB.Command")
    With cmd
    
        'в качестве подключени€ используем только что настроенное подключение
        Set .ActiveConnection = conn
        'выполн€ем основной запрос
        .CommandType = adCmdStoredProc
        .CommandText = "GetDataFromSales"
        .Parameters.Append .CreateParameter("@db", adDBTimeStamp, adParamInput, , db)
        .Parameters.Append .CreateParameter("@de", adDBTimeStamp, adParamInput, , de)
        'результаты записываем в Recordset
        Set rs = .Execute
        
    End With
    
    'копируем данные из Recordset на лист новой книги
    With Workbooks.Add.Worksheets(1)
        
        .Name = "отчЄт"
        'удал€ем старые данные
        .UsedRange.Delete
    
        'заполн€ем заголовки
        Dim i As Integer
        For i = 0 To rs.fields.Count - 1
            .Cells(1, i + 1).value = rs.fields(i).Name
        Next
        .Rows(1).EntireRow.Font.Bold = True
        
        'вставл€ем значени€
        .Range("A2").CopyFromRecordset rs
        
        'расшир€ем колонки
        .UsedRange.EntireColumn.AutoFit
        
        'закрепл€ем первую строку
        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        ActiveWindow.FreezePanes = True
    End With
    
    'закрываем подключение
    conn.Close
End Sub

'----------------------------------------------------------------------
' ‘ормирование xml из значений диапазона.
'
' Ќеобходим диапазон, состо€щий как минимум из двух строк, где перва€
' строка содержит заголовки таблицы.
'
' јргументы:
' ----------
'     source - диапазон с данными дл€ формировани€ xml.
'
' ¬озвращаемое значение:
' ----------------------
'     —трока xml.
'----------------------------------------------------------------------
Function get_xml(ByRef source As Range) As String
    
    'массив-приемник дл€ строк xml
    Dim xml_rows() As String
    'провер€ем количество строк в диапазоне
    If source.Rows.Count < 2 Then Exit Function
    ReDim xml_rows(source.Rows.Count - 1)
    'копируем значени€ из диапазона в массив
    Dim arr As Variant
    arr = source.value
    'формируем xml
    Dim row_index As Long, column_index As Long, xml_row As String, value As Variant
    For row_index = 2 To UBound(arr, 1)
        xml_row = "<row "
        For column_index = 1 To UBound(arr, 2)
            If IsDate(arr(row_index, column_index)) Then
                value = Format(arr(row_index, column_index), "YYYY-MM-DD")
            Else
                value = arr(row_index, column_index)
            End If
            xml_row = xml_row & " " & arr(1, column_index) & " = """ & value & """"
        Next column_index
        xml_row = xml_row & "/>"
        xml_rows(row_index - 2) = xml_row
    Next row_index
    get_xml = Join(xml_rows, vbCrLf)
End Function

'----------------------------------------------------------------------
' «апрос расположени€ файла-источника у пользовател€.
'
' ¬озвращаемое значение:
' ----------------------
'     ѕуть к файлу.
'----------------------------------------------------------------------
Function get_path() As String
    Dim fd As Object
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .Title = "¬ыберете файл дл€ загрузки в базу данных"
        .Filters.Clear
        .Filters.Add "Excel 2003", "*.xls?"
        If .Show = True Then get_path = Dir(.SelectedItems(1))
    End With
End Function

'----------------------------------------------------------------------
' ¬ыполнение запроса SQL дл€ записи данных в таблицу базы данных.
'----------------------------------------------------------------------
Sub upload_data()

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    Dim xml As String
    
    With Workbooks.Open(get_path())
        xml = get_xml(.Worksheets(1).UsedRange)
        .Close SaveChanges:=False
    End With
    'настраиваем подключение к базе данных
    Set conn = CreateObject("ADODB.Connection")
    
    'формируем строку подключени€ к базе данных
    conn.ConnectionString = "DSN=docker_mssql;" & _
                            "Database=Test;" & _
                            "Uid=sa;" & _
                            "Pwd=Hh#!098765_"
    'открываем подключение
    conn.Open
    
    'выполн€ем команды на сервере SQL
    Set cmd = CreateObject("ADODB.Command")
    With cmd
    
        'в качестве подключени€ используем только что настроенное подключение
        Set .ActiveConnection = conn
        'выполн€ем основной запрос
        .CommandType = adCmdStoredProc
        .CommandText = "UploadDataToSales"
        .Parameters.Append .CreateParameter("@xml", adLongVarChar, adParamInput, 2147483647, xml)
        'результаты записываем в Recordset
        .Execute
        
    End With
    'закрываем подключение
    conn.Close
    MsgBox "«агрузка завершена"
End Sub




