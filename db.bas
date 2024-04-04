Attribute VB_Name = "db"
'----------------------------------------------------------------------
' Ìîäóëü äëÿ ðàáîòû ñ áàçàìè äàííûõ.
'----------------------------------------------------------------------

Option Explicit

'----------------------------------------------------------------------
' Âûïîëíåíèå çàïðîñà SQL è ñîõðàíåíèå ðåçóëüòàòîâ íà ëèñò.
'----------------------------------------------------------------------
Sub get_data()

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    Dim db As String, de As String
    db = Format(Worksheets("óïðàâëåíèå").Range("C3").value, "YYYY-MM-DD")
    de = Format(Worksheets("óïðàâëåíèå").Range("C4").value, "YYYY-MM-DD")
    'íàñòðàèâàåì ïîäêëþ÷åíèå ê áàçå äàííûõ
    Set conn = CreateObject("ADODB.Connection")
    
    'ôîðìèðóåì ñòðîêó ïîäêëþ÷åíèÿ ê áàçå äàííûõ
    conn.ConnectionString = "DSN=docker_mssql;" & _
                            "Database=Test;" & _
                            "Uid=sa;" & _
                            "Pwd=Hh#!098765_"
    'îòêðûâàåì ïîäêëþ÷åíèå
    conn.Open
    
    'âûïîëíÿåì êîìàíäû íà ñåðâåðå SQL
    Set cmd = CreateObject("ADODB.Command")
    With cmd
    
        'â êà÷åñòâå ïîäêëþ÷åíèÿ èñïîëüçóåì òîëüêî ÷òî íàñòðîåííîå ïîäêëþ÷åíèå
        Set .ActiveConnection = conn
        'âûïîëíÿåì îñíîâíîé çàïðîñ
        .CommandType = adCmdStoredProc
        .CommandText = "GetDataFromSales"
        .Parameters.Append .CreateParameter("@db", adDBTimeStamp, adParamInput, , db)
        .Parameters.Append .CreateParameter("@de", adDBTimeStamp, adParamInput, , de)
        'ðåçóëüòàòû çàïèñûâàåì â Recordset
        Set rs = .Execute
        
    End With
    
    'êîïèðóåì äàííûå èç Recordset íà ëèñò íîâîé êíèãè
    With Workbooks.Add.Worksheets(1)
        
        .Name = "îò÷¸ò"
        'óäàëÿåì ñòàðûå äàííûå
        .UsedRange.Delete
    
        'çàïîëíÿåì çàãîëîâêè
        Dim i As Integer
        For i = 0 To rs.fields.Count - 1
            .Cells(1, i + 1).value = rs.fields(i).Name
        Next
        .Rows(1).EntireRow.Font.Bold = True
        
        'âñòàâëÿåì çíà÷åíèÿ
        .Range("A2").CopyFromRecordset rs
        
        'ðàñøèðÿåì êîëîíêè
        .UsedRange.EntireColumn.AutoFit
        
        'çàêðåïëÿåì ïåðâóþ ñòðîêó
        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        ActiveWindow.FreezePanes = True
    End With
    
    'çàêðûâàåì ïîäêëþ÷åíèå
    conn.Close
End Sub

'----------------------------------------------------------------------
' Ôîðìèðîâàíèå xml èç çíà÷åíèé äèàïàçîíà.
'
' Íåîáõîäèì äèàïàçîí, ñîñòîÿùèé êàê ìèíèìóì èç äâóõ ñòðîê, ãäå ïåðâàÿ
' ñòðîêà ñîäåðæèò çàãîëîâêè òàáëèöû.
'
' Àðãóìåíòû:
' ----------
'     source - äèàïàçîí ñ äàííûìè äëÿ ôîðìèðîâàíèÿ xml.
'
' Âîçâðàùàåìîå çíà÷åíèå:
' ----------------------
'     Ñòðîêà xml.
'----------------------------------------------------------------------
Function get_xml(ByRef source As Range) As String
    
    'ìàññèâ-ïðèåìíèê äëÿ ñòðîê xml
    Dim xml_rows() As String
    'ïðîâåðÿåì êîëè÷åñòâî ñòðîê â äèàïàçîíå
    If source.Rows.Count < 2 Then Exit Function
    ReDim xml_rows(source.Rows.Count - 1)
    'êîïèðóåì çíà÷åíèÿ èç äèàïàçîíà â ìàññèâ
    Dim arr As Variant
    arr = source.value
    'ôîðìèðóåì xml
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
' Çàïðîñ ðàñïîëîæåíèÿ ôàéëà-èñòî÷íèêà ó ïîëüçîâàòåëÿ.
'
' Âîçâðàùàåìîå çíà÷åíèå:
' ----------------------
'     Ïóòü ê ôàéëó.
'----------------------------------------------------------------------
Function get_path() As String
    Dim fd As Object
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .Title = "Âûáåðåòå ôàéë äëÿ çàãðóçêè â áàçó äàííûõ"
        .Filters.Clear
        .Filters.Add "Excel 2003", "*.xls?"
        If .Show = True Then get_path = Dir(.SelectedItems(1))
    End With
End Function

'----------------------------------------------------------------------
' Âûïîëíåíèå çàïðîñà SQL äëÿ çàïèñè äàííûõ â òàáëèöó áàçû äàííûõ.
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
    'íàñòðàèâàåì ïîäêëþ÷åíèå ê áàçå äàííûõ
    Set conn = CreateObject("ADODB.Connection")
    
    'ôîðìèðóåì ñòðîêó ïîäêëþ÷åíèÿ ê áàçå äàííûõ
    conn.ConnectionString = "DSN=docker_mssql;" & _
                            "Database=Test;" & _
                            "Uid=sa;" & _
                            "Pwd=Hh#!098765_"
    'îòêðûâàåì ïîäêëþ÷åíèå
    conn.Open
    
    'âûïîëíÿåì êîìàíäû íà ñåðâåðå SQL
    Set cmd = CreateObject("ADODB.Command")
    With cmd
    
        'â êà÷åñòâå ïîäêëþ÷åíèÿ èñïîëüçóåì òîëüêî ÷òî íàñòðîåííîå ïîäêëþ÷åíèå
        Set .ActiveConnection = conn
        'âûïîëíÿåì îñíîâíîé çàïðîñ
        .CommandType = adCmdStoredProc
        .CommandText = "UploadDataToSales"
        .Parameters.Append .CreateParameter("@xml", adLongVarChar, adParamInput, 2147483647, xml)
        'ðåçóëüòàòû çàïèñûâàåì â Recordset
        .Execute
        
    End With
    'çàêðûâàåì ïîäêëþ÷åíèå
    conn.Close
    MsgBox "Çàãðóçêà çàâåðøåíà"
End Sub




