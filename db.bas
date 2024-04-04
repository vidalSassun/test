Attribute VB_Name = "db"
'----------------------------------------------------------------------
' ������ ��� ������ � ������ ������.
'----------------------------------------------------------------------

Option Explicit

'----------------------------------------------------------------------
' ���������� ������� SQL � ���������� ����������� �� ����.
'----------------------------------------------------------------------
Sub get_data()

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    Dim db As String, de As String
    db = Format(Worksheets("����������").Range("C3").value, "YYYY-MM-DD")
    de = Format(Worksheets("����������").Range("C4").value, "YYYY-MM-DD")
    '����������� ����������� � ���� ������
    Set conn = CreateObject("ADODB.Connection")
    
    '��������� ������ ����������� � ���� ������
    conn.ConnectionString = "DSN=docker_mssql;" & _
                            "Database=Test;" & _
                            "Uid=sa;" & _
                            "Pwd=Hh#!098765_"
    '��������� �����������
    conn.Open
    
    '��������� ������� �� ������� SQL
    Set cmd = CreateObject("ADODB.Command")
    With cmd
    
        '� �������� ����������� ���������� ������ ��� ����������� �����������
        Set .ActiveConnection = conn
        '��������� �������� ������
        .CommandType = adCmdStoredProc
        .CommandText = "GetDataFromSales"
        .Parameters.Append .CreateParameter("@db", adDBTimeStamp, adParamInput, , db)
        .Parameters.Append .CreateParameter("@de", adDBTimeStamp, adParamInput, , de)
        '���������� ���������� � Recordset
        Set rs = .Execute
        
    End With
    
    '�������� ������ �� Recordset �� ���� ����� �����
    With Workbooks.Add.Worksheets(1)
        
        .Name = "�����"
        '������� ������ ������
        .UsedRange.Delete
    
        '��������� ���������
        Dim i As Integer
        For i = 0 To rs.fields.Count - 1
            .Cells(1, i + 1).value = rs.fields(i).Name
        Next
        .Rows(1).EntireRow.Font.Bold = True
        
        '��������� ��������
        .Range("A2").CopyFromRecordset rs
        
        '��������� �������
        .UsedRange.EntireColumn.AutoFit
        
        '���������� ������ ������
        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        ActiveWindow.FreezePanes = True
    End With
    
    '��������� �����������
    conn.Close
End Sub

'----------------------------------------------------------------------
' ������������ xml �� �������� ���������.
'
' ��������� ��������, ��������� ��� ������� �� ���� �����, ��� ������
' ������ �������� ��������� �������.
'
' ���������:
' ----------
'     source - �������� � ������� ��� ������������ xml.
'
' ������������ ��������:
' ----------------------
'     ������ xml.
'----------------------------------------------------------------------
Function get_xml(ByRef source As Range) As String
    
    '������-�������� ��� ����� xml
    Dim xml_rows() As String
    '��������� ���������� ����� � ���������
    If source.Rows.Count < 2 Then Exit Function
    ReDim xml_rows(source.Rows.Count - 1)
    '�������� �������� �� ��������� � ������
    Dim arr As Variant
    arr = source.value
    '��������� xml
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
' ������ ������������ �����-��������� � ������������.
'
' ������������ ��������:
' ----------------------
'     ���� � �����.
'----------------------------------------------------------------------
Function get_path() As String
    Dim fd As Object
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .Title = "�������� ���� ��� �������� � ���� ������"
        .Filters.Clear
        .Filters.Add "Excel 2003", "*.xls?"
        If .Show = True Then get_path = Dir(.SelectedItems(1))
    End With
End Function

'----------------------------------------------------------------------
' ���������� ������� SQL ��� ������ ������ � ������� ���� ������.
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
    '����������� ����������� � ���� ������
    Set conn = CreateObject("ADODB.Connection")
    
    '��������� ������ ����������� � ���� ������
    conn.ConnectionString = "DSN=docker_mssql;" & _
                            "Database=Test;" & _
                            "Uid=sa;" & _
                            "Pwd=Hh#!098765_"
    '��������� �����������
    conn.Open
    
    '��������� ������� �� ������� SQL
    Set cmd = CreateObject("ADODB.Command")
    With cmd
    
        '� �������� ����������� ���������� ������ ��� ����������� �����������
        Set .ActiveConnection = conn
        '��������� �������� ������
        .CommandType = adCmdStoredProc
        .CommandText = "UploadDataToSales"
        .Parameters.Append .CreateParameter("@xml", adLongVarChar, adParamInput, 2147483647, xml)
        '���������� ���������� � Recordset
        .Execute
        
    End With
    '��������� �����������
    conn.Close
    MsgBox "�������� ���������"
End Sub




