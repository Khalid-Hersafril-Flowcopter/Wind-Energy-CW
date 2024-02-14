Attribute VB_Name = "ExtractHours"
Function ExtractHours()
    Dim work_sheet As Worksheet
    Set work_sheet = ActiveSheet
    
    Dim date_header As String
    date_header = "Date and Time"
    
    Dim last_column As Long
    last_column = work_sheet.Cells(1, work_sheet.Columns.count).End(xlToLeft).Column
    
    Dim date_header_found As Boolean
    date_header_found = False
    
    Dim i As Integer
    Dim date_header_pos As Integer
    For i = 1 To last_column
        
        If LCase(work_sheet.Cells(1, i).Value) = LCase(date_header) Then
            date_header_found = True
            date_header_pos = i
            Exit For
        End If
    Next i
    
    If Not date_header_found Then
        MsgBox "Header '" & date_header & "' not found."
    End If
    
End Function
