Attribute VB_Name = "WindDataProcess"
Public Function process_selected_range(ByRef datetimeRange As Range, ByRef uCompRange As Range, ByRef vCompRange As Range, tgtColRange As Range)
    ' Process the range however you need
    ' For example, just print out the address to the Immediate Window
    Debug.Print datetimeRange.Address
    Debug.Print uCompRange.Address
    Debug.Print vCompRange.Address
    Debug.Print tgtColRange.Address
    
    Dim datetimeValues As Variant
    Dim uCompValues As Variant
    Dim vCompValues As Variant
    Dim tgtColValues As Variant
    datetimeValues = datetimeRange.Value
    uCompValues = uCompRange.Value
    vCompValues = vCompRange.Value
    tgtColValues = tgtColRange.Value
    
    ' Get header name
    ' It is assumed that the header is at the first row of the column
    Dim datetime_header As String
    Dim u_comp_header As String
    Dim v_comp_header As String
    Dim tgt_col_header As String
    datetime_header = datetimeValues(1, 1)
    u_comp_header = uCompValues(1, 1)
    v_comp_header = vCompValues(1, 1)
    tgt_col_header = tgtColValues(1, 1)
    
    Dim lastUsedRow As Long
    With datetimeRange
        ' Find the last row with data starting from the bottom of the selected column range
        lastUsedRow = .Cells(.Rows.Count, 1).End(xlUp).Row - .Row + 1
    End With
    
    datetime_len = getRowLen(datetimeRange)
    u_comp_len = getRowLen(uCompRange)
    v_comp_len = getRowLen(vCompRange)
    tgt_col_len = getRowLen(tgtColRange)
    
    Debug.Print datetime_len
    Debug.Print u_comp_len
    Debug.Print v_comp_len
    Debug.Print tgt_col_len

End Function

Function getRowLen(rng As Range)
    Dim last_used_row As Long
    With rng
        last_used_row = .Cells(.Rows.Count, 1).End(xlUp).Row - .Row + 1
    End With
    
    getRowLen = last_used_row
End Function

