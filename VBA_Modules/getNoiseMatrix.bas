Attribute VB_Name = "getNoiseMatrix"
Function getNoiseMatrixFunction(ByRef wind_turbine_data_range As Range, ByRef property_data_range As Range, ByRef noise_matrix_write_range As Range, _
                                ByRef turbine_diameter As Double, ByRef tip_speed As Double, ByRef alpha As Double)
                                
    Debug.Print wind_turbine_data_range.address, property_data_range.address, noise_matrix_write_range.address, turbine_diameter, tip_speed, alpha
    
    ' Assert the first row of the box is the labels
    Dim header As Range
    Set header = wind_turbine_data_range.Rows(1)

    ' Check if the first row contains labels (you could add more checks here)
    If IsEmpty(header.Cells(1, 1).value) Or IsEmpty(header.Cells(1, 2).value) Or IsEmpty(header.Cells(1, 3).value) Then
        MsgBox "The first row does not contain header!"
        Exit Function
    End If

    Dim turbinesData As Scripting.Dictionary
    Set turbinesData = GetTurbineData(wind_turbine_data_range)
    
    ' Example of how to use the turbinesData
    Dim k As Variant
    For Each k In turbinesData.keys
        Debug.Print "Turbine: " & k & ", Coordinates: (" & turbinesData(k)(0) & ", " & turbinesData(k)(1) & ")"
    Next k

    MsgBox "Data parsing complete."
    
End Function

Private Function getColumnLetter(columnRef As String) As String
    ' Replace removes the $ signs, and then we take the first character since the column is the same.
    getColumnLetter = Replace(columnRef, "$", "")
    ' If there's a colon indicating a range of the same column, it is removed as well.
    If InStr(getColumnLetter, ":") > 0 Then
        getColumnLetter = Left(getColumnLetter, InStr(getColumnLetter, ":") - 1)
    End If
End Function

Private Function colLetterToNumber(col_letter As String) As Double
    colLetterToNumber = Range(col_letter & 1).Column
End Function

Private Function colNumberToLetter(col_number As Double) As String
    colNumberToLetter = Split(Cells(, col_number).address, "$")(1)
End Function

Function splitAddress(address As String) As Variant
    Dim i As Integer
    Dim letterPart As String
    Dim numberPart As Integer
    
    ' Loop through each character in the string
    For i = 1 To Len(address)
        If IsNumeric(Mid(address, i, 1)) Then
            numberPart = numberPart & Mid(address, i, 1)
        Else
            letterPart = letterPart & Mid(address, i, 1)
        End If
    Next i
    
    ' Return both parts as an array
    splitAddress = Array(letterPart, numberPart)
End Function

Function GetTurbineData(rng As Range) As Scripting.Dictionary
    Dim dict As New Scripting.Dictionary
    Dim cell As Range
    Dim key As String
    Dim x As Long
    Dim y As Long
    
    ' Loop through each row in the range, skipping the header
    For Each cell In rng.Offset(1, 0).Resize(rng.Rows.Count - 1, 1).Cells
        key = cell.value ' Turbine name
        x = cell.Offset(0, 1).value ' X coordinate
        y = cell.Offset(0, 2).value ' Y coordinate
        dict(key) = Array(x, y) ' Add to dictionary as an array (which is like a tuple)
    Next cell
    
    Set GetTurbineData = dict
End Function
