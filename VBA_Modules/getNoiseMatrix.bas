Attribute VB_Name = "getNoiseMatrix"
Function getNoiseMatrixFunction(ByRef wind_turbine_data_range As Range, ByRef property_data_range As Range, _
                                ByRef write_matrix_range As Range, ByRef alpha_val As Double)
                                
    Debug.Print wind_turbine_data_range.address, property_data_range.address, write_matrix_range.address, alpha_val
    
    ' Assert the first row of the box is the labels
    Dim header As Range
    Set header = wind_turbine_data_range.Rows(1)

    ' Check if the first row contains labels (you could add more checks here)
    If IsEmpty(header.Cells(1, 1).value) Or IsEmpty(header.Cells(1, 2).value) Or IsEmpty(header.Cells(1, 3).value) Then
        MsgBox "The first row does not contain header!"
        Exit Function
    End If

    ' I don't understand how tf I cannot change the function name without breaking it, so Im leaving the name as it is
    ' although GetTurbineData is a generic function that parses "object", "x", "y" data
    Dim wind_turbine_dict As Scripting.Dictionary: Set wind_turbine_dict = GetTurbineData(wind_turbine_data_range)
    Dim property_dict As Scripting.Dictionary: Set property_dict = GetTurbineData(property_data_range)
    
    ' Example of how to use the turbinesData
    Dim k As Variant
    Dim n As Variant
    For Each k In wind_turbine_dict.keys
        Debug.Print "Turbine: " & k & ", Coordinates: (" & wind_turbine_dict(k)(0) & ", " & wind_turbine_dict(k)(1) & ", " & wind_turbine_dict(k)(2) & ")"
        
        For Each n In property_dict.keys
            Debug.Print "Property: " & n & ", Coordinates: (" & property_dict(n)(0) & ", " & property_dict(n)(1) & ", " & property_dict(n)(2) & ")"
        Next n
    Next k
    
    Dim init_matrix_col As String: init_matrix_col = getColumnLetter(write_matrix_range.address)
    
    ' Force the property names to be written at row 3 of the sheets
    Dim init_matrix_col_num As Long: init_matrix_col_num = colLetterToNumber(init_matrix_col)
    Dim property_str_col As String: property_str_col = init_matrix_col
    ' Force the wind_turbine names written at 2nd row of the sheets
    Dim wind_turbine_str_col_init As String: wind_turbine_str_col_init = colNumberToLetter(init_matrix_col_num + 1)
    Dim distance_val_str_col As String: distance_val_str_col = colNumberToLetter(init_matrix_col_num + 1) & 3
    
    ' Offset the noise calculation by the numbers of wind turbine
    Dim noise_val_str_col As String: noise_val_str_col = colNumberToLetter(colLetterToNumber(distance_val_str_col) _
                                                            + (wind_turbine_dict.Count + 1)) & 3
    
    Dim i As Long
    For i = 0 To property_dict.Count - 1
        ' Force the property names to be written incrementally from the 3rd Row
        Range(init_matrix_col & (3 + i)) = property_dict.keys()(i)
    Next i
    
    For i = 0 To wind_turbine_dict.Count - 1
        ' Since the wind turbine is written column to column, we have to incrementally increase the letter value
        ' but force it to be at 2nd row
        Dim wind_turbine_str_col As String: wind_turbine_str_col = colNumberToLetter(colLetterToNumber(wind_turbine_str_col_init) + i)
        Range(wind_turbine_str_col & 2) = wind_turbine_dict.keys()(i)
    Next i
    
    Debug.Print property_str_col, wind_turbine_str_col, distance_val_str_col, ActiveSheet.Name
    
    ' Define the starting cell on the sheet
    Dim startCell As Range
    Set startCell = ThisWorkbook.Sheets(ActiveSheet.Name).Range(distance_val_str_col) ' Change to your actual sheet name and start cell
    Dim distance_matrix As Variant
    distance_matrix = createDistanceMatrix(wind_turbine_dict, property_dict) ' Example: 10 rows, 5 columns
    ' Write the matrix to the sheet
    WriteMatrixToSheet distance_matrix, startCell


    ' Define the starting cell on the sheet
    Dim noiseStartCell As Range
    Set noiseStartCell = ThisWorkbook.Sheets(ActiveSheet.Name).Range(noise_val_str_col) ' Change to your actual sheet name and start cell
    Debug.Print noise_val_str_col
    Dim noise_matrix As Variant
    noise_matrix = createNoiseMatrix(wind_turbine_dict, property_dict, alpha_val, distance_matrix)
    WriteMatrixToSheet noise_matrix, noiseStartCell
    
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
    Dim noise_lvl As Double
    
    ' Loop through each row in the range, skipping the header
    For Each cell In rng.Offset(1, 0).Resize(rng.Rows.Count - 1, 1).Cells
        key = cell.value ' Turbine name
        x = cell.Offset(0, 1).value ' X coordinate
        y = cell.Offset(0, 2).value ' Y coordinate
        noise_lvl = cell.Offset(0, 3) ' Noise Level
        alpha = cell.Offset(0, 4)
        dict(key) = Array(x, y, noise_lvl, alpha) ' Add to dictionary as an array (which is like a tuple)
    Next cell
    
    Set GetTurbineData = dict
End Function

' Generate Matrix for Distance
Function createDistanceMatrix(dict_1 As Scripting.Dictionary, dict_2 As Scripting.Dictionary)
    col_count = dict_1.Count
    row_count = dict_2.Count
    
    Dim matrix() As Double
    ReDim matrix(1 To row_count, 1 To col_count) As Double
    x = 1
    
    ' VBA Matrix index starts with 1 so matrix(0, 0) does not exist!
    For curr_row = 1 To row_count
       For curr_col = 1 To col_count
            ' This is required since VBA Dictionary index starts at 0 which is kind of stupid and unintuitive
            Dim i As Long: i = curr_row - 1
            Dim j As Long: j = curr_col - 1
            
            ' Writing it this way to make the code more readable
            Dim dict1_pos As Variant: dict1_pos = dict_1.Items()(j)
            Dim dict2_pos As Variant: dict2_pos = dict_2.Items()(i)
            matrix(curr_row, curr_col) = Sqr((dict1_pos(0) - dict2_pos(0)) ^ 2 + (dict1_pos(1) - dict2_pos(1)) ^ 2)
        Next curr_col
    Next curr_row
    
    createDistanceMatrix = matrix
End Function

' Generate Matrix for Distance
Function createNoiseMatrix(dict_1 As Scripting.Dictionary, dict_2 As Scripting.Dictionary, alpha_val As Double, ParamArray distance_matrix() As Variant)

    col_count = dict_1.Count
    row_count = dict_2.Count
    
    Dim matrix() As Double
    ReDim matrix(1 To row_count, 1 To col_count) As Double
    x = 1
    
    ' VBA Matrix index starts with 1 so matrix(0, 0) does not exist!
    For curr_row = 1 To row_count
       For curr_col = 1 To col_count
            ' This is required since VBA Dictionary index starts at 0 which is kind of stupid and unintuitive
            Dim i As Long: i = curr_row - 1
            Dim j As Long: j = curr_col - 1
            
            ' Writing it this way to make the code more readable
            Dim dict1_array As Variant: dict1_array = dict_1.Items()(j)
            Dim dict2_array As Variant: dict2_array = dict_2.Items()(i)
            
            ' The only issue with this implementation is that if the user decides to flip wind turbine and property,
            ' The noise calculation would then be wrong since we have force to use dict1_array's noise value
            ' Remember that the dictionary format is
            ' Name, x, y, noise
            Dim R As Double: R = distance_matrix(0)(curr_row, curr_col)
            matrix(curr_row, curr_col) = dict1_array(2) - 10 * Application.WorksheetFunction.Log10(2 * 3.14159 * R ^ 2) - dict1_array(3) * R
        Next curr_col
    Next curr_row
    
    createNoiseMatrix = matrix
End Function

Function WriteMatrixToSheet(matrix As Variant, startCell As Range)
    ' Determine the size of the matrix
    Dim numRows As Long
    Dim numCols As Long
    numRows = UBound(matrix, 1) - LBound(matrix, 1) + 1
    numCols = UBound(matrix, 2) - LBound(matrix, 2) + 1
    
    ' Define the range that will be populated with the matrix values
    With startCell
        Dim endCell As Range
        Set endCell = .Offset(numRows - 1, numCols - 1)
        Dim writeRange As Range
        Set writeRange = .Worksheet.Range(.address & ":" & endCell.address)
        
        ' Write the matrix to the sheet
        writeRange.value = matrix
    End With
End Function

