Attribute VB_Name = "WakeLoss"
Function wakeLossAnalysis(ByRef wind_turbine_data As Range, ByRef data_write_range As Range, _
                            ByRef row_offset As Double)
                                
    Debug.Print wind_turbine_data.address
    
    ' Assert the first row of the box is the labels
    Dim header As Range
    Set header = wind_turbine_data.Rows(1)

    ' Check if the first row contains labels (you could add more checks here)
    If IsEmpty(header.Cells(1, 1).Value) Or IsEmpty(header.Cells(1, 2).Value) Or IsEmpty(header.Cells(1, 3).Value) Then
        MsgBox "The first row does not contain header!"
        Exit Function
    End If

    ' I don't understand how tf I cannot change the function name without breaking it, so Im leaving the name as it is
    ' although GetTurbineData is a generic function that parses "object", "x", "y" data
    Dim wind_data_dict As Scripting.Dictionary
    Set wind_data_dict = GetTurbineData(wind_turbine_data)
    
    Dim init_matrix_col As String: init_matrix_col = getColumnLetter(data_write_range.address)
    
    ' Force the property names to be written at row 3 of the sheets
    Dim init_matrix_col_num As Long: init_matrix_col_num = colLetterToNumber(init_matrix_col)
    
    Dim col_data_write_str_init As String
    Dim row_data_write_str_init As String
    Dim distance_val_str_col As String
    Dim direction_val_str_col As String
    Dim vel_factor_val_str_col As String
    
    Dim data_count As Double
    
    row_data_write_str_init = init_matrix_col
    ' Force the wind_turbine names written at 2nd row of the sheets
    col_data_write_str_init = colNumberToLetter(init_matrix_col_num + 1)
    data_count = wind_data_dict.count
    
    ' This is useful if user wants to write their sound analysis on the same column but with an offset of a certain row length
    distance_val_str_col = colNumberToLetter(init_matrix_col_num + 1) & (3 + row_offset)
    
    ' Write this below the distance matrix
    direction_val_str_col = colNumberToLetter(init_matrix_col_num + 1) & (4 + (data_count + 1) + row_offset)
    
    ' Write this below the direction matrix (there will be n * data count rows above)
    vel_factor_val_str_col = colNumberToLetter(init_matrix_col_num + 1) & (4 + (2 * data_count + 3) + row_offset)

    Dim i As Long
    
    For i = 0 To wind_data_dict.count - 1
        ' Force the property names to be written incrementally from the 3rd Row
        Range(init_matrix_col & ((3 + row_offset) + i)) = wind_data_dict.keys()(i)
        Range(init_matrix_col & ((3 + (data_count + 2) + row_offset) + i)) = wind_data_dict.keys()(i)
        Range(init_matrix_col & ((3 + 2 * (data_count + 2) + row_offset) + i)) = wind_data_dict.keys()(i)
    Next i
    
    Dim col_data_write_str As String
    For i = 0 To wind_data_dict.count - 1
        ' Since the wind turbine is written column to column, we have to incrementally increase the letter value
        ' but force it to be at 2nd row
        col_data_write_str = colNumberToLetter(colLetterToNumber(col_data_write_str_init) + i)
        Range(col_data_write_str & (2 + row_offset)) = wind_data_dict.keys()(i)
        
        ' Here I have used 3 instead of 2 when adding it with the row offset because I want to leave some
        ' space between the headers
        Range(col_data_write_str & (3 + (data_count + 1) + row_offset)) = wind_data_dict.keys()(i)
        
        ' I dont know why adding 3 there works but my brain is fried so I am just leaving this here
        Range(col_data_write_str & (3 + (2 * data_count + 3) + row_offset)) = wind_data_dict.keys()(i)
    Next i
    
    ' Define the starting cell on the sheet
    Dim startCell As Range
    Set startCell = ThisWorkbook.Sheets(ActiveSheet.Name).Range(distance_val_str_col) ' Change to your actual sheet name and start cell
    Dim distance_matrix As Variant: distance_matrix = createDistanceMatrix(wind_data_dict, wind_data_dict)

    ' Write the matrix to the sheet
    WriteMatrixToSheet distance_matrix, startCell
    
    Set startCell = ThisWorkbook.Sheets(ActiveSheet.Name).Range(direction_val_str_col) ' Change to your actual sheet name and start cell
    Dim direction_matrix As Variant: direction_matrix = createDirectionMatrix(wind_data_dict, wind_data_dict)

    ' Write the matrix to the sheet
    WriteMatrixToSheet direction_matrix, startCell
    
    Set startCell = ThisWorkbook.Sheets(ActiveSheet.Name).Range(vel_factor_val_str_col) ' Change to your actual sheet name and start cell
    Dim vel_factor_matrix As Variant: vel_factor_matrix = createVelFactorMatrix(wind_data_dict, wind_data_dict, distance_matrix)

    ' Write the matrix to the sheet
    WriteMatrixToSheet vel_factor_matrix, startCell
    
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

Private Function splitAddress(address As String) As Variant
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

Private Function GetTurbineData(rng As Range) As Scripting.Dictionary
    Dim dict As New Scripting.Dictionary
    Dim cell As Range
    Dim key As String
    Dim x As Double
    Dim y As Double
    Dim noise_lvl As Double
    
    ' Loop through each row in the range, skipping the header
    For Each cell In rng.Offset(1, 0).Resize(rng.Rows.count - 1, 1).Cells
        key = cell.Value ' Turbine name
        x = cell.Offset(0, 1).Value ' X coordinate
        y = cell.Offset(0, 2).Value ' Y coordinate
        diameter = cell.Offset(0, 3) ' Diameter of turbine
        setback_distance = cell.Offset(0, 4).Value ' Setback distance
        dict(key) = Array(x, y, diameter, setback_distance) ' Add to dictionary as an array (which is like a tuple)
    Next cell
    
    Set GetTurbineData = dict
End Function

' Generate Matrix for Distance
Private Function createDistanceMatrix(dict_1 As Scripting.Dictionary, dict_2 As Scripting.Dictionary)

    col_count = dict_1.count
    row_count = dict_2.count

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
Private Function createDirectionMatrix(dict_1 As Scripting.Dictionary, dict_2 As Scripting.Dictionary)

    col_count = dict_1.count
    row_count = dict_2.count

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
            
            Dim x_dist As Double: x_dist = dict1_pos(0) - dict2_pos(0)
            Dim y_dist As Double: y_dist = dict1_pos(1) - dict2_pos(1)
            
            Dim angle_rad As Double
            Dim angle_deg As Double
            If Not (x_dist = 0 Or y_dist = 0) Then
                angle_rad = Application.WorksheetFunction.Atan2(y_dist, x_dist)
                angle_deg = radToDeg(angle_rad)
            Else
                angle_deg = 0
            End If
            
            If angle_deg = 0 Then
                angle_deg = 0
            ElseIf angle_deg < 0 Then
                angle_deg = (360 + angle_deg)
            End If
            
            matrix(curr_row, curr_col) = angle_deg
        Next curr_col
    Next curr_row
    
    createDirectionMatrix = matrix
    
End Function

' Generate Matrix for Distance
Function createVelFactorMatrix(dict_1 As Scripting.Dictionary, dict_2 As Scripting.Dictionary, ParamArray distance_matrix() As Variant)

    col_count = dict_1.count
    row_count = dict_2.count
    
    Dim matrix() As Double
    ReDim matrix(1 To row_count, 1 To col_count) As Double
    
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
            Dim x As Double: x = distance_matrix(0)(curr_row, curr_col)
            
            Dim Ct As Double: Ct = 0.89
            Dim k As Double: k = 0.075
            Dim D As Double: D = dict1_array(2)
            
            If Not ((dict1_array(0) = dict2_array(0)) And (dict1_array(1) = dict2_array(1))) Then
                matrix(curr_row, curr_col) = 1 - (1 - Sqr(1 - Ct)) * (D / (D + 2 * k * x)) ^ 2
            Else
                ' For turbines that is referencing itself, the velocity ratio should be 1
                matrix(curr_row, curr_col) = 1
            End If

        Next curr_col
    Next curr_row
    
    createVelFactorMatrix = matrix
End Function

Private Function WriteMatrixToSheet(matrix As Variant, startCell As Range)
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
        writeRange.Value = matrix
    End With
End Function

Private Function radToDeg(angle_rad As Double) As Double
    radToDeg = Application.WorksheetFunction.Degrees(angle_rad)
End Function

Private Function degToRad(angle_deg As Double) As Double
    degToRad = Application.WorksheetFunction.Radians(angle_deg)
End Function





