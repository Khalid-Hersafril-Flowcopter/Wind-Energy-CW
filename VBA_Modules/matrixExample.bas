Attribute VB_Name = "DevExample"
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

Sub ExampleUsage()
Attribute ExampleUsage.VB_ProcData.VB_Invoke_Func = "t\n14"
    ' Create the matrix using the createMatrix function
    Dim wind_turbine_dict As Object: Set wind_turbine_dict = CreateObject("Scripting.Dictionary")
    wind_turbine_dict.Add "T1", Array(1, 2)
    wind_turbine_dict.Add "T2", Array(3, 4)
    wind_turbine_dict.Add "T3", Array(3, 4)
    wind_turbine_dict.Add "T4", Array(3, 4)
    
    Dim property_dict As Scripting.Dictionary
    Set property_dict = New Scripting.Dictionary
    property_dict.Add "P1", Array(1, 2)
    property_dict.Add "P2", Array(3, 4)
    
'    Dim i As Long
'    For i = 0 To wind_turbine_dict.Count - 1
'
'        Debug.Print wind_turbine_dict.keys()(i), wind_turbine_dict.Items()(i)(0), wind_turbine_dict.Items()(i)(1)
'    Next i
    
    Dim myMatrix As Variant
    myMatrix = createMatrix(wind_turbine_dict, property_dict) ' Example: 10 rows, 5 columns

    ' Define the starting cell on the sheet
    Dim startCell As Range
    Set startCell = ThisWorkbook.Sheets("Sheet1").Range("F5") ' Change to your actual sheet name and start cell

    ' Write the matrix to the sheet
    WriteMatrixToSheet myMatrix, startCell
End Sub

Function createTempMatrix(n, m)
   Dim matrix() As Integer
   ReDim matrix(1 To n, 1 To m) As Integer
   x = 1

    For i = 1 To n
       For j = 1 To m
            matrix(i, j) = x
            x = (x + 1)
        Next j
    Next i

    createMatrix = matrix
End Function

Function createMatrix(dict_1 As Scripting.Dictionary, dict_2 As Scripting.Dictionary)
    m = dict_1.Count
    n = dict_2.Count
    
    Dim matrix() As Integer
    ReDim matrix(1 To n, 1 To m) As Integer
    x = 1
    
    For i = 1 To n
       For j = 1 To m
            matrix(i, j) = x
            x = (x + 1)
        Next j
    Next i
    
    createMatrix = matrix
End Function
