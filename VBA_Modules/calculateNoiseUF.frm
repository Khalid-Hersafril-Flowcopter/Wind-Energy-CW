VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} calculateNoiseUF 
   Caption         =   "Noise Data Parser"
   ClientHeight    =   3720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4980
   OleObjectBlob   =   "calculateNoiseUF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "calculateNoiseUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub calculateNoiseMatrixButton_Click()
    Dim wind_turbine_data_range As Range
    Dim new_wind_turbine_data_range As Range
    Dim property_data_range As Range
    Dim noise_matrix_write_range As Range
    Dim transpose_flag As Boolean
    Dim row_offset As Long
    
    On Error Resume Next ' In case the Value in RefEdit is not a valid range
    Set wind_turbine_data_range = Range(windTurbineData.value)
    Set property_data_range = Range(propertyData.value)
    Set noise_matrix_write_range = Range(noiseMatrixWrite.value)
    transpose_flag = transposeCheckbox.value
    row_offset = rowOffset.value
    On Error GoTo 0 ' Stop error handling
    
    ' TODO (Khalid): Validate each data
    If Not wind_turbine_data_range Is Nothing Then
        ' Call the function and pass the selected range to it
            getNoiseMatrixFunction wind_turbine_data_range, property_data_range, noise_matrix_write_range, transpose_flag, row_offset
    Else
        MsgBox "The selected range is not valid."
    End If
    
    Unload Me ' Close the UserForm
End Sub

