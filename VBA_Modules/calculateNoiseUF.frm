VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} calculateNoiseUF 
   Caption         =   "Noise Data Parser"
   ClientHeight    =   3540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4965
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
    Dim alpha_val As Double
    
    On Error Resume Next ' In case the Value in RefEdit is not a valid range
    Set wind_turbine_data_range = Range(windTurbineData.value)
    Set property_data_range = Range(propertyData.value)
    Set noise_matrix_write_range = Range(noiseMatrixWrite.value)
    alpha_val = alphaVal.value
    On Error GoTo 0 ' Stop error handling
    
    ' TODO (Khalid): Validate each data
    If Not wind_turbine_data_range Is Nothing Then
        ' Call the function and pass the selected range to it
            getNoiseMatrixFunction wind_turbine_data_range, property_data_range, noise_matrix_write_range, alpha_val
    Else
        MsgBox "The selected range is not valid."
    End If
    
    Unload Me ' Close the UserForm
End Sub

