VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} shadowFlickerAnalysisUF 
   Caption         =   "Shadow Flicker Analysis"
   ClientHeight    =   4845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5490
   OleObjectBlob   =   "shadowFlickerAnalysisUF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "shadowFlickerAnalysisUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub analyzeShadowFlicker_Click()
    Dim wind_turbine_data_range As Range
    Dim property_data_range As Range
    Dim data_write_range As Range
    Dim transpose_flag As Boolean
    Dim row_offset As Double
    Dim avg_sunlight As Double
    Dim correction_factor As Double
    
    On Error Resume Next ' In case the Value in RefEdit is not a valid range
    Set wind_turbine_data_range = Range(windTurbineData.Value)
    Set property_data_range = Range(propertyData.Value)
    Set data_write_range = Range(dataWrite.Value)
    transpose_flag = transposeCheckbox.Value
    row_offset = rowOffset.Value
    avg_sunlight = avgSunlight.Value
    correction_factor = corrFactor.Value
    On Error GoTo 0 ' Stop error handling
    
    ' TODO (Khalid): Validate each data
    If Not wind_turbine_data_range Is Nothing Then
        ' Call the function and pass the selected range to it
            shadowFlickerAnalysis wind_turbine_data_range, property_data_range, avg_sunlight, correction_factor, data_write_range, transpose_flag, row_offset
            MsgBox "Data parsing complete."
    Else
        MsgBox "The selected range is not valid."
    End If
    
    Unload Me ' Close the UserForm
End Sub
