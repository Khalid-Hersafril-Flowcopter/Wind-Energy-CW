VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WakeLossAnalysisUF 
   Caption         =   "UserForm1"
   ClientHeight    =   2925
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5580
   OleObjectBlob   =   "WakeLossAnalysisUF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WakeLossAnalysisUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub analyzeWakeLoss_Click()
    Dim wind_turbine_data_range As Range
    Dim data_write_range As Range
    Dim row_offset As Double
    
    On Error Resume Next ' In case the Value in RefEdit is not a valid range
    Set wind_turbine_data_range = Range(windTurbineData.Value)
    Set data_write_range = Range(dataWrite.Value)
    row_offset = rowOffset.Value
    On Error GoTo 0 ' Stop error handling
    
    ' TODO (Khalid): Validate each data
    If Not wind_turbine_data_range Is Nothing Then
        ' Call the function and pass the selected range to it
            wakeLossAnalysis wind_turbine_data_range, data_write_rang, row_offset
            MsgBox "Data parsing complete."
    Else
        MsgBox "The selected range is not valid."
    End If
    
    Unload Me ' Close the UserForm
End Sub
