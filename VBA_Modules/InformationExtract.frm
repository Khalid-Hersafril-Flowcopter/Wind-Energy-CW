VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InformationExtract 
   Caption         =   "Wind Data Parser"
   ClientHeight    =   5625
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6390
   OleObjectBlob   =   "InformationExtract.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InformationExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub avgWindTurbulenceCheckbox_Click()
    If avgWindTurbulenceCheckbox.Value Then
        windTurbulenceData.Enabled = True
    Else
        windTurbulenceData.Enabled = False
    End If
End Sub

Private Sub SetInformationValue_Click()
    Dim datetime_selected_range As Range
    Dim wind_speed_selected_range As Range
    Dim new_datetime_selected_range As Range
    Dim wind_speed_average_selected_range As Range
    Dim wind_speed_direction_selected_range As Range
    Dim type_selection As String
    Dim write_uv_comp_flag As Boolean
    Dim wind_turbulence_flag As Boolean
    Dim wind_turbulence_selected_range As Range
     
    On Error Resume Next ' In case the Value in RefEdit is not a valid range
    Set datetime_selected_range = Range(DatetimeData.Value)
    Set wind_speed_selected_range = Range(targetColData.Value)
    Set new_datetime_selected_range = Range(newDatetimeData.Value)
    Set wind_speed_average_selected_range = Range(avgWindSpeedData.Value)
    Set wind_speed_direction_selected_range = Range(windSpeedDirectionData.Value)
    wind_turbulence_flag = avgWindTurbulenceCheckbox.Value
    type_selection = typeSelectionCBox.Value
    write_uv_comp_flag = writeUVCheckBox.Value
    
    If Not wind_turbulence_flag Then
        ' Ensure that we maintain the same function convention and not crash if data is empty
        Set wind_turbulence_selected_range = GetDummyRange()
    Else
        Set wind_turbulence_selected_range = Range(windTurbulenceData.Value)
        
        If wind_turbulence_selected_range Is Nothing Then
            MsgBox "You have enabled Wind Turbulence but entered nothing. Please fill in the data for wind turbulence. Thank you :)"
            Unload Me
            Exit Sub
        End If
    End If
    
    On Error GoTo 0 ' Stop error handling
    
    ' TODO (Khalid): Validate each data
    If Not datetime_selected_range Is Nothing Then
        ' Call the function and pass the selected range to it
            process_selected_range datetime_selected_range, wind_speed_selected_range, wind_speed_direction_selected_range, _
                                    new_datetime_selected_range, wind_speed_average_selected_range, type_selection, _
                                    write_uv_comp_flag, wind_turbulence_selected_range, wind_turbulence_flag
    Else
        MsgBox "The selected range is not valid."
    End If
    
    Unload Me ' Close the UserForm
End Sub

Private Function GetDummyRange() As Range
    ' This function returns a dummy range. Adjust the sheet name and range address as needed.
    Set GetDummyRange = Application.ActiveSheet.Range("A1")
End Function
