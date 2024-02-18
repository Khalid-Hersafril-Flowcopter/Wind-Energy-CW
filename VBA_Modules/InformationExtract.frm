VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InformationExtract 
   Caption         =   "UserForm1"
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6315
   OleObjectBlob   =   "InformationExtract.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InformationExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SetInformationValue_Click()
    Dim datetime_selected_range As Range
    Dim wind_speed_selected_range As Range
    Dim new_datetime_selected_range As Range
    Dim wind_speed_average_selected_range As Range
    Dim wind_speed_direction_selected_range As Range
    Dim type_selection As String
    Dim write_uv_comp_flag As Boolean
    
    On Error Resume Next ' In case the Value in RefEdit is not a valid range
    Set datetime_selected_range = Range(DatetimeData.Value)
    Set wind_speed_selected_range = Range(targetColData.Value)
    Set new_datetime_selected_range = Range(newDatetimeData.Value)
    Set wind_speed_average_selected_range = Range(avgWindSpeedData.Value)
    Set wind_speed_direction_selected_range = Range(windSpeedDirectionData.Value)
    type_selection = typeSelectionCBox.Value
    write_uv_comp_flag = writeUVCheckBox.Value
    On Error GoTo 0 ' Stop error handling
    
    ' TODO (Khalid): Validate each data
    If Not datetime_selected_range Is Nothing Then
        ' Call the function and pass the selected range to it
            process_selected_range datetime_selected_range, wind_speed_selected_range, wind_speed_direction_selected_range, _
                                    new_datetime_selected_range, wind_speed_average_selected_range, type_selection, _
                                    write_uv_comp_flag
    Else
        MsgBox "The selected range is not valid."
    End If
    
    Unload Me ' Close the UserForm
End Sub
