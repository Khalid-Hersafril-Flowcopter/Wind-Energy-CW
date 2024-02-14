VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InformationExtract 
   Caption         =   "UserForm1"
   ClientHeight    =   3300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6075
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
    Dim u_comp_selected_range As Range
    Dim v_comp_selected_range As Range
    Dim target_col_selected_range As Range
    
    On Error Resume Next ' In case the Value in RefEdit is not a valid range
    Set datetime_selected_range = Range(DatetimeData.Value)
    Set u_comp_selected_range = Range(uCompData.Value)
    Set v_comp_selected_range = Range(vCompData.Value)
    Set target_col_selected_range = Range(targetColData.Value)
    On Error GoTo 0 ' Stop error handling
    
    ' TODO (Khalid): Validate each data
    If Not datetime_selected_range Is Nothing Then
        ' Call the function and pass the selected range to it
            process_selected_range datetime_selected_range, u_comp_selected_range, v_comp_selected_range, target_col_selected_range
    Else
        MsgBox "The selected range is not valid."
    End If
    
    Unload Me ' Close the UserForm
End Sub
