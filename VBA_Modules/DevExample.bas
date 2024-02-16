Attribute VB_Name = "DevExample"
Function test()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("VBA Development") ' Replace "SheetName" with the actual name of your sheet
    Dim lastRow As Long
    Dim dataRange As Range
    Dim windSpeedRange As Range
'    Dim vRange As Range

    ' Find the last row with data in column N
    lastRow = ws.Cells(ws.Rows.Count, "N").End(xlUp).Row

    ' Set the range for the entire column N up to the last row with data
    Set dataRange = ws.Range("N1:N" & lastRow)
    Set windSpeedRange = ws.Range("O1:O" & lastRow)
'    Set vRange = ws.Range("P1:P" & lastRow)
    
    Dim datetimeValues As Variant
    Dim windSpeedValues As Variant
'    Dim vCompValues As Variant
    
    datetimeValues = dataRange.Value
    windSpeedValues = windSpeedRange.Value
'    vCompValues = vRange.Value
    
    Dim wind_speed_dict As Object
    Set wind_speed_dict = CreateObject("Scripting.Dictionary")
    
    Dim hour_key As Integer
    Dim wind_speed_sum As Double
    Dim data_count As Integer
    
    ' Initialize the new datetime with the first date available from the range
    Dim curr_datetime As Date: curr_datetime = datetimeValues(2, 1)
    Dim init_date As String: init_date = getOnlyDate(curr_datetime)
    Dim init_hour_int As Integer: init_hour_int = 0
    Dim new_datetime As String: new_datetime = generateNewDatetime(init_date, init_hour_int)
    
    ' Initialize control variable for checking change in conditions
    Dim prev_hour As Integer: prev_hour = init_hour_int
    Dim prev_date As String: prev_date = init_date
    Dim hour_changed As Boolean: hour_changed = False
    Dim day_changed As Boolean: day_changed = False

    Dim i As Long
    
    Dim date_val As Date
    
    ' Initialize day change
    
    wind_speed_sum = 0
    data_count = 0
    
    For i = 2 To lastRow

        date_val = datetimeValues(i, 1)

        ' TODO (Khalid): Currently, there is a bug where dd/mm/yyyy 00:00:00 returns Nothing and this is not captured
        ' We can choose to ignore this value, but that'd mean for 0 hours, you'd get one less elements
        If IsDate(date_val) Then
            Dim curr_date As String: curr_date = getOnlyDate(date_val)
            Dim curr_hour As Integer: curr_hour = hour(date_val)
            
            If IsNumeric(windSpeedValues(i, 1)) Then
                
                Dim curr_wind_speed As Double: curr_wind_speed = windSpeedValues(i, 1)

            
'            Debug.Print curr_date, curr_hour, curr_wind_speed
            
                day_changed = DateDiff("d", curr_date, prev_date)
                hour_changed = curr_hour <> prev_hour
                
                If Not hour_changed Then
                    wind_speed_sum = wind_speed_sum + curr_wind_speed
                    data_count = data_count + 1
                    
                    If i = lastRow Then
                        wind_speed_average = wind_speed_sum / data_count
                        wind_speed_dict.Add new_datetime, wind_speed_average
                        Debug.Print new_datetime, wind_speed_dict(new_datetime), i
                        Debug.Print "Finished processing the wind speed date."
                    End If
    
                Else
                    
                    ' Ensure that all data receieved is not corrupted before doing average calculations
                    If IsNumeric(wind_speed_sum) And data_count <> 0 Then
                        wind_speed_average = wind_speed_sum / data_count
                        wind_speed_dict.Add new_datetime, wind_speed_average
                        
                        Debug.Print new_datetime, wind_speed_dict(new_datetime), i
                    Else
                        Debug.Print "Average data is corrupted with " & wind_speed_sum & " Returning NaN!"
                        wind_speed_dict.Add new_datetime, "NaN"
                    End If

                    ' Since now we are in the next hour, we should populate the first sum as the current wind speed value
                    wind_speed_sum = curr_wind_speed
                    data_count = 1
                    new_datetime = generateNewDatetime(curr_date, curr_hour)
                End If
    
    '            If day_changed Then
    '                wind_speed_dict.Add new_datetime, Array(u_avg, v_avg)
    '                new_datetime = generateNewDatetime(curr_date, curr_hour)
    '            End If
                
                prev_hour = curr_hour
                prev_date = curr_date
            Else
                ' Ignore and skip to the next iteration if the data is corrupted
                Debug.Print "Data is corrupted on " & date_val & " with " & windSpeedValues(i, 1); ". Ignoring this data!"
            End If
        End If
    
    Next i
    
End Function

Private Function getOnlyDate(datetime As Date)
    Dim output_date As String
    output_date = Day(datetime) & "/" & Month(datetime) & "/" & Year(datetime)
    getOnlyDate = output_date
End Function

Private Function getRowLen(rng As Range)
    ' TODO (Khalid): Bug - If the user does not select the whole column, it returns 1
    Dim last_used_row As Long
    With rng
        last_used_row = .Cells(.Rows.Count, 1).End(xlUp).Row - .Row + 1
    End With
    
    getRowLen = last_used_row
End Function

Private Function generateNewDatetime(curr_date As String, hour_int As Integer)
    Dim time_str As String
    Dim new_datetime_fmt As String
    
    ' Force it to be formatted with 00:00:00 clock
    If hour_int = 0 Then
        time_str = "00" & ":" & "00:00"
        new_datetime_fmt = curr_date & " " & time_str
        generateNewDatetime = new_datetime_fmt
    Else
        time_str = hour_int & ":" & "00:00"
        new_datetime_fmt = curr_date & " " & time_str
        generateNewDatetime = CDate(new_datetime_fmt)
    End If
End Function
