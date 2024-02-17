Attribute VB_Name = "WindDataProcess"
Public Function process_selected_range(ByRef datetimeRange As Range, ByRef windSpeedRange As Range, ByRef newDatetimeRange As Range, ByRef windSpeedAvgRange As Range)
    ' Process the range however you need
    ' For example, just print out the address to the Immediate Window
    Debug.Print datetimeRange.Address
    Debug.Print windSpeedRange.Address
    Debug.Print newDatetimeRange.Address
    Debug.Print windSpeedAvgRange.Address

    Dim datetimeValues As Variant: datetimeValues = datetimeRange.Value
    Dim windSpeedValues As Variant: windSpeedValues = windSpeedRange.Value

    ' Get header name
    ' It is assumed that the header is at the first row of the column
    Dim datetime_header As String: datetime_header = datetimeValues(1, 1)
    Dim wind_speed_header As String: wind_speed_header = windSpeedValues(1, 1)
    
    ' Get the column letter for where to write the data
    Dim datetimeCol As String: datetimeCol = getColumnLetter(datetimeRange.Address)
    Dim windSpeedCol As String: windSpeedCol = getColumnLetter(windSpeedRange.Address)
    
    ' Find the last row with data in column N
'    lastRow = ws.Cells(ws.Rows.Count, datetimeCol).End(xlUp).Row
    
'    Dim lastUsedRow As Long
'    With datetimeRange
'        ' Find the last row with data starting from the bottom of the selected column range
'        lastUsedRow = .Cells(.Rows.Count, 1).End(xlUp).Row - .Row + 1
'    End With
    
    Dim datetime_len As Long: datetime_len = getRowLen(datetimeRange)
    Dim wind_speed_len As Long: wind_speed_len = getRowLen(windSpeedRange)
    
    Debug.Print "Data length: " & datetime_len & ", " & wind_speed_len
    
    ' Get the column letter for where to write the data
    Dim dateWriteCol As String: dateWriteCol = getColumnLetter(newDatetimeRange.Address)
    Dim windSpeedAverageWriteCol As String: windSpeedAverageWriteCol = getColumnLetter(windSpeedAvgRange.Address)
    
    ' Write the header for the new generated data
    Range(dateWriteCol & 1) = "Date and Time"
    Range(windSpeedAverageWriteCol & 1) = "Wind Speed Average (m/s)"
    
    Dim wind_speed_dict As Object
    Set wind_speed_dict = CreateObject("Scripting.Dictionary")
    
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

    ' Initialize variables required for data processing
    Dim hour_key As Integer
    Dim wind_speed_sum As Double: wind_speed_sum = 0
    Dim data_count As Integer: data_count = 0

    Dim i As Long
    Dim date_val As Date
    Dim prev_date_val As Date: prev_date_val = datetimeValues(2, 1)

    For i = 2 To datetime_len
        
        ' TODO (Khalid): Currently, there is a bug where dd/mm/yyyy 00:00:00 returns Nothing and this is not captured
        ' We can choose to ignore this value, but that'd mean for 0 hours, you'd get one less elements
        If IsDate(datetimeValues(i, 1)) Then
            date_val = datetimeValues(i, 1)
            Dim curr_date As String: curr_date = getOnlyDate(date_val)
            Dim curr_hour As Integer: curr_hour = hour(date_val)
            
            If IsNumeric(windSpeedValues(i, 1)) Then
                
                Dim curr_wind_speed As Double: curr_wind_speed = windSpeedValues(i, 1)
            
                day_changed = DateDiff("d", curr_date, prev_date)
                hour_changed = curr_hour <> prev_hour
                
                If Not hour_changed Then
                    wind_speed_sum = wind_speed_sum + curr_wind_speed
                    data_count = data_count + 1
                    
                    If i = datetime_len Then
                        wind_speed_average = wind_speed_sum / data_count
                        wind_speed_dict.Add new_datetime, wind_speed_average
                        Range(dateWriteCol & wind_speed_dict.Count + 1) = CDate(new_datetime)
                        Range(windSpeedAverageWriteCol & wind_speed_dict.Count + 1) = wind_speed_average
                        Debug.Print new_datetime, wind_speed_dict(new_datetime), i
                        Debug.Print "Finished processing the wind speed date."
                    End If
    
                Else
                    
                    ' Ensure that all data receieved is not corrupted before doing average calculations
                    If IsNumeric(wind_speed_sum) And data_count <> 0 Then
                        wind_speed_average = wind_speed_sum / data_count
                        wind_speed_dict.Add new_datetime, wind_speed_average
                        Range(dateWriteCol & wind_speed_dict.Count + 1) = CDate(new_datetime)
                        Range(windSpeedAverageWriteCol & wind_speed_dict.Count + 1) = wind_speed_average
                        
                        Debug.Print new_datetime, wind_speed_dict(new_datetime), i
                    Else
                        ' This condition generallly occurs if the hour interval keeps returning faulty data
                        Debug.Print "Average data is corrupted with " & wind_speed_sum & " Returning NaN!"
                        wind_speed_dict.Add new_datetime, "NaN"
                        Range(dateWriteCol & wind_speed_dict.Count + 1) = CDate(new_datetime)
                        Range(windSpeedAverageWriteCol & wind_speed_dict.Count + 1) = "NaN"
                    End If
                    
'                    Dim missing_hours As Long: missing_hours = curr_hour - prev_hour
'
'                    Dim missing_hours As Long: missing_hours = HoursDifference(CDate(curr_date), CDate(prev_date))
'
'                    ' Handles missing datetime. This is to ensure that the data length is equal to each other
'                    ' which is extremely important when performing correlation
'                    If missing_hours > 1 Then
'                        For n = 1 To missing_hours - 1
'                            Dim missing_datetime As Date: missing_datetime = generateNewDatetime(curr_date, prev_hour + n)
'                            wind_speed_dict.Add missing_datetime, "NaN"
'                            Range(dateWriteCol & wind_speed_dict.Count + 1) = CDate(missing_datetime)
'                            Range(windSpeedAverageWriteCol & wind_speed_dict.Count + 1) = "NaN"
'                        Next n
'                    End If
                    
                    Dim hours_diff As Double: hours_diff = hoursDifference(CDate(prev_date_val), CDate(date_val))
                    
                    Debug.Print "Difference in hours: " & hours_diff
                    
                    ' Loop through dates from StartDate to EndDate with 1-hour increments
                    Do While Round(hours_diff, 0) > 1
                        ' Increment CurrentDate by 1 hour
                        prev_date_val = DateAdd("h", 1, prev_date_val)
                        ' Print each hour in 24-hour format
                        Debug.Print Format(prev_date_val, "dd/mm/yyyy HH:mm:ss")
                        
                        wind_speed_dict.Add prev_date_val, "NaN"
                        Range(dateWriteCol & wind_speed_dict.Count + 1) = CDate(prev_date_val)
                        Range(windSpeedAverageWriteCol & wind_speed_dict.Count + 1) = "NaN"
                        hours_diff = hours_diff - 1
                        
                    Loop

                    ' Since now we are in the next hour, we should populate the first sum as the current wind speed value
                    wind_speed_sum = curr_wind_speed
                    data_count = 1
                    new_datetime = generateNewDatetime(curr_date, curr_hour)
                End If
          
                prev_hour = curr_hour
                prev_date = curr_date
            Else
                ' Ignore and skip to the next iteration if the data is corrupted
                Debug.Print "Data is corrupted on " & date_val & " with " & windSpeedValues(i, 1); ". Ignoring this data!"
            End If
            
            prev_date_val = date_val
        End If
    
    Next i
    
End Function

Private Function getRowLen(rng As Range)
    ' TODO (Khalid): Bug - If the user does not select the whole column, it returns 1
    Dim last_used_row As Long
    With rng
        last_used_row = .Cells(.Rows.Count, 1).End(xlUp).Row - .Row + 1
    End With
    
    getRowLen = last_used_row
End Function

Private Function getOnlyDate(datetime As Date)
    Dim output_date As String
    output_date = Day(datetime) & "/" & Month(datetime) & "/" & Year(datetime)
    getOnlyDate = output_date
End Function

Function getColumnLetter(columnRef As String) As String
    ' Replace removes the $ signs, and then we take the first character since the column is the same.
    getColumnLetter = Replace(columnRef, "$", "")
    ' If there's a colon indicating a range of the same column, it is removed as well.
    If InStr(getColumnLetter, ":") > 0 Then
        getColumnLetter = Left(getColumnLetter, InStr(getColumnLetter, ":") - 1)
    End If
End Function

Private Function generateNewDatetime(curr_date As String, hour_int As Integer)
    Dim time_str As String
    Dim new_datetime_fmt As String
    
    ' Force it to be formatted with 00:00:00 clock
    If hour_int = 0 Then
        time_str = "00" & ":" & "00:00"
        new_datetime_fmt = curr_date & " " & time_str
    Else
        time_str = hour_int & ":" & "00:00"
        new_datetime_fmt = curr_date & " " & time_str
    End If
    
    generateNewDatetime = CDate(new_datetime_fmt)
End Function

Function hoursDifference(StartDate As Date, EndDate As Date) As Double
    hoursDifference = (EndDate - StartDate) * 24
End Function

