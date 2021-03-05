
Function buildFileName()
    Dim currentDate, currentTime
    currentDate = Date
    currentTime = Time

    buildFileName = Year(currentDate) & "." & Month(currentDate) & "." & Day(currentDate) & "-" & Hour(currentTime) & "." & Minute(currentTime) & "." & Second(currentTime)
End Function

MsgBox(buildFileName())