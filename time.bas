Attribute VB_Name = "Module1"

Public Enum TimeFormatType
DaysHoursMinutesSecondsMilliseconds = 0
DaysHoursMinutesSeconds = 1
DHMSMColonSeparated = 2
DaysHoursMinutes = 3
End Enum

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

Public Function FormatCount(Count As Long, Optional FormatType As TimeFormatType = 0) As String
    Dim Hours As Long, Minutes As Long, Seconds As Long
    
    Miliseconds = Count Mod 1000
    Count = Count \ 1000
    Days = Count \ (24& * 3600&)
    If Days > 0 Then Count = Count - (24& * 3600& * Days)
    Hours = Count \ 3600&
    If Hours > 0 Then Count = Count - (3600& * Hours)
    Minutes = Count \ 60
    Seconds = Count Mod 60


    Select Case FormatType
            Case 0


        FormatCount = Hours & " hours, " & _
            Minutes & " minutes, " & Seconds & " seconds "
            
            Case 1


            FormatCount = Hours & " hours, " & _
                Minutes & " minutes, " & Seconds & " seconds"
            Case 2


                FormatCount = Hours & ":" & _
                    Minutes & ":" & Seconds
            
            Case 3
            
            FormatCount = Hours & " hours, " & _
                Minutes & " minutes"
End Select
End Function


