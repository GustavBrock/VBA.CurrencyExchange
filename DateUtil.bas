Attribute VB_Name = "DateUtil"
Option Compare Text
Option Explicit

' DateUtil v1.0.1
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Timezone
'
' Selected constants, enums, and functions from project VBA.Date.
' If VBA.Timezone is used combined with VBA.Date, this module is
' superfluous and must be omitted.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)


' Code constants.
'
    Public Const ZeroDateValue      As Date = #12:00:00 AM#
    Public Const HoursPerDay        As Long = 24
    Public Const MinutesPerHour     As Long = 60
    Public Const SecondsPerMinute   As Long = 60
    Public Const SecondsPerHour     As Long = MinutesPerHour * SecondsPerMinute
    Public Const SecondsPerDay      As Long = HoursPerDay * SecondsPerHour
    Public Const DaysInWeek         As Integer = 7
    
    ' Unix Time.
    Public Const UtOffset           As Long = -25569
    
' Enums.
'
    ' Enum for error values for use with Err.Raise.
    Public Enum DtError
        dtInvalidProcedureCallOrArgument = 5
        dtOverflow = 6
        dtTypeMismatch = 13
    End Enum

' Returns the date of a specified Unix Time with a resolution of 1 ms.
' UnixDate can be any value that will return a valid VBA Date value.
'
' Minimum value:  -59011459200
'   ->  100-01-01 00:00:00.000
' Maximum value:  253402300799.999
'   -> 9999-12-31 23:59:59.999
'
' 2017-11-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateUnix( _
    ByVal UnixDate As Variant) _
    As Date
    
    Dim ResultDate  As Date
    
    ResultDate = (CDec(UnixDate) / SecondsPerDay) - CDec(UtOffset)
    
    DateUnix = ResultDate
    
End Function

' Returns the time of a specified Unix Time with a resolution of 1 ms.
' UnixTime can be any value that will return a valid VBA Date value.
'
' Zero value   :             0
'   ->            00:00:00.000
' Minimum value:  -56802297600
'   ->  100-01-01 00:00:00.000
' Maximum value:  255611462399.999
'   -> 9999-12-31 23:59:59.999
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function TimeUnix( _
    ByVal UnixTime As Variant) _
    As Date
    
    Dim ResultTime  As Date
    
    ResultTime = CDec(UnixTime) / SecondsPerDay
    
    TimeUnix = ResultTime
    
End Function

' Returns the Unix Time in seconds for a specified date.
' UtcDate can be any Date value of VBA with a resolution of one millisecond.
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function UnixDate( _
    ByVal UtcDate As Date) _
    As Variant
    
    Dim Result  As Variant
    
    Result = Int((CDec(UtcDate) + CDec(UtOffset)) * CDec(SecondsPerDay) + 0.5)
    
    UnixDate = Result
    
End Function

' Returns the Unix Time in seconds for a specified date.
' UtcTime can be any Date value of VBA with a resolution of one millisecond.
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function UnixTime( _
    ByVal UtcTime As Date) _
    As Variant
    
    Dim Result  As Variant
    
    Result = Int(CDec(UtcTime) * CDec(SecondsPerDay) + 0.5)
    
    UnixTime = Result
    
End Function

' Calculates the date of the occurrence of Weekday in the month of DateInMonth.
'
' If Occurrence is 0 or negative, the first occurrence of Weekday in the month is assumed.
' If Occurrence is 5 or larger, the last occurrence of Weekday in the month is assumed.
'
' If Weekday is invalid or not specified, the weekday of DateInMonth is used.
'
' 2016-06-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateWeekdayInMonth( _
    ByVal DateInMonth As Date, _
    Optional ByVal Occurrence As Integer, _
    Optional ByVal Weekday As VbDayOfWeek = vbUseSystemDayOfWeek) _
    As Date
    
    Dim Offset          As Integer
    Dim Month           As Integer
    Dim Year            As Integer
    Dim ResultDate      As Date
    
    ' Validate Weekday.
    Select Case Weekday
        Case _
            vbMonday, _
            vbTuesday, _
            vbWednesday, _
            vbThursday, _
            vbFriday, _
            vbSaturday, _
            vbSunday
        Case Else
            ' vbUseSystemDayOfWeek, zero, none or invalid value for VbDayOfWeek.
            Weekday = VBA.Weekday(DateInMonth)
    End Select
    
    ' Validate Occurence.
    If Occurrence <= 0 Then
        ' Find first occurrence.
        Occurrence = 1
    ElseIf Occurrence > 5 Then
        ' Find last occurrence.
        Occurrence = 5
    End If
    
    ' Start date.
    Month = VBA.Month(DateInMonth)
    Year = VBA.Year(DateInMonth)
    ResultDate = DateSerial(Year, Month, 1)
    
    ' Find offset of Weekday from first day of month.
    Offset = DaysInWeek * (Occurrence - 1) + (Weekday - VBA.Weekday(ResultDate) + DaysInWeek) Mod DaysInWeek
    ' Calculate result date.
    ResultDate = DateAdd("d", Offset, ResultDate)
    
    If Occurrence = 5 Then
        ' The latest occurrency of Weekday is requested.
        ' Check if there really is a fifth occurrence of Weekday in this month.
        If VBA.Month(ResultDate) <> Month Then
            ' There are only four occurrencies of Weekday in this month.
            ' Return the fourth as the latest.
            ResultDate = DateAdd("d", -DaysInWeek, ResultDate)
        End If
    End If
    
    DateWeekdayInMonth = ResultDate
  
End Function

' Returns the sign of Expression, +, -, for positive or negative
' values, or a space (or a plus sign) for zero or a non-numeric value.
'
' If ZeroPlus is True, + will be returned also for zero values.
'
' Examples:
'   0.78    -> "+"
'   "-23,9" -> "-"
'   Null    -> " "
'   Date()  -> " "
'   "Yes"   -> " "
'
' 2016-07-05. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function FormatSign( _
    ByVal Expression As Variant, _
    Optional ByVal ZeroPlus As Boolean) _
    As String
    
    Const Signs As String = "- +"
    
    ' Always return exactly one character.
    Dim Sign    As String * 1
    
    If IsNumeric(Expression) Or IsDate(Expression) Then
        Sign = Mid(Signs, 2 + Sgn(Expression))
        If ZeroPlus = True And Sign = " " Then
            Sign = Mid(Signs, 3)
        End If
    End If
    
    FormatSign = Sign

End Function



