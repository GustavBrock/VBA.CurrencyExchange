Attribute VB_Name = "ZoneUtil"
Option Compare Database
Option Explicit

' ZoneUtil v1.0.0
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Timezone
'
' Selected constants, enums, and functions from project VBA.Timezone.
' If VBA.CurrencyConvert is used combined with VBA.Timezone, this module is
' superfluous and must be omitted.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)

    Public Type SystemTime
        wYear                           As Integer
        wMonth                          As Integer
        wDayOfWeek                      As Integer
        wDay                            As Integer
        wHour                           As Integer
        wMinute                         As Integer
        wSecond                         As Integer
        wMilliseconds                   As Integer
    End Type

' Declarations.

' Returns the current UTC time.
Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" ( _
    ByRef lpSystemTime As SystemTime)
'

' Retrieves the current date and time from the local computer as UTC.
' By cutting off the milliseconds, the resolution is one second to mimic Now().
'
' 2016-06-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function UtcNow() As Date

    Dim SysTime     As SystemTime
    Dim Datetime    As Date
    
    ' Retrieve current UTC date/time.
    GetSystemTime SysTime
    
    Datetime = _
        DateSerial(SysTime.wYear, SysTime.wMonth, SysTime.wDay) + _
        TimeSerial(SysTime.wHour, SysTime.wMinute, SysTime.wSecond)
    
    UtcNow = Datetime
    
End Function

