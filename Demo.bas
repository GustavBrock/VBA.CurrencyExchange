Attribute VB_Name = "Demo"
Option Compare Text
Option Explicit

' CurrencyExchange Demo V1.0.0
' (c) Gustav Brock, Cactus Data ApS, CPH


' Fill table CurrencyRate with exchange rates from a source of choice.
'
' Example:
'
'   FillCurrencyRates ExchangeRatesDkk
'
' Note, that some sources don't supply the currency name, only the code.
'
' 2018-10-07. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub FillCurrencyRates(ByRef Rates As Variant)

    Const TableName     As String = "CurrencyRate"
    
    Dim Records         As DAO.Recordset
    
    Dim FieldNames      As Variant
    Dim Sql             As String
    Dim Index           As Integer
    Dim Item            As Integer
    
    If Not IsArray(Rates) Then Exit Sub
    
    ' Field names must match the order of array Rates.
    FieldNames = Array("[Date]", "[Code]", "[Rate]", "[Name]")
    
    ' Clean table.
    Sql = "Delete * From " & TableName & ";"
    CurrentDb.Execute Sql
    
    ' Fill table.
    Sql = "Select " & Join(FieldNames, ",") & " From " & TableName & ";"
    Set Records = CurrentDb.OpenRecordset(Sql)
    For Index = LBound(Rates, 1) To UBound(Rates, 1)
        Records.AddNew
        For Item = LBound(Rates, 2) To UBound(Rates, 2)
            Records.Fields(Item).Value = Rates(Index, Item)
        Next
        Records.Update
    Next
    Records.Close
    
End Sub

