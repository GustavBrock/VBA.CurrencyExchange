Attribute VB_Name = "ExchangeService"
Option Compare Text
Option Explicit

' ExchangeRate V1.5.1
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.CurrencyExchange


' API id or key. Guid string, 0, 24, 26, or 32 characters.
'
' Currency Converter API:           "00000000-0000-0000-0000-000000000000"
' Leave empty for the free plan:    ""
Public Const CcaApiId   As String = "00000000-0000-0000-0000-000000000000"
' apilayer/currencylayer:           "--------------------------------"
Public Const ApiApiId   As String = "00000000000000000000000000000000"
' ExchangeRate-API:                 "------------------------"
Public Const EraApiId   As String = "000000000000000000000000"
' Fixer:                            "--------------------------------"
Public Const FxrApiId   As String = "00000000000000000000000000000000"
' Open Exchange Rates:              "--------------------------------"
Public Const OxrApiId   As String = "00000000000000000000000000000000"
' XE Account ID, like:              "organisationname0000000"
Public Const XeAccount  As String = "aaaaaaaaaaaaa1234567"
' XE Account API Key:               "--------------------------"
Public Const XeApiId    As String = "00000000000000000000000000"

' Compiler constants.
'
' Select Early Binding (True) or Late Binding (False).
#Const EarlyBinding = True

' Enums.
'
' HTTP status codes, reduced.
Private Enum HttpStatus
    OK = 200
    BadRequest = 400
    Unauthorized = 401
    Forbidden = 403
End Enum
'
' Dimensions of array holding rates.
Private Enum RateDetail
    Date = 0
    Code = 1
    Rate = 2
    Name = 3
End Enum
'
' Dimensions of array holding parameters.
Private Enum ParameterDetail
    Name = 0
    Value = 1
End Enum

' Application constants.
'
' Currency code for Danish krone.
Public Const DanishKroneCode    As String = "DKK"
' Currency code for Euro.
Public Const EuroCode           As String = "EUR"
' Currency code for US Dollar.
Public Const USDollarCode       As String = "USD"
' Currency code for neutral currency.
Public Const RubelCode          As String = "RUB"
' Currency code for neutral currency.
Public Const NeutralCode        As String = "XXX"
' Currency name for neutral currency.
Public Const NeutralName        As String = "No currency"
' Exchange rate for no currency.
Public Const NeutralRate        As Double = 1
' Currency code for no (invalid) currency.
Public Const NoCode             As String = ""
' Exchange rate for no (invalid) currency.
Public Const NoRate             As Double = 0
' Publishing/value date when unknown.
Public Const NoValueDate        As Date = #1/1/1970#

' Returns the current conversion factor from Rubel to another currency based on
' the official exchange rates published by the Central Bank of the Russian
' Federation.
'
' Optionally, the conversion factor can be calculated from any other of the
' published exchange rates. Exchange rates from or to other currencies than
' RUB are calculated from RUB by triangular calculation.
'
' If an invalid or unpublished currency code is passed, a conversion factor
' of zero is returned.
'
' Examples, typical:
'   CurrencyConvertCbr("DKK")           ->  0.0973738278625471
'   CurrencyConvertCbr("DKK", "EUR")    ->  7.46477501777072
'   CurrencyConvertCbr("AUD")           ->  0.021253081696846
'   CurrencyConvertCbr("AUD", "DKK")    ->  0.2182627731021
'   CurrencyConvertCbr("DKK", "AUD")    ->  4.58163334858857
'   CurrencyConvertCbr("EUR", "DKK")    ->  0.133962510272498
'   CurrencyConvertCbr("", "DKK")       -> 10.2697
'   CurrencyConvertCbr("EUR")           ->  0.013044442415309
' Examples, neutral code.
'   CurrencyConvertCbr("AUD", "XXX")    ->  1
'   CurrencyConvertCbr("XXX", "AUD")    ->  1
'   CurrencyConvertCbr("XXX")           ->  1
' Examples, invalid code.
'   CurrencyConvertCbr("XYZ")           ->  0
'   CurrencyConvertCbr("DKK", "XYZ")    ->  0
'
' 2018-10-07. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function CurrencyConvertCbr( _
    ByVal IsoTo As String, _
    Optional ByVal IsoFrom As String = RubelCode) _
    As Double
    
    Dim Rates()     As Variant
    
    Dim RateTo      As Double
    Dim RateFrom    As Double
    Dim Factor      As Double
    Dim Index       As Integer
    
    If IsoFrom = "" Then
        IsoFrom = RubelCode
    End If
    If IsoTo = "" Then
        IsoTo = RubelCode
    End If
    
    If IsoTo = NeutralCode Or IsoFrom = NeutralCode Then
        Factor = NeutralRate
    ElseIf IsoTo = IsoFrom Then
        Factor = NeutralRate
    Else
        Rates() = ExchangeRatesCbr
    
        If IsoTo = RubelCode Then
            RateTo = NeutralRate
        Else
            For Index = LBound(Rates) To UBound(Rates)
                If Rates(Index, RateDetail.Code) = IsoTo Then
                    RateTo = Rates(Index, RateDetail.Rate)
                    Exit For
                End If
            Next
        End If
        
        If RateTo > NoRate Then
            If IsoFrom = RubelCode Then
                RateFrom = NeutralRate
            Else
                For Index = LBound(Rates) To UBound(Rates)
                    If Rates(Index, RateDetail.Code) = IsoFrom Then
                        RateFrom = Rates(Index, RateDetail.Rate)
                        Exit For
                    End If
                Next
            End If
            Factor = RateFrom / RateTo
        End If
        
    End If
    
    CurrencyConvertCbr = Factor

End Function

' Returns the current conversion factor from one currency to another
' based on the exchange rates published by "Currency Converter API".
' By default, conversion is from Euro to another currency.
'
' If an invalid or unpublished currency code is passed, a conversion factor
' of zero is returned.
'
' Examples, typical:
'   CurrencyConvertCca("DKK")           ->  7.47139
'   CurrencyConvertCca("DKK", "EUR")    ->  7.47139
'   CurrencyConvertCca("AUD")           ->  1.61313
'   CurrencyConvertCca("AUD", "DKK")    ->  0.215908
'   CurrencyConvertCca("DKK", "AUD")    ->  4.63161
'   CurrencyConvertCca("EUR", "DKK")    ->  0.133844
'   CurrencyConvertCca("", "DKK")       ->  0.157527
'   CurrencyConvertCca("USD")           ->  1.176948
' Examples, neutral code.
'   CurrencyConvertCca("AUD", "XXX")    ->  1
'   CurrencyConvertCca("XXX", "AUD")    ->  1
'   CurrencyConvertCca("XXX")           ->  1
' Examples, invalid code.
'   CurrencyConvertCca("XYZ")           ->  0
'   CurrencyConvertCca("DKK", "XYZ")    ->  0
'
' 2018-09-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function CurrencyConvertCca( _
    ByVal IsoTo As String, _
    Optional ByVal IsoFrom As String = EuroCode) _
    As Double
    
    Dim Rates()     As Variant
    
    Dim IsoBase     As String
    Dim Factor      As Double
    
    If IsoFrom = "" Then
        IsoFrom = EuroCode
    End If
    If IsoTo = "" Then
        IsoTo = USDollarCode
    End If
    
    If IsoTo = NeutralCode Or IsoFrom = NeutralCode Then
        Factor = NeutralRate
    ElseIf IsoTo = IsoFrom Then
        Factor = NeutralRate
    Else
        ' Retrieve the current rate.
        IsoBase = IsoFrom
        Rates() = ExchangeRatesCca(IsoBase, IsoTo)
        Factor = Rates(0, RateDetail.Rate)
    End If
    
    CurrencyConvertCca = Factor

End Function

' Returns the current conversion factor from US Dollar to another currency
' based on the exchange rates published by "Currencylayer API".
'
' Optionally, the conversion factor can be calculated from any other of the
' published exchange rates.
' For the free plan, exchange rates for other base currencies are
' calculated from USD by triangular calculation.
'
' If an invalid or unpublished currency code is passed, a conversion factor
' of zero is returned.
'
' Examples, typical:
'   CurrencyConvertCla("DKK")           ->  6.3456
'   CurrencyConvertCla("DKK", "EUR")    ->  7.45922499573883
'   CurrencyConvertCla("AUD")           ->  1.37655
'   CurrencyConvertCla("AUD", "DKK")    ->  0.216929841149773
'   CurrencyConvertCla("DKK", "AUD")    ->  4.60978533289746
'   CurrencyConvertCla("EUR", "DKK")    ->  0.134062184820978
'   CurrencyConvertCla("", "DKK")       ->  0.157589510842158
'   CurrencyConvertCla("USD")           ->  1
' Examples, neutral code.
'   CurrencyConvertCla("AUD", "XXX")    ->  1
'   CurrencyConvertCla("XXX", "AUD")    ->  1
'   CurrencyConvertCla("XXX")           ->  1
' Examples, invalid code.
'   CurrencyConvertCla("XYZ")           ->  0
'   CurrencyConvertCla("DKK", "XYZ")    ->  0
'
' 2018-09-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function CurrencyConvertCla( _
    ByVal IsoTo As String, _
    Optional ByVal IsoFrom As String = USDollarCode) _
    As Double
    
    Dim Rates()     As Variant
    
    Dim IsoBase     As String
    Dim RateTo      As Double
    Dim RateFrom    As Double
    Dim Factor      As Double
    Dim Index       As Integer
    
    If IsoFrom = "" Then
        IsoFrom = USDollarCode
    End If
    If IsoTo = "" Then
        IsoTo = USDollarCode
    End If
    
    If IsoTo = NeutralCode Or IsoFrom = NeutralCode Then
        Factor = NeutralRate
    ElseIf IsoTo = IsoFrom Then
        Factor = NeutralRate
    Else
        ' Retrieve current rates using IsoFrom as the base currency.
        ' If using the free license, USD is always the base currency, thus
        ' triangular calculation of the Factor will be used for other
        ' base currencies than USD.
        IsoBase = IsoFrom
        Rates() = ExchangeRatesCla(IsoBase)
        
        ' Look up the Factor of IsoBase.
        For Index = LBound(Rates) To UBound(Rates)
            If Rates(Index, RateDetail.Code) = IsoFrom Then
                RateFrom = Rates(Index, RateDetail.Rate)
                Exit For
            End If
        Next
        
        If RateFrom > NoRate Then
            ' Look up the Factor of IsoTo.
            For Index = LBound(Rates) To UBound(Rates)
                If Rates(Index, RateDetail.Code) = IsoTo Then
                    RateTo = Rates(Index, RateDetail.Rate)
                    Exit For
                End If
            Next
            Factor = RateTo / RateFrom
        End If
    End If
    
    CurrencyConvertCla = Factor

End Function

' Returns the current conversion factor from Danish Krone to another currency
' based on the official exchange rates published by the Danish National Bank.
'
' Optionally, the conversion factor can be calculated from any other of the
' published exchange rates. Exchange rates for other base currencies are
' calculated from DKK by triangular calculation.
'
' Source:
'   http://www.nationalbanken.dk/en/statistics/exchange_rates/Pages/default.aspx
'
' If an invalid or unpublished currency code is passed, a conversion factor
' of zero is returned.
'
' Examples, typical:
'   CurrencyConvertDkk("EUR")           ->  0.134062634062634
'   CurrencyConvertDkk("EUR", "DKK")    ->  0.134062634062634
'   CurrencyConvertDkk("AUD")           ->  0.21661901048436
'   CurrencyConvertDkk("AUD", "DKK")    ->  0.21661901048436
'   CurrencyConvertDkk("DKK", "AUD")    ->  4.6164
'   CurrencyConvertDkk("DKK", "EUR")    ->  7.4592
'   CurrencyConvertDkk("AUD", "EUR")    ->  1.61580452300494

'   CurrencyConvertDkk("", "EUR")       ->  7.4592
'   CurrencyConvertDkk("DKK")           ->  1
' Examples, neutral code.
'   CurrencyConvertDkk("AUD", "XXX")    ->  1
'   CurrencyConvertDkk("XXX", "AUD")    ->  1
'   CurrencyConvertDkk("XXX")           ->  1
' Examples, invalid code.
'   CurrencyConvertDkk("XYZ")           ->  0
'   CurrencyConvertDkk("EUR", "XYZ")    ->  0
'
' 2018-09-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function CurrencyConvertDkk( _
    ByVal IsoTo As String, _
    Optional ByVal IsoFrom As String = DanishKroneCode) _
    As Double
    
    Dim Rates()     As Variant
    
    Dim RateTo      As Double
    Dim RateFrom    As Double
    Dim Factor      As Double
    Dim Index       As Integer
    
    If IsoFrom = "" Then
        IsoFrom = DanishKroneCode
    End If
    If IsoTo = "" Then
        IsoTo = DanishKroneCode
    End If
    
    If IsoTo = NeutralCode Or IsoFrom = NeutralCode Then
        Factor = NeutralRate
    ElseIf IsoTo = IsoFrom Then
        Factor = NeutralRate
    Else
        Rates() = ExchangeRatesDkk
    
        If IsoTo = DanishKroneCode Then
            RateTo = NeutralRate
        Else
            For Index = LBound(Rates) To UBound(Rates)
                If Rates(Index, RateDetail.Code) = IsoTo Then
                    RateTo = Rates(Index, RateDetail.Rate)
                    Exit For
                End If
            Next
        End If
        
        If RateTo > NoRate Then
            If IsoFrom = DanishKroneCode Then
                RateFrom = NeutralRate
            Else
                For Index = LBound(Rates) To UBound(Rates)
                    If Rates(Index, RateDetail.Code) = IsoFrom Then
                        RateFrom = Rates(Index, RateDetail.Rate)
                        Exit For
                    End If
                Next
            End If
            Factor = RateFrom / RateTo
        End If
        
    End If
    
    CurrencyConvertDkk = Factor

End Function

' Returns the current conversion factor from Euro to another currency
' based on the official exchange rates published by the European Central Bank.
'
' Optionally, the conversion factor can be calculated from any other of the
' published exchange rates. Exchange rates from or to other currencies than
' EUR are calculated from EUR by triangular calculation.
'
' If an invalid or unpublished currency code is passed, a conversion factor
' of zero is returned.
'
' Examples, typical:
'   CurrencyConvertEcb("DKK")           ->  7.4592
'   CurrencyConvertEcb("DKK", "EUR")    ->  7.4592
'   CurrencyConvertEcb("AUD")           ->  1.6158
'   CurrencyConvertEcb("AUD", "DKK")    ->  0.216618404118404
'   CurrencyConvertEcb("DKK", "AUD")    ->  4.61641292239139
'   CurrencyConvertEcb("EUR", "DKK")    ->  0.134062634062634
'   CurrencyConvertEcb("", "DKK")       ->  0.134062634062634
'   CurrencyConvertEcb("EUR")           ->  1
' Examples, neutral code.
'   CurrencyConvertEcb("AUD", "XXX")    ->  1
'   CurrencyConvertEcb("XXX", "AUD")    ->  1
'   CurrencyConvertEcb("XXX")           ->  1
' Examples, invalid code.
'   CurrencyConvertEcb("XYZ")           ->  0
'   CurrencyConvertEcb("DKK", "XYZ")    ->  0
'
' 2018-09-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function CurrencyConvertEcb( _
    ByVal IsoTo As String, _
    Optional ByVal IsoFrom As String = EuroCode) _
    As Double
    
    Dim Rates()     As Variant
    
    Dim RateTo      As Double
    Dim RateFrom    As Double
    Dim Factor      As Double
    Dim Index       As Integer
    
    If IsoFrom = "" Then
        IsoFrom = EuroCode
    End If
    If IsoTo = "" Then
        IsoTo = EuroCode
    End If
    
    If IsoTo = NeutralCode Or IsoFrom = NeutralCode Then
        Factor = NeutralRate
    ElseIf IsoTo = IsoFrom Then
        Factor = NeutralRate
    Else
        Rates() = ExchangeRatesEcb
    
        If IsoFrom = EuroCode Then
            RateFrom = NeutralRate
        Else
            For Index = LBound(Rates) To UBound(Rates)
                If Rates(Index, RateDetail.Code) = IsoFrom Then
                    RateFrom = Rates(Index, RateDetail.Rate)
                    Exit For
                End If
            Next
        End If
        
        If RateFrom > NoRate Then
            If IsoTo = EuroCode Then
                RateTo = NeutralRate
            Else
                For Index = LBound(Rates) To UBound(Rates)
                    If Rates(Index, RateDetail.Code) = IsoTo Then
                        RateTo = Rates(Index, RateDetail.Rate)
                        Exit For
                    End If
                Next
            End If
            Factor = RateTo / RateFrom
        End If
    End If
    
    CurrencyConvertEcb = Factor

End Function

' Returns the current conversion factor from any currency to another currency
' based on the exchange rates published by "ExchangeRate API".
'
' Optionally, the conversion factor can be calculated from any other of the
' published exchange rates.
'
' If an invalid or unpublished currency code is passed, a conversion factor
' of zero is returned.
'
' Examples, typical:
'   CurrencyConvertEra("DKK")           ->  7.46073161
'   CurrencyConvertEra("DKK", "EUR")    ->  7.46073161
'   CurrencyConvertEra("AUD")           ->  1.61848928
'   CurrencyConvertEra("AUD", "DKK")    ->  0.21695816
'   CurrencyConvertEra("DKK", "AUD")    ->  4.60920808
'   CurrencyConvertEra("EUR", "DKK")    ->  0.13403512
'   CurrencyConvertEra("", "DKK")       ->  0.13403512
'   CurrencyConvertEra("EUR")           ->  1
' Examples, neutral code.
'   CurrencyConvertEra("AUD", "XXX")    ->  1
'   CurrencyConvertEra("XXX", "AUD")    ->  1
'   CurrencyConvertEra("XXX")           ->  1
' Examples, invalid code.
'   CurrencyConvertEra("XYZ")           ->  0
'   CurrencyConvertEra("DKK", "XYZ")    ->  0
'
' 2018-09-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function CurrencyConvertEra( _
    ByVal IsoTo As String, _
    Optional ByVal IsoFrom As String = EuroCode) _
    As Double
    
    Dim Rates()     As Variant
    
    Dim IsoBase     As String
    Dim RateTo      As Double
    Dim RateFrom    As Double
    Dim Factor      As Double
    Dim Index       As Integer
    
    If IsoFrom = "" Then
        IsoFrom = EuroCode
    End If
    If IsoTo = "" Then
        IsoTo = EuroCode
    End If
    
    If IsoTo = NeutralCode Or IsoFrom = NeutralCode Then
        Factor = NeutralRate
    ElseIf IsoTo = IsoFrom Then
        Factor = NeutralRate
    Else
        ' Retrieve current rates using IsoFrom as the base currency.
        IsoBase = IsoFrom
        Rates() = ExchangeRatesEra(IsoBase)
        
        ' Look up the Factor of IsoBase.
        For Index = LBound(Rates) To UBound(Rates)
            If Rates(Index, RateDetail.Code) = IsoFrom Then
                RateFrom = Rates(Index, RateDetail.Rate)
                Exit For
            End If
        Next
        
        If RateFrom > NoRate Then
            ' Look up the Factor of IsoTo.
            For Index = LBound(Rates) To UBound(Rates)
                If Rates(Index, RateDetail.Code) = IsoTo Then
                    RateTo = Rates(Index, RateDetail.Rate)
                    Exit For
                End If
            Next
            Factor = RateTo / RateFrom
        End If
    End If
    
    CurrencyConvertEra = Factor

End Function

' Returns the current conversion factor from Euro to another currency
' based on the exchange rates published by "Fixer".
'
' Optionally, the conversion factor can be calculated from any other of the
' published exchange rates.
' For the free plan, exchange rates for other base currencies are
' calculated from EUR by triangular calculation.
'
' If an invalid or unpublished currency code is passed, a conversion factor
' of zero is returned.
'
' Examples, typical:
'   CurrencyConvertFxr("DKK")           ->  7.459225
'   CurrencyConvertFxr("DKK", "EUR")    ->  7.459225
'   CurrencyConvertFxr("AUD")           ->  1.618129
'   CurrencyConvertFxr("AUD", "DKK")    ->  0.216929908938261
'   CurrencyConvertFxr("DKK", "AUD")    ->  4.60978389238435
'   CurrencyConvertFxr("EUR", "DKK")    ->  0.134062184744394
'   CurrencyConvertFxr("", "DKK")       ->  0.134062184744394
'   CurrencyConvertFxr("EUR")           ->  1
' Examples, neutral code.
'   CurrencyConvertFxr("AUD", "XXX")    ->  1
'   CurrencyConvertFxr("XXX", "AUD")    ->  1
'   CurrencyConvertFxr("XXX")           ->  1
' Examples, invalid code.
'   CurrencyConvertFxr("XYZ")           ->  0
'   CurrencyConvertFxr("DKK", "XYZ")    ->  0
'
' 2018-09-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function CurrencyConvertFxr( _
    ByVal IsoTo As String, _
    Optional ByVal IsoFrom As String = EuroCode) _
    As Double
    
    Dim Rates()     As Variant
    
    Dim IsoBase     As String
    Dim RateTo      As Double
    Dim RateFrom    As Double
    Dim Factor      As Double
    Dim Index       As Integer
    
    If IsoFrom = "" Then
        IsoFrom = EuroCode
    End If
    If IsoTo = "" Then
        IsoTo = EuroCode
    End If
    
    If IsoTo = NeutralCode Or IsoFrom = NeutralCode Then
        Factor = NeutralRate
    ElseIf IsoTo = IsoFrom Then
        Factor = NeutralRate
    Else
        ' Retrieve current rates using IsoFrom as the base currency.
        ' If using the free license, EUR is always the base currency, thus
        ' triangular calculation of the Factor will be used for other
        ' base currencies than EUR.
        IsoBase = IsoFrom
        Rates() = ExchangeRatesFxr(IsoBase)
        
        ' Look up the Factor of IsoBase.
        For Index = LBound(Rates) To UBound(Rates)
            If Rates(Index, RateDetail.Code) = IsoFrom Then
                RateFrom = Rates(Index, RateDetail.Rate)
                Exit For
            End If
        Next
        
        If RateFrom > NoRate Then
            ' Look up the Factor of IsoTo.
            For Index = LBound(Rates) To UBound(Rates)
                If Rates(Index, RateDetail.Code) = IsoTo Then
                    RateTo = Rates(Index, RateDetail.Rate)
                    Exit For
                End If
            Next
            Factor = RateTo / RateFrom
        End If
    End If
    
    CurrencyConvertFxr = Factor

End Function

' Returns the current conversion factor from US Dollar to another currency
' based on the exchange rates published by "Open Exchange Rates".
'
' Optionally, the conversion factor can be calculated from any other of the
' published exchange rates.
'
' If an invalid or unpublished currency code is passed, a conversion factor
' of zero is returned.
'
' Examples, typical:
'   CurrencyConvertOxr("DKK")           ->  6.324293
'   CurrencyConvertOxr("DKK", "EUR")    ->  7.45837642771642
'   CurrencyConvertOxr("AUD")           ->  1.370338
'   CurrencyConvertOxr("AUD", "DKK")    ->  0.216678449274883
'   CurrencyConvertOxr("DKK", "AUD")    ->  4.61513363856217
'   CurrencyConvertOxr("EUR", "DKK")    ->  0.134077437588676
'   CurrencyConvertOxr("", "DKK")       ->  0.158120441288852
'   CurrencyConvertOxr("USD")           ->  1
' Examples, neutral code.
'   CurrencyConvertOxr("AUD", "XXX")    ->  1
'   CurrencyConvertOxr("XXX", "AUD")    ->  1
'   CurrencyConvertOxr("XXX")           ->  1
' Examples, invalid code.
'   CurrencyConvertOxr("XYZ")           ->  0
'   CurrencyConvertOxr("DKK", "XYZ")    ->  0
'
' 2018-10-12. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function CurrencyConvertOxr( _
    ByVal IsoTo As String, _
    Optional ByVal IsoFrom As String = USDollarCode) _
    As Double
    
    Dim Rates()     As Variant
    
    Dim IsoBase     As String
    Dim RateTo      As Double
    Dim RateFrom    As Double
    Dim Factor      As Double
    Dim Index       As Integer
    
    If IsoFrom = "" Then
        IsoFrom = USDollarCode
    End If
    If IsoTo = "" Then
        IsoTo = USDollarCode
    End If
    
    If IsoTo = NeutralCode Or IsoFrom = NeutralCode Then
        Factor = NeutralRate
    ElseIf IsoTo = IsoFrom Then
        Factor = NeutralRate
    Else
        ' Retrieve current rates using IsoFrom as the base currency.
        ' If using the free plan, USD is always the base currency.
        ' Thus, when using the free plan with a base currency other
        ' than USD, triangular calculation of the rate will be used.
        IsoBase = IsoFrom
        Rates() = ExchangeRatesOxr(IsoBase)
        
        ' Look up the rate of IsoFrom.
        For Index = LBound(Rates) To UBound(Rates)
            If Rates(Index, RateDetail.Code) = IsoFrom Then
                RateFrom = Rates(Index, RateDetail.Rate)
                Exit For
            End If
        Next
        
        If RateFrom > NoRate Then
            ' Look up the rate of Isoto.
            For Index = LBound(Rates) To UBound(Rates)
                If Rates(Index, RateDetail.Code) = IsoTo Then
                    RateTo = Rates(Index, RateDetail.Rate)
                    Exit For
                End If
            Next
            Factor = RateTo / RateFrom
        End If
    End If
    
    CurrencyConvertOxr = Factor

End Function

' Returns the current conversion factor from US Dollar to another currency
' based on the exchange rates published by "XE".
'
' Optionally, the conversion factor can be calculated from any other of the
' published exchange rates.
'
' If an invalid or unpublished currency code is passed, a conversion factor
' of zero is returned.
'
' Examples, typical:
'   CurrencyConvertXec("DKK")           ->  6.453107743
'   CurrencyConvertXec("DKK", "EUR")    ->  7.4699364684
'   CurrencyConvertXec("AUD")           ->  1.406057001
'   CurrencyConvertXec("AUD", "DKK")    ->  0.2178883504
'   CurrencyConvertXec("DKK", "AUD")    ->  4.5895064983
'   CurrencyConvertXec("EUR", "DKK")    ->  0.1338699471
'   CurrencyConvertXec("", "DKK")       ->  0.1549640948
'   CurrencyConvertXec("USD")           ->  1
' Examples, neutral code.
'   CurrencyConvertXec("AUD", "XXX")    ->  1
'   CurrencyConvertXec("XXX", "AUD")    ->  1
'   CurrencyConvertXec("XXX")           ->  1
' Examples, invalid code.
'   CurrencyConvertXec("XYZ")           ->  0
'   CurrencyConvertXec("DKK", "XYZ")    ->  0
'
' 2018-09-20. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function CurrencyConvertXec( _
    ByVal IsoTo As String, _
    Optional ByVal IsoFrom As String = USDollarCode) _
    As Double
    
    Dim Rates()     As Variant
    
    Dim IsoBase     As String
    Dim RateTo      As Double
    Dim RateFrom    As Double
    Dim Factor      As Double
    Dim Index       As Integer
    
    If IsoFrom = "" Then
        IsoFrom = USDollarCode
    End If
    If IsoTo = "" Then
        IsoTo = USDollarCode
    End If
    
    If IsoTo = NeutralCode Or IsoFrom = NeutralCode Then
        Factor = NeutralRate
    ElseIf IsoTo = IsoFrom Then
        Factor = NeutralRate
    Else
        ' Retrieve current rates using IsoFrom as the base currency.
        ' If using the free plan, USD is always the base currency.
        ' Thus, when using the free plan with a base currency other
        ' than USD, triangular calculation of the rate will be used.
        IsoBase = IsoFrom
        Rates() = ExchangeRatesXec(IsoBase)
        
        ' Look up the rate of IsoFrom.
        For Index = LBound(Rates) To UBound(Rates)
            If Rates(Index, RateDetail.Code) = IsoFrom Then
                RateFrom = Rates(Index, RateDetail.Rate)
                Exit For
            End If
        Next
        
        If RateFrom > NoRate Then
            ' Look up the rate of Isoto.
            For Index = LBound(Rates) To UBound(Rates)
                If Rates(Index, RateDetail.Code) = IsoTo Then
                    RateTo = Rates(Index, RateDetail.Rate)
                    Exit For
                End If
            Next
            Factor = RateTo / RateFrom
        End If
    End If
    
    CurrencyConvertXec = Factor

End Function

' Retrieve the current exchange rates from the Central Bank of the Russian
' Federation having RUB as the base currency.
' The rates are returned as an array and cached until the next update.
' The rates are updated once a day at about UTC 13:00.
'
' Source:
'   https://cbr.ru/eng/currency_base/daily/
'
' Note:
'   The Central Bank of the Russian Federation has set the exchange rates of
'   foreign currencies against the ruble without assuming any liability to
'   buy or sell foreign currency at the rates.
'
' Example:
'   Dim Rates As Variant
'   Rates = ExchangeRatesCbr()
'   Rates(9, 0) -> 2018-10-06       ' Publishing date.
'   Rates(9, 1) -> "DKK"            ' Currency code.
'   Rates(9, 2) -> 10.2697          ' Exchange rate.
'   Rates(9, 3) -> "Danish Krone"   ' Currency name in English.
'
' 2018-10-07. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ExchangeRatesCbr( _
    Optional ByVal LanguageCode As String) _
    As Variant

    ' Operational constants.
    '
    ' API endpoints.
    Const RuServiceUrl  As String = "https://cbr.ru/currency_base/daily/"
    Const EnServiceUrl  As String = "https://cbr.ru/eng/currency_base/daily/"
    
    ' Functional constants.
    '
    ' Page encoding.
    Const Characterset  As String = "UTF-8"
    ' Async setting.
    Const Async         As Variant = False
    ' Class name of data table.
    Const DataClassName As String = "data"
    ' Field items of html table.
    Const CodeField     As Integer = 1
    Const NameField     As Integer = 3
    Const UnitField     As Integer = 2
    Const RateField     As Integer = 4
    ' Locater/header for publishing date: "DT":".
    Const DateHeader    As String = """DT"":"""
    ' Length of formatted date: 2000-01-01.
    Const DateLength    As Integer = 10
    
    ' Update hour (UTC).
    Const UpdateHour    As Date = #1:00:00 PM#
    ' Update interval: 24 hours.
    Const UpdatePause   As Integer = 24
    ' English language code.
    Const EnglishCode   As String = "en"
    ' Russion language code.
    Const RussianCode   As String = "ru"
    

#If EarlyBinding Then
    ' Microsoft XML, v6.0.
    Dim XmlHttp         As MSXML2.ServerXMLHTTP60
    ' Microsoft ActiveX Data Objects 6.1 Library.
    Dim Stream          As ADODB.Stream
    ' Microsoft HTML Object Library.
    Dim Document        As MSHTML.HTMLDocument
    Dim Scripts         As MSHTML.IHTMLElementCollection
    Dim Script          As MSHTML.HTMLHtmlElement
    Dim Tables          As MSHTML.IHTMLElementCollection
    Dim Table           As MSHTML.HTMLHtmlElement
    Dim Rows            As MSHTML.IHTMLElementCollection
    Dim Row             As MSHTML.HTMLHtmlElement
    Dim Fields          As MSHTML.IHTMLElementCollection

    Set XmlHttp = New MSXML2.ServerXMLHTTP60
    Set Stream = New ADODB.Stream
    Set Document = New MSHTML.HTMLDocument
#Else
    Dim XmlHttp         As Object
    Dim Stream          As Object
    Dim Document        As Object
    Dim Scripts         As Object
    Dim Script          As Object
    Dim Tables          As Object
    Dim Table           As Object
    Dim Rows            As Object
    Dim Row             As Object
    Dim Fields          As Object
    
    Set XmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    Set Stream = CreateObject("ADODB.Stream")
    Set Document = CreateObject("htmlfile")
#End If

    Static Rates()      As Variant
    Static LastCall     As Date
    Static LastCode     As String
    
    Dim ServiceUrl      As String
    Dim RateCount       As Integer
    Dim Published       As String
    Dim ValueDate       As Date
    Dim ThisCall        As Date
    Dim Text            As String
    Dim Index           As Integer
    Dim Unit            As Double
    Dim ScaledRate      As Double
    Dim TrueRate        As Double
    
    If StrComp(LanguageCode, RussianCode, vbTextCompare) = 0 Then
        LanguageCode = RussianCode
        ServiceUrl = RuServiceUrl
    Else
        LanguageCode = EnglishCode
        ServiceUrl = EnServiceUrl
    End If
    
    If LastCode = LanguageCode And DateDiff("h", LastCall, UtcNow) < UpdatePause Then
        ' Return cached rates.
    Else
        ' Retrieve updated rates.
    
        ' Define default result array.
        ' Redim for four dimensions: date, code, rate, name.
        ReDim Rates(0, 0 To 3)
        Rates(0, RateDetail.Date) = NoValueDate
        Rates(0, RateDetail.Code) = NeutralCode
        Rates(0, RateDetail.Rate) = NeutralRate
        Rates(0, RateDetail.Name) = NeutralName
        
        ' Retrieve data.
        XmlHttp.Open "GET", ServiceUrl, Async
        XmlHttp.Send
        If XmlHttp.Status = HttpStatus.OK Then
            ' Retrieve and convert the page.
            ' The default character set cannot be used. See:
            ' https://stackoverflow.com/a/23812869/3527297
            
            ' Write the raw bytes to the stream.
            Stream.Open
            Stream.Type = adTypeBinary
            Stream.Write XmlHttp.responseBody
            ' Read text characters from the stream applying the character set.
            Stream.Position = 0
            Stream.Type = adTypeText
            Stream.Charset = Characterset
            ' Copy the page to the document object.
            Document.body.innerHTML = Stream.ReadText
        
            ' Search the scripts to locate the publishing date.
            Set Scripts = Document.getElementsByTagName("script")
            ValueDate = Date
            For Each Script In Scripts
                Text = Script.innerHTML
                If InStr(Text, "uniDbQuery_Data =") > 0 Then
                    Published = Left(Split(Text, DateHeader)(1), DateLength)
                    If IsDate(Published) Then
                        ValueDate = CDate(Published)
                    End If
                    Exit For
                End If
            Next
        
            ' Search the tables to locate the data table.
            ' Doesn't work with late binding.
            ' Set Tables = Document.getElementsByClassName("data")
            Set Tables = Document.getElementsByTagName("table")
            For Each Table In Tables
                If Table.className = DataClassName Then
                    Exit For
                End If
            Next
            
            If Not Table Is Nothing Then
                ' The table was found.
                Set Rows = Table.getElementsByTagName("tr")
                ' Reduce the count by one to skip the header row.
                RateCount = Rows.Length - 1
                ' Redim for four dimensions: date, code, rate, name.
                ReDim Rates(0 To RateCount - 1, 0 To 3)
                
                ' Fill the array of rates.
                For Index = LBound(Rates, 1) To UBound(Rates, 1)
                    ' Offset Index by one to skip the header row.
                    Set Row = Rows.Item(Index + 1)
                    ' Get the fields of this rate.
                    Set Fields = Row.getElementsByTagName("td")
                    
                    ' The returned rates are scaled to hold four decimals only.
                    ' Calculate the true (non-scaled) rate.
                    ScaledRate = Val(Replace(Fields.Item(RateField).innerText, ",", "."))
                    Unit = Val(Fields.Item(UnitField).innerText)
                    TrueRate = ScaledRate / Unit
                    
                    Rates(Index, RateDetail.Date) = ValueDate
                    Rates(Index, RateDetail.Code) = Fields.Item(CodeField).innerText
                    Rates(Index, RateDetail.Rate) = TrueRate
                    Rates(Index, RateDetail.Name) = Fields.Item(NameField).innerHTML
                Next
            End If
            
            ThisCall = ValueDate + UpdateHour
            ' Record requested language and publishing time of retrieved rates.
            LastCode = LanguageCode
            LastCall = ThisCall
            
        End If
    End If
    
    ExchangeRatesCbr = Rates

End Function

' Retrieve the current exchange rate from "Currency Converter API" for one base currency.
' The requested rate is returned as an array and cached until the next update.
' All retrieved rates are cached in a collection until the next update.
' The rates are updated from once per hour down to once per minute.
'
' Default base currency is EUR.
' Default rate is for USD.
'
' Source:
'   https://currencyconverterapi.com/
'   https://currencyconverterapi.com/docs
'
' Note:
'   The services are provided as is and without warranty.
'
' Example:
'   Dim Rates As Variant
'   Rates = ExchangeRatesCca()
'   Rates(0, 0) -> 2018-09-24 07:56:50  ' Publishing date.
'   Rates(0, 1) -> "USD"                ' Currency code.
'   Rates(0, 2) -> 1.17395              ' Exchange rate.
'
' 2018-09-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ExchangeRatesCca( _
    Optional ByVal IsoBase As String = EuroCode, _
    Optional ByVal IsoTo As String = USDollarCode) _
    As Variant
    
    ' Operational constants.
    '
    ' API endpoint.
    Const FreeSubdomain As String = "free"
    Const PaidSubdomain As String = "api"
    Const TempSubdomain As String = "xxx"
    ' API version must be 3 or higher.
    Const ApiVersion    As String = "6"
    Const ServiceUrl    As String = "https://" & TempSubdomain & ".currencyconverterapi.com/api/v" & ApiVersion & "/convert"
    ' Data styles. For reference only; must be "ultra".
    Const CompactStyle  As String = "ultra"
    Const ExtendedStyle As String = ""
    ' Update interval: 60, 15, or 1 minutes.
    Const UpdatePause   As Integer = 60
    
    ' Function constants.
    '
    ' Default currency code. Can be any valid currency codes.
    Const DefaultBase   As String = EuroCode
    Const DefaultTo     As String = USDollarCode
    ' Node names in retrieved collection.
    Const RootNodeName  As String = "root"
    ' ResponseText when invalid currency code is passed.
    Const EmptyResponse As String = "{}"
    
    Static CodePairs    As Collection
    
    Static Rates()      As Variant
    Static LastCodePair As String
    Static LastCall     As Date
    
    Dim DataCollection  As Collection
    
    Dim Parameter()     As String
    Dim Parameters()    As String
    Dim UrlParts(1)     As String
    
    Dim Subdomain       As String
    Dim CodePair        As String
    Dim RateItem        As Variant
    Dim Index           As Integer
    Dim Url             As String
    Dim ResponseText    As String
    Dim ValueDate       As Date
    Dim ThisCall        As Date
    Dim IsCurrent       As Boolean
    
    ' Assemple code pair.
    If IsoBase = "" Then
        IsoBase = DefaultBase
    End If
    If IsoTo = "" Then
        IsoTo = DefaultTo
    End If
    CodePair = Trim(Left(UCase(IsoBase), 3)) & "_" & Trim(Left(UCase(IsoTo), 3))
    
    ' Is the current collection of rates up-to-date?
    IsCurrent = DateDiff("n", LastCall, Now) < UpdatePause
    
    If IsCurrent And LastCodePair = CodePair Then
        ' Return cached rate.
    ElseIf IsCurrent And IsCollectionItem(CodePairs, CodePair) Then
        ' Return stored rate from collection.
        Rates = CodePairs(CodePair)
        LastCodePair = CodePair
    Else
        ' Retrieve the code pair and add it to the collection of code pairs.
        If IsCurrent Then
            ' Keep the stored code pairs.
        Else
            ' Clear all stored code pairs.
            Set CodePairs = New Collection
        End If
        
        ' Set subdomain to call.
        If CcaApiId = "" Then
            ' Free plan is used.
            Subdomain = FreeSubdomain
        Else
            ' Paid plan is used.
            Subdomain = PaidSubdomain
        End If
        
        ' Define parameter array.
        ' Redim for two dimensions: name, value.
        ReDim Parameter(0 To 2, 0 To 1)
        ' Parameter names.
        Parameter(0, ParameterDetail.Name) = "q"
        Parameter(1, ParameterDetail.Name) = "compact"
        Parameter(2, ParameterDetail.Name) = "apiKey"
        ' Parameter values.
        Parameter(0, ParameterDetail.Value) = CodePair
        Parameter(1, ParameterDetail.Value) = CompactStyle
        Parameter(2, ParameterDetail.Value) = CcaApiId
        
        ' Assemble parameters.
        ReDim Parameters(LBound(Parameter, 1) To UBound(Parameter, 1))
        For Index = LBound(Parameters) To UBound(Parameters)
            Parameters(Index) = Parameter(Index, 0) & "=" & Parameter(Index, 1)
        Next
        
        ' Assemble URL.
        UrlParts(0) = Replace(ServiceUrl, TempSubdomain, Subdomain)
        UrlParts(1) = Join(Parameters, "&")
        Url = Join(UrlParts, "?")
        ' Uncomment for debugging.
        'Debug.Print Url
        
        ' Define default result array.
        ' Redim for three dimensions: date, code, rate.
        ReDim Rates(0, 0 To 2)
        Rates(0, RateDetail.Date) = NoValueDate
        Rates(0, RateDetail.Code) = NeutralCode
        Rates(0, RateDetail.Rate) = NeutralRate
        
        If RetrieveDataResponse(Url, ResponseText) = True Then
            Set DataCollection = CollectJson(ResponseText)
        End If
    
        If DataCollection Is Nothing Then
            ' Error. ResponseText holds the error code.
            ' Optional error handling.
            Select Case ResponseText
                Case HttpStatus.BadRequest
                    ' Typical for invalid api key, or API limit reached.
                Case EmptyResponse
                    ' Invalid currency code.
                Case Else
                    ' Other error.
            End Select
            ' Set "not found" return values.
            Rates(0, RateDetail.Code) = NoCode
            Rates(0, RateDetail.Rate) = NoRate
        End If
        
        If Not DataCollection Is Nothing Then
            ' The rate was retrieved.
            ' Get the UTC value date and time for the rate.
            ValueDate = UtcNow
            
            ' The retrieved rate item is an array.
            RateItem = DataCollection(RootNodeName)(CollectionItem.Data)(1)
            Rates(0, RateDetail.Date) = ValueDate
            Rates(0, RateDetail.Code) = Split(RateItem(CollectionItem.Name), "_")(1)
            Rates(0, RateDetail.Rate) = RateItem(CollectionItem.Data)
            
            ' Store this code pair in the collection of code pairs.
            CodePairs.Add Rates, CodePair
            
            Set DataCollection = Nothing
            
            ' Round the call time down to the start of the update interval.
            ThisCall = CDate(Fix(Now * 24 * 60 / UpdatePause) / (24 * 60 / UpdatePause))
            ' Record hour of retrieval.
            LastCall = ThisCall
        End If
        ' Record requested base currency.
        LastCodePair = CodePair
    End If
    
    ExchangeRatesCca = Rates

End Function

' Retrieve the current exchange rates from "Currencylayer API" for one base currency.
' The rates are returned as an array and cached until the next update.
' The rates are updated from once per hour down to once per minute.
'
' Default base currency is USD.
' For the free plan, exchange rates for other base currencies are
' calculated from USD by triangular calculation.
'
' Source:
'   https://currencylayer.com/
'   https://currencylayer.com/documentation
'
' Note:
'   Exchange rates are classed as indicative rates and are accurate enough to display price estimations.
'   The rates are unsuitable for forex trading or processing cross currency settlements.
'
' Example:
'   Dim Rates As Variant
'   Rates = ExchangeRatesCla()
'   Rates(12, 0) -> 2018-09-20 08:54:06 ' Publishing date.
'   Rates(12, 1) -> "BDT"               ' Currency code.
'   Rates(12, 2) -> 84.064038           ' Exchange rate.
'
' 2018-10-07. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ExchangeRatesCla( _
    Optional ByVal IsoBase As String) _
    As Variant
    
    ' Operational constants.
    '
    ' API endpoint for the free plan.
    ' For the paid plans, https may be used.
    Const ServiceUrl    As String = "http://www.apilayer.net/api/live"
    ' Update interval: 60, 10, or 1 minutes.
    Const UpdatePause   As Integer = 60
    
    ' Function constants.
    '
    ' Default base currency code.
    Const DefaultBase   As String = USDollarCode
    ' Node names in retrieved collection.
    Const RootNodeName  As String = "root"
    Const TimeNodeName  As String = "timestamp"
    Const RateNodeName  As String = "quotes"
    Const FirstNodeName As String = "success"
    Const ErrorNodeName As String = "error"
    Const CodeNodeName  As String = "code"
    ' Error code for invalid or missing access key.
    Const KeyErrorCode  As Long = 101
    ' Error code for restricted access to base currency.
    Const BaseErrorCode As Long = 105
    ' Error code for invalid currency code.
    Const CodeErrorCode As Long = 201
    
    Static Rates()      As Variant
    Static LastCode     As String
    Static LastCall     As Date
    
    Dim DataCollection  As Collection
    
    Dim Parameters()    As String
    Dim Parameter()     As String
    Dim UrlParts(1)     As String
    
    Dim RateCount       As Integer
    Dim RateItem        As Variant
    Dim BaseRate        As Double
    Dim Index           As Integer
    Dim Url             As String
    Dim ResponseText    As String
    Dim ValueDate       As Date
    Dim ThisCall        As Date
    Dim ErrorCode       As Long
    
    If IsoBase = "" Then
        IsoBase = DefaultBase
    End If
    
    If LastCode = IsoBase And DateDiff("n", LastCall, Now) < UpdatePause Then
        ' Return cached rates.
    Else
        ' Retrieve updated rates.
        
        ' Define parameter array.
        ' Redim for two dimensions: name, value.
        ReDim Parameter(0 To 1, 0 To 1)
        ' Parameter names.
        Parameter(0, ParameterDetail.Name) = "access_key"
        Parameter(1, ParameterDetail.Name) = "source"
        ' Parameter values.
        Parameter(0, ParameterDetail.Value) = ApiApiId
        Parameter(1, ParameterDetail.Value) = IsoBase
        
        ' Assemble parameters.
        ReDim Parameters(LBound(Parameter, 1) To UBound(Parameter, 1))
        For Index = LBound(Parameters) To UBound(Parameters)
            Parameters(Index) = Parameter(Index, 0) & "=" & Parameter(Index, 1)
        Next
        
        ' Assemble URL.
        UrlParts(0) = ServiceUrl
        UrlParts(1) = Join(Parameters, "&")
        Url = Join(UrlParts, "?")
        ' Uncomment for debugging.
        ' Debug.Print Url
        
        ' Define default result array.
        ' Redim for three dimensions: date, code, rate.
        ReDim Rates(0, 0 To 2)
        Rates(0, RateDetail.Date) = NoValueDate
        Rates(0, RateDetail.Code) = NeutralCode
        Rates(0, RateDetail.Rate) = NeutralRate
        
        If RetrieveDataResponse(Url, ResponseText) = True Then
            Set DataCollection = CollectJson(ResponseText)
        Else
            ' Give up.
            Set DataCollection = Nothing
        End If
    
        If Not DataCollection Is Nothing Then
            If DataCollection(RootNodeName)(CollectionItem.Data)(1)(CollectionItem.Name) = FirstNodeName Then
                If DataCollection(RootNodeName)(CollectionItem.Data)(FirstNodeName)(CollectionItem.Data) = False Then
                    ErrorCode = DataCollection(RootNodeName)(CollectionItem.Data)(ErrorNodeName)(CollectionItem.Data)(CodeNodeName)(CollectionItem.Data)
                    Select Case ErrorCode
                        Case KeyErrorCode
                            ' Missing or invalid access key.
                            Set DataCollection = Nothing
                        Case CodeErrorCode, BaseErrorCode
                            ' Typical for invalid currency code, or if free license and base <> USD, respectively.
                            ' Rebuld Url to use base = USD.
                            Parameter(1, 1) = DefaultBase
                            ' Reassemble parameters.
                            For Index = LBound(Parameters) To UBound(Parameters)
                                Parameters(Index) = Parameter(Index, 0) & "=" & Parameter(Index, 1)
                            Next
                            
                            ' Reassemble URL.
                            UrlParts(0) = ServiceUrl
                            UrlParts(1) = Join(Parameters, "&")
                            Url = Join(UrlParts, "?")
                            
                            ' Try once more to retrieve the rates.
                            If RetrieveDataResponse(Url, ResponseText) = True Then
                                Set DataCollection = CollectJson(ResponseText)
                                If DataCollection(RootNodeName)(CollectionItem.Data)(FirstNodeName)(CollectionItem.Data) = False Then
                                    ' Give up.
                                    Set DataCollection = Nothing
                                End If
                            End If
                            ' Rebuld Url to use base = USD.
                            Parameter(1, 1) = DefaultBase
                            ' Reassemble parameters.
                            For Index = LBound(Parameters) To UBound(Parameters)
                                Parameters(Index) = Parameter(Index, 0) & "=" & Parameter(Index, 1)
                            Next
                            
                            ' Reassemble URL.
                            UrlParts(0) = ServiceUrl
                            UrlParts(1) = Join(Parameters, "&")
                            Url = Join(UrlParts, "?")
                            
                            ' Try once more to retrieve the rates.
                            If RetrieveDataResponse(Url, ResponseText) = True Then
                                Set DataCollection = CollectJson(ResponseText)
                                If DataCollection(RootNodeName)(CollectionItem.Data)(FirstNodeName)(CollectionItem.Data) = False Then
                                    ' Give up.
                                    Set DataCollection = Nothing
                                End If
                            End If
                    End Select
                End If
            End If
        End If
        
        If Not DataCollection Is Nothing Then
            ' Rates were retrieved.
            ' Get the UTC value date and time for the rates.
            ValueDate = DateUnix(DataCollection(RootNodeName)(CollectionItem.Data)(TimeNodeName)(CollectionItem.Data))
            ' Get count of rates.
            RateCount = DataCollection(RootNodeName)(CollectionItem.Data)(RateNodeName)(CollectionItem.Data).Count
            ' Redim for three dimensions: date, code, rate.
            ReDim Rates(RateCount - 1, 0 To 2)
            BaseRate = NeutralRate
    
            ' Fill the array from the collection items.
            For Index = 1 To RateCount
                ' A retrieved rate item is an array.
                RateItem = DataCollection(RootNodeName)(CollectionItem.Data)(RateNodeName)(CollectionItem.Data)(Index)
                Rates(Index - 1, RateDetail.Date) = ValueDate
                Rates(Index - 1, RateDetail.Code) = Right(RateItem(CollectionItem.Name), 3)
                Rates(Index - 1, RateDetail.Rate) = RateItem(CollectionItem.Data)
                If Right(RateItem(CollectionItem.Name), 3) = IsoBase And RateItem(CollectionItem.Data) <> NeutralRate Then
                    ' Prepare triangular calculation.
                    BaseRate = RateItem(CollectionItem.Data)
                End If
            Next
            If BaseRate <> NeutralRate Then
                For Index = 1 To RateCount
                    ' Perform triangular calculation of the exchange rates.
                    If Rates(Index - 1, RateDetail.Code) = IsoBase Then
                        Rates(Index - 1, RateDetail.Rate) = NeutralRate
                    Else
                        Rates(Index - 1, RateDetail.Rate) = Rates(Index - 1, RateDetail.Rate) / BaseRate
                    End If
                Next
            End If
            
            Set DataCollection = Nothing
            
            ' Round the call time down to the start of the update interval.
            ThisCall = CDate(Fix(Now * 24 * 60 / UpdatePause) / (24 * 60 / UpdatePause))
            ' Record requested base currency and hour of retrieval.
            LastCode = IsoBase
            LastCall = ThisCall
        End If
    End If
    
    ExchangeRatesCla = Rates

End Function

' Retrieve the current exchange rates from the National Bank of Denmark
' having DKK as the base currency.
' The rates are returned as an array and cached until the next update.
' The rates are updated once a day at about UTC 15:00.
'
' Source:
'   http://www.nationalbanken.dk/en/statistics/exchange_rates/Pages/default.aspx
'
' Note:
'   The exchange rates on Danmarks Nationalbank's website are indicative rates
'   that are not intended to be used in any market transaction.
'   The rates are intended for information purposes only.
'
' Defaults to English currency names.
' Optionally, setting parameter LanguageCode to "da", Danish names are retrieved.
'
' Example:
'   Dim Rates As Variant
'   Rates = ExchangeRatesDkk()
'   Rates(7, 0) -> 2018-05-30   ' Publishing date.
'   Rates(7, 1) -> "EUR"        ' Currency code.
'   Rates(7, 2) -> 7.4432       ' Exchange rate.
'   Rates(7, 3) -> "Euro"       ' Currency name, English or Danish.
'
' 2018-10-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ExchangeRatesDkk( _
    Optional ByVal LanguageCode As String) _
    As Variant

    ' Operational constants.
    '
    ' Base URL for Danmarks Nationalbank exchange rates.
    Const ServiceUrl    As String = "http://www.nationalbanken.dk/_vti_bin/DN/DataService.svc/CurrencyRatesXML"
    ' Update hour (UTC).
    Const UpdateHour    As Date = #3:00:00 PM#
    ' Update interval: 24 hours.
    Const UpdatePause   As Integer = 24
    ' Default language code.
    Const EnglishCode   As String = "en"
    ' Local language code.
    Const DanishCode    As String = "da"
    ' Base amount.
    Const BaseAmount    As Currency = 100
    
    ' Function constants.
    '
    ' Async setting.
    Const Async         As Variant = False
    ' XML node and attribute names.
    Const RootNodeName  As String = "exchangerates"
    Const TimeNodeName  As String = "dailyrates"
    Const TimeItemName  As String = "id"
    Const CodeItemName  As String = "code"
    Const RateItemName  As String = "rate"
    Const NameItemName  As String = "desc"
  
#If EarlyBinding Then
    ' Microsoft XML, v6.0.
    Dim Document        As MSXML2.DOMDocument60
    Dim XmlHttp         As MSXML2.ServerXMLHTTP60
    Dim RootNodeList    As MSXML2.IXMLDOMNodeList
    Dim TimeNodeList    As MSXML2.IXMLDOMNodeList
    Dim RateNodeList    As MSXML2.IXMLDOMNodeList
    Dim RootNode        As MSXML2.IXMLDOMNode
    Dim TimeNode        As MSXML2.IXMLDOMNode
    Dim RateNode        As MSXML2.IXMLDOMNode
    Dim RateAttribute   As MSXML2.IXMLDOMAttribute

    Set Document = New MSXML2.DOMDocument60
    Set XmlHttp = New MSXML2.ServerXMLHTTP60
#Else
    Dim Document        As Object
    Dim XmlHttp         As Object
    Dim RootNodeList    As Object
    Dim TimeNodeList    As Object
    Dim RateNodeList    As Object
    Dim RootNode        As Object
    Dim TimeNode        As Object
    Dim RateNode        As Object
    Dim RateAttribute   As Object

    Set Document = CreateObject("MSXML2.DOMDocument")
    Set XmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
#End If

    Static Rates()      As Variant
    Static LastCall     As Date
    Static LastCode     As String
    
    Dim Parameter()     As String
    Dim Parameters()    As String
    Dim UrlParts(1)     As String
    
    Dim Index           As Integer
    Dim Url             As String
    Dim CurrencyCode    As String
    Dim CurrencyName    As String
    Dim Rate            As String
    Dim ValueDate       As Date
    Dim ThisCall        As Date
    Dim Item            As Integer
    
    If StrComp(LanguageCode, DanishCode, vbTextCompare) = 0 Then
        LanguageCode = DanishCode
    Else
        LanguageCode = EnglishCode
    End If
    
    If LastCode = LanguageCode And DateDiff("h", LastCall, UtcNow) < UpdatePause Then
        ' Return cached rates.
    Else
        ' Retrieve updated rates.
    
        ' Define parameter array.
        ' Redim for two dimensions: name, value.
        ReDim Parameter(0 To 0, 0 To 1)
        ' Parameter names.
        Parameter(0, ParameterDetail.Name) = "lang"
        ' Parameter values.
        Parameter(0, ParameterDetail.Value) = LanguageCode
        
        ' Assemble parameters.
        ReDim Parameters(LBound(Parameter, 1) To UBound(Parameter, 1))
        For Index = LBound(Parameters) To UBound(Parameters)
            Parameters(Index) = Parameter(Index, 0) & "=" & Parameter(Index, 1)
        Next
        
        ' Assemble URL.
        UrlParts(0) = ServiceUrl
        UrlParts(1) = Join(Parameters, "&")
        Url = Join(UrlParts, "?")
        ' Uncomment for debugging.
        ' Debug.Print Url
        
        ' Define default result array.
        ' Redim for four dimensions: date, code, rate, name.
        ReDim Rates(0, 0 To 3)
        Rates(0, RateDetail.Date) = NoValueDate
        Rates(0, RateDetail.Code) = NeutralCode
        Rates(0, RateDetail.Rate) = NeutralRate
        Rates(0, RateDetail.Name) = NeutralName
        
        ' Retrieve data.
        XmlHttp.Open "GET", Url, Async
        XmlHttp.Send
        
        If XmlHttp.Status = HttpStatus.OK Then
            ' File retrieved successfully.
            Document.loadXML XmlHttp.ResponseText
        
            Set RootNodeList = Document.getElementsByTagName(RootNodeName)
            ' Find root node.
            For Each RootNode In RootNodeList
                If RootNode.nodeName = RootNodeName Then
                    Exit For
                Else
                    Set RootNode = Nothing
                End If
            Next
            
            If Not RootNode Is Nothing Then
                If RootNode.hasChildNodes Then
                    ' Find first level node.
                    Set TimeNodeList = RootNode.childNodes
                    For Each TimeNode In TimeNodeList
                        If TimeNode.nodeName = TimeNodeName Then
                            Exit For
                        Else
                            Set TimeNode = Nothing
                        End If
                    Next
                End If
            End If
            
            If Not TimeNode Is Nothing Then
                If TimeNode.hasChildNodes Then
                    ' Find value date.
                    ValueDate = CDate(TimeNode.Attributes.getNamedItem(TimeItemName).nodeValue)
                    
                    ' Find the exchange rates.
                    Set RateNodeList = TimeNode.childNodes
                    ' Redim for four dimensions: date, code, rate, name.
                    ReDim Rates(RateNodeList.Length - 1, 0 To 3)
                    For Each RateNode In RateNodeList
                        Rates(Item, RateDetail.Date) = ValueDate
                        If RateNode.Attributes.Length > 0 Then
                            ' Get the ISO currency code.
                            Set RateAttribute = RateNode.Attributes.getNamedItem(CodeItemName)
                            If Not RateAttribute Is Nothing Then
                                CurrencyCode = RateAttribute.nodeValue
                            End If
                            ' Get the exchange rate for this currency code.
                            Set RateAttribute = RateNode.Attributes.getNamedItem(RateItemName)
                            If Not RateAttribute Is Nothing Then
                                Rate = RateAttribute.nodeValue
                            End If
                            ' Get the currency name.
                            Set RateAttribute = RateNode.Attributes.getNamedItem(NameItemName)
                            If Not RateAttribute Is Nothing Then
                                CurrencyName = RateAttribute.nodeValue
                            End If
                            ' Fill this result item.
                            Rates(Item, RateDetail.Code) = CurrencyCode
                            ' Replace Danish decimal separator, comma, with dot.
                            Rates(Item, RateDetail.Rate) = CDbl(Val(Replace(Rate, ",", "."))) / BaseAmount
                            Rates(Item, RateDetail.Name) = CurrencyName
                        End If
                        Item = Item + 1
                    Next RateNode
                End If
            End If
            
            ThisCall = ValueDate + UpdateHour
            ' Record requested language and publishing time of retrieved rates.
            LastCode = LanguageCode
            LastCall = ThisCall

        End If
    End If
    
    ExchangeRatesDkk = Rates

End Function

' Retrieve the current exchange rates from the European Central Bank, ECB,
' for Euro having each of the listed currencies as the base currency.
' The rates are returned as an array and cached until the next update.
' The rates are updated once a day at about UTC 15:00.
'
' Source:
'   http://www.ecb.europa.eu/stats/policy_and_exchange_rates/euro_reference_exchange_rates/html/index.en.html
'
' Note:
'   The exchange rates on the European Central Bank's website are indicative rates
'   that are not intended to be used in any market transaction.
'   The rates are intended for information purposes only.
'
' Example:
'   Dim Rates As Variant
'   Rates = ExchangeRatesEcb()
'   Rates(7, 0) -> 2018-05-30       ' Publishing date.
'   Rates(7, 1) -> "PLN"            ' Currency code.
'   Rates(7, 2) -> 4.3135           ' Exchange rate.
'
' 2018-06-07. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ExchangeRatesEcb() As Variant

    ' Operational constants.
    '
    ' Base URL for European Central Bank exchange rates.
    Const ServiceUrl    As String = "http://www.ecb.europa.eu/stats/eurofxref/"
    ' File to look up.
    Const Filename      As String = "eurofxref-daily.xml"
    ' Update hour (UTC).
    Const UpdateHour    As Date = #3:00:00 PM#
    ' Update interval: 24 hours.
    Const UpdatePause   As Integer = 24
    
    ' Function constants.
    '
    ' Async setting.
    Const Async         As Variant = False
    ' XML node and attribute names.
    Const RootNodeName  As String = "gesmes:Envelope"
    Const CubeNodeName  As String = "Cube"
    Const TimeNodeName  As String = "Cube"
    Const TimeItemName  As String = "time"
    Const CodeItemName  As String = "currency"
    Const RateItemName  As String = "rate"
  
#If EarlyBinding Then
    ' Microsoft XML, v6.0.
    Dim Document        As MSXML2.DOMDocument60
    Dim XmlHttp         As MSXML2.ServerXMLHTTP60
    Dim RootNodeList    As MSXML2.IXMLDOMNodeList
    Dim CubeNodeList    As MSXML2.IXMLDOMNodeList
    Dim RateNodeList    As MSXML2.IXMLDOMNodeList
    Dim RootNode        As MSXML2.IXMLDOMNode
    Dim CubeNode        As MSXML2.IXMLDOMNode
    Dim TimeNode        As MSXML2.IXMLDOMNode
    Dim RateNode        As MSXML2.IXMLDOMNode
    Dim RateAttribute   As MSXML2.IXMLDOMAttribute

    Set Document = New MSXML2.DOMDocument60
    Set XmlHttp = New MSXML2.ServerXMLHTTP60
#Else
    Dim Document        As Object
    Dim XmlHttp         As Object
    Dim RootNodeList    As Object
    Dim CubeNodeList    As Object
    Dim RateNodeList    As Object
    Dim RootNode        As Object
    Dim CubeNode        As Object
    Dim TimeNode        As Object
    Dim RateNode        As Object
    Dim RateAttribute   As Object

    Set Document = CreateObject("MSXML2.DOMDocument")
    Set XmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
#End If

    Static Rates()      As Variant
    Static LastCall     As Date
    
    Dim Url             As String
    Dim CurrencyCode    As String
    Dim Rate            As String
    Dim ValueDate       As Date
    Dim ThisCall        As Date
    Dim Item            As Integer
    
    
    If DateDiff("h", LastCall, UtcNow) < UpdatePause Then
        ' Return cached rates.
    Else
        ' Retrieve updated rates.
    
        ' Define default result array.
        ' Redim for three dimensions: date, code, rate.
        ReDim Rates(0, 0 To 2)
        Rates(0, RateDetail.Date) = NoValueDate
        Rates(0, RateDetail.Code) = NeutralCode
        Rates(0, RateDetail.Rate) = NeutralRate
        
        Url = ServiceUrl & Filename
        
        ' Retrieve data.
        XmlHttp.Open "GET", Url, Async
        XmlHttp.Send
        
        If XmlHttp.Status = HttpStatus.OK Then
            ' File retrieved successfully.
            Document.loadXML XmlHttp.ResponseText
        
            Set RootNodeList = Document.getElementsByTagName(RootNodeName)
            ' Find root node.
            For Each RootNode In RootNodeList
                If RootNode.nodeName = RootNodeName Then
                    Exit For
                Else
                    Set RootNode = Nothing
                End If
            Next
            
            If Not RootNode Is Nothing Then
                If RootNode.hasChildNodes Then
                    ' Find first level Cube node.
                    Set CubeNodeList = RootNode.childNodes
                    For Each CubeNode In CubeNodeList
                        If CubeNode.nodeName = CubeNodeName Then
                            Exit For
                        Else
                            Set CubeNode = Nothing
                        End If
                    Next
                End If
            End If
            If Not CubeNode Is Nothing Then
                If CubeNode.hasChildNodes Then
                    ' Find second level Cube node.
                    Set CubeNodeList = CubeNode.childNodes
                    For Each TimeNode In CubeNodeList
                        If TimeNode.nodeName = TimeNodeName Then
                            Exit For
                        Else
                            Set TimeNode = Nothing
                        End If
                    Next
                End If
            End If
            
            If Not TimeNode Is Nothing Then
                If TimeNode.hasChildNodes Then
                    ' Find value date.
                    ValueDate = CDate(TimeNode.Attributes.getNamedItem(TimeItemName).nodeValue)
                    
                    ' Find the exchange rates.
                    Set RateNodeList = TimeNode.childNodes
                    ' Redim for three dimensions: date, code, rate.
                    ReDim Rates(RateNodeList.Length - 1, 0 To 2)
                    For Each RateNode In RateNodeList
                        Rates(Item, RateDetail.Date) = ValueDate
                        If RateNode.Attributes.Length > 0 Then
                            ' Get the ISO currency code.
                            Set RateAttribute = RateNode.Attributes.getNamedItem(CodeItemName)
                            If Not RateAttribute Is Nothing Then
                                CurrencyCode = RateAttribute.nodeValue
                            End If
                            ' Get the exchange rate for this currency code.
                            Set RateAttribute = RateNode.Attributes.getNamedItem(RateItemName)
                            If Not RateAttribute Is Nothing Then
                                Rate = RateAttribute.nodeValue
                            End If
                            Rates(Item, RateDetail.Code) = CurrencyCode
                            Rates(Item, RateDetail.Rate) = CDbl(Val(Rate))
                        End If
                        Item = Item + 1
                    Next RateNode
                End If
            End If
            
            ThisCall = ValueDate + UpdateHour
            ' Record requested language and publishing time of retrieved rates.
            LastCall = ThisCall
            
        End If
    End If
    
    ExchangeRatesEcb = Rates

End Function

' Retrieve the current exchange rates from "ExchangeRate API" for one base currency.
' The rates are returned as an array and cached until the next update.
' The rates are updated from once per hour down to once per minute.
'
' Default base currency is EUR.
'
' Source:
'   https://www.exchangerate-api.com/
'   https://www.exchangerate-api.com/documentation
'
' Note:
'   Exchange rates are classed as indicative rates and are accurate enough to display price estimations.
'   The rates are unsuitable for forex trading or processing cross currency settlements.
'
' Example:
'   Dim Rates As Variant
'   Rates = ExchangeRatesFxr()
'   Rates(12, 0) -> 2018-09-20 08:54:06 ' Publishing date.
'   Rates(12, 1) -> "BDT"               ' Currency code.
'   Rates(12, 2) -> 98.26592            ' Exchange rate.
'
' 2018-10-07. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ExchangeRatesEra( _
    Optional ByVal IsoBase As String, _
    Optional ByVal IsoTo As String) _
    As Variant
    
    ' Operational constants.
    '
    ' API endpoint.
    Const ServiceUrl    As String = "https://v3.exchangerate-api.com/"
    ' Update interval: 60, 15, or 5 minutes.
    Const UpdatePause   As Integer = 60
    
    ' Function constants.
    '
    ' Default base currency code. Can be any valid currency code.
    Const DefaultBase   As String = EuroCode
    ' URL paths.
    Const AllRatesPath  As String = "bulk"
    Const OneRatePath   As String = "pair"
    ' Node names in retrieved collection.
    Const RootNodeName  As String = "root"
    Const TimeNodeName  As String = "timestamp"
    Const FromNodeName  As String = "from"
    Const ToNodeName    As String = "to"
    Const RateNodeName  As String = "rates"
    Const Rate1NodeName As String = "rate"
    Const FirstNodeName As String = "result"
    Const ErrorNodeName As String = "error"
    Const UBoundBulk    As Long = 3
    Const UBoundPair    As Long = UBoundBulk + 1
    
    ' Result codes.
    '
    Const ResultError   As String = "error"
    Const ResultSuccess As String = "success"
    
    ' Error codes.
    '
    ' Error code for unknown currency code.
    Const UnknownCode   As String = "unknown-code"
    ' Error code for invalid api key.
    Const InvalidKey    As String = "invalid-key"
    
    Static Rates()      As Variant
    Static LastFromCode As String
    Static LastToCode   As String
    Static LastCall     As Date
    
    Dim DataCollection  As Collection
    
    Dim UrlPath()       As String
    
    Dim BulkRates       As Boolean
    Dim RateCount       As Integer
    Dim RateItem        As Variant
    Dim Index           As Integer
    Dim Url             As String
    Dim CodePath        As String
    Dim ResponseText    As String
    Dim ValueDate       As Date
    Dim ThisCall        As Date
    Dim ResultCode      As String
    Dim ErrorCode       As String
    
    If IsoBase = "" Then
        IsoBase = DefaultBase
    End If
    
    If LastFromCode = IsoBase And LastToCode = IsoTo And DateDiff("n", LastCall, Now) < UpdatePause Then
        ' Return cached rates.
    Else
        ' Retrieve updated rates.
        BulkRates = (IsoTo = "")
        If BulkRates Then
            ' Fetch the full list of exchange rates for IsoBase.
            ReDim UrlPath(UBoundBulk)
            CodePath = AllRatesPath
        Else
            ' Fetch only the exchange rates of IsoTo for IsoBase.
            ReDim UrlPath(UBoundPair)
            CodePath = OneRatePath
        End If
        UrlPath(0) = ServiceUrl
        UrlPath(1) = CodePath
        UrlPath(2) = EraApiId
        UrlPath(3) = IsoBase
        If Not BulkRates Then
            UrlPath(UBoundPair) = IsoTo
        End If
        
        Url = Join(UrlPath, "/")
        ' Uncomment for debugging.
        ' Debug.Print Url
        
        ' Define default result array.
        ' Redim for three dimensions: date, code, rate.
        ReDim Rates(0, 0 To 2)
        Rates(0, RateDetail.Date) = NoValueDate
        Rates(0, RateDetail.Code) = NeutralCode
        Rates(0, RateDetail.Rate) = NeutralRate
        
        If RetrieveDataResponse(Url, ResponseText) = True Then
            Set DataCollection = CollectJson(ResponseText)
        Else
            ' Give up.
            Set DataCollection = Nothing
        End If
    
        If Not DataCollection Is Nothing Then
            If DataCollection(RootNodeName)(CollectionItem.Data)(1)(CollectionItem.Name) = FirstNodeName Then
                ResultCode = DataCollection(RootNodeName)(CollectionItem.Data)(FirstNodeName)(CollectionItem.Data)
                Select Case ResultCode
                    Case ResultSuccess
                        ' Data fetched successfully. Continue.
                    Case ResultError
                        ErrorCode = DataCollection(RootNodeName)(CollectionItem.Data)(ErrorNodeName)(CollectionItem.Data)
                        ' Optional error handling.
                        Select Case ErrorCode
                            Case InvalidKey
                                ' Invalid api key.
                            Case UnknownCode
                                ' Invalid currency code.
                        End Select
                        Set DataCollection = Nothing
                End Select
            Else
                ' Unexpected data.
                Set DataCollection = Nothing
            End If
        End If
        
        If ResultCode = ResultSuccess And Not DataCollection Is Nothing Then
            ' One or all rates were retrieved.
            ' Get the UTC value date and time for the rate(s).
            ValueDate = DateUnix(DataCollection(RootNodeName)(CollectionItem.Data)(TimeNodeName)(CollectionItem.Data))
            
            If BulkRates Then
                ' Get count of rates.
                RateCount = DataCollection(RootNodeName)(CollectionItem.Data)(RateNodeName)(CollectionItem.Data).Count
                ' Redim for three dimensions: date, code, rate.
                ReDim Rates(RateCount - 1, 0 To 2)
                ' Fill the array from the collection items.
                For Index = 1 To RateCount
                    ' A retrieved rate item is an array.
                    RateItem = DataCollection(RootNodeName)(CollectionItem.Data)(RateNodeName)(CollectionItem.Data)(Index)
                    Rates(Index - 1, RateDetail.Date) = ValueDate
                    Rates(Index - 1, RateDetail.Code) = RateItem(CollectionItem.Name)
                    Rates(Index - 1, RateDetail.Rate) = RateItem(CollectionItem.Data)
                Next
            Else
                ' Redim for two items ("from" and "to" currency codes) and three dimensions: date, code, rate.
                ReDim Rates(1, 0 To 2)
                ' "From" currency data.
                Rates(0, RateDetail.Date) = ValueDate
                Rates(0, RateDetail.Code) = DataCollection(RootNodeName)(CollectionItem.Data)(FromNodeName)(CollectionItem.Data)
                Rates(0, RateDetail.Rate) = NeutralRate
                ' "To" currency data.
                ' The retrieved rate item is not an array.
                Rates(1, RateDetail.Date) = ValueDate
                Rates(1, RateDetail.Code) = DataCollection(RootNodeName)(CollectionItem.Data)(ToNodeName)(CollectionItem.Data)
                Rates(1, RateDetail.Rate) = DataCollection(RootNodeName)(CollectionItem.Data)(Rate1NodeName)(CollectionItem.Data)
            End If
            
            Set DataCollection = Nothing
            
            ' Round the call time down to the start of the update interval.
            ThisCall = CDate(Fix(Now * 24 * 60 / UpdatePause) / (24 * 60 / UpdatePause))
            ' Record requested base currency and hour of retrieval.
            LastFromCode = IsoBase
            LastToCode = IsoTo
            LastCall = ThisCall
        End If
    End If
    
    ExchangeRatesEra = Rates

End Function

' Retrieve the current exchange rates from "Fixer" for one base currency.
' The rates are returned as an array and cached until the next update.
' The rates are updated from once per hour down to once per minute.
'
' Default base currency is EUR.
' For the free plan, exchange rates for other base currencies are
' calculated from EUR by triangular calculation.
'
' Source:
'   https://fixer.io/
'   https://fixer.io/documentation
'
' Note:
'   Exchange rates are classed as indicative rates and are accurate enough to display price estimations.
'   The rates are unsuitable for forex trading or processing cross currency settlements.
'
' Example:
'   Dim Rates As Variant
'   Rates = ExchangeRatesFxr()
'   Rates(12, 0) -> 2018-09-20 08:54:06 ' Publishing date.
'   Rates(12, 1) -> "BDT"               ' Currency code.
'   Rates(12, 2) -> 98.26592            ' Exchange rate.
'
' 2018-09-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ExchangeRatesFxr( _
    Optional ByVal IsoBase As String) _
    As Variant
    
    ' Operational constants.
    '
    ' API endpoint for the free plan.
    ' For the paid plans, https may be used.
    Const ServiceUrl    As String = "http://data.fixer.io/api/latest"
    ' Update interval: 60, 10, or 1 minutes.
    Const UpdatePause   As Integer = 60
    
    ' Function constants.
    '
    ' Default base currency code.
    Const DefaultBase   As String = EuroCode
    ' Node names in retrieved collection.
    Const RootNodeName  As String = "root"
    Const TimeNodeName  As String = "timestamp"
    Const RateNodeName  As String = "rates"
    Const FirstNodeName As String = "success"
    Const ErrorNodeName As String = "error"
    Const CodeNodeName  As String = "code"
    ' Error code for invalid or missing access key.
    Const KeyErrorCode  As Long = 101
    ' Error code for restricted access to base currency.
    Const BaseErrorCode As Long = 105
    ' Error code for invalid currency code.
    Const CodeErrorCode As Long = 201
    
    Static Rates()      As Variant
    Static LastCode     As String
    Static LastCall     As Date
    
    Dim DataCollection  As Collection
    
    Dim Parameter()     As String
    Dim Parameters()    As String
    Dim UrlParts(1)     As String
    
    Dim RateCount       As Integer
    Dim RateItem        As Variant
    Dim BaseRate        As Double
    Dim Index           As Integer
    Dim Url             As String
    Dim ResponseText    As String
    Dim ValueDate       As Date
    Dim ThisCall        As Date
    Dim ErrorCode       As Long
    
    If IsoBase = "" Then
        IsoBase = DefaultBase
    End If
    
    If LastCode = IsoBase And DateDiff("n", LastCall, Now) < UpdatePause Then
        ' Return cached rates.
    Else
        ' Retrieve updated rates.
        
        ' Define parameter array.
        ' Redim for two dimensions: name, value.
        ReDim Parameter(0 To 1, 0 To 1)
        ' Parameter names.
        Parameter(0, ParameterDetail.Name) = "access_key"
        Parameter(1, ParameterDetail.Name) = "base"
        ' Parameter values.
        Parameter(0, ParameterDetail.Value) = FxrApiId
        Parameter(1, ParameterDetail.Value) = IsoBase
        
        ' Assemble parameters.
        ReDim Parameters(LBound(Parameter, 1) To UBound(Parameter, 1))
        For Index = LBound(Parameters) To UBound(Parameters)
            Parameters(Index) = Parameter(Index, 0) & "=" & Parameter(Index, 1)
        Next
        
        ' Assemble URL.
        UrlParts(0) = ServiceUrl
        UrlParts(1) = Join(Parameters, "&")
        Url = Join(UrlParts, "?")
        ' Uncomment for debugging.
        ' Debug.Print Url
        
        ' Define default result array.
        ' Redim for three dimensions: date, code, rate.
        ReDim Rates(0, 0 To 2)
        Rates(0, RateDetail.Date) = NoValueDate
        Rates(0, RateDetail.Code) = NeutralCode
        Rates(0, RateDetail.Rate) = NeutralRate
        
        If RetrieveDataResponse(Url, ResponseText) = True Then
            Set DataCollection = CollectJson(ResponseText)
        Else
            ' Give up.
            Set DataCollection = Nothing
        End If
    
        If Not DataCollection Is Nothing Then
            If DataCollection(RootNodeName)(CollectionItem.Data)(1)(CollectionItem.Name) = FirstNodeName Then
                If DataCollection(RootNodeName)(CollectionItem.Data)(FirstNodeName)(CollectionItem.Data) = False Then
                    ErrorCode = DataCollection(RootNodeName)(CollectionItem.Data)(ErrorNodeName)(CollectionItem.Data)(CodeNodeName)(CollectionItem.Data)
                    Select Case ErrorCode
                        Case KeyErrorCode
                            ' Missing or invalid access key.
                            Set DataCollection = Nothing
                        Case CodeErrorCode, BaseErrorCode
                            ' Typical for invalid currency code, or if free license and base <> USD, respectively.
                            ' Rebuld Url to use base = USD.
                            Parameter(1, 1) = DefaultBase
                            ' Reassemble parameters.
                            For Index = LBound(Parameters) To UBound(Parameters)
                                Parameters(Index) = Parameter(Index, 0) & "=" & Parameter(Index, 1)
                            Next
                            
                            ' Reassemble URL.
                            UrlParts(0) = ServiceUrl
                            UrlParts(1) = Join(Parameters, "&")
                            Url = Join(UrlParts, "?")
                            
                            ' Try once more to retrieve the rates.
                            If RetrieveDataResponse(Url, ResponseText) = True Then
                                Set DataCollection = CollectJson(ResponseText)
                                If DataCollection(RootNodeName)(CollectionItem.Data)(FirstNodeName)(CollectionItem.Data) = False Then
                                    ' Give up.
                                    Set DataCollection = Nothing
                                End If
                            End If
                            ' Rebuld Url to use base = USD.
                            Parameter(1, 1) = DefaultBase
                            ' Reassemble parameters.
                            For Index = LBound(Parameters) To UBound(Parameters)
                                Parameters(Index) = Parameter(Index, 0) & "=" & Parameter(Index, 1)
                            Next
                            
                            ' Reassemble URL.
                            UrlParts(0) = ServiceUrl
                            UrlParts(1) = Join(Parameters, "&")
                            Url = Join(UrlParts, "?")
                            
                            ' Try once more to retrieve the rates.
                            If RetrieveDataResponse(Url, ResponseText) = True Then
                                Set DataCollection = CollectJson(ResponseText)
                                If DataCollection(RootNodeName)(CollectionItem.Data)(FirstNodeName)(CollectionItem.Data) = False Then
                                    ' Give up.
                                    Set DataCollection = Nothing
                                End If
                            End If
                    End Select
                End If
            End If
        End If
        
        If Not DataCollection Is Nothing Then
            ' Rates were retrieved.
            ' Get the UTC value date and time for the rates.
            ValueDate = DateUnix(DataCollection(RootNodeName)(CollectionItem.Data)(TimeNodeName)(CollectionItem.Data))
            ' Get count of rates.
            RateCount = DataCollection(RootNodeName)(CollectionItem.Data)(RateNodeName)(CollectionItem.Data).Count
            ' Redim for three dimensions: date, code, rate.
            ReDim Rates(RateCount - 1, 0 To 2)
            BaseRate = NeutralRate
            
            ' Fill the array from the collection items.
            For Index = 1 To RateCount
                ' A retrieved rate item is an array.
                RateItem = DataCollection(RootNodeName)(CollectionItem.Data)(RateNodeName)(CollectionItem.Data)(Index)
                Rates(Index - 1, RateDetail.Date) = ValueDate
                Rates(Index - 1, RateDetail.Code) = RateItem(CollectionItem.Name)
                Rates(Index - 1, RateDetail.Rate) = RateItem(CollectionItem.Data)
                If RateItem(CollectionItem.Name) = IsoBase And RateItem(CollectionItem.Data) <> NeutralRate Then
                    ' Prepare triangular calculation.
                    BaseRate = RateItem(CollectionItem.Data)
                End If
            Next
            If BaseRate <> NeutralRate Then
                For Index = 1 To RateCount
                    ' Perform triangular calculation of the exchange rates.
                    If Rates(Index - 1, RateDetail.Code) = IsoBase Then
                        Rates(Index - 1, RateDetail.Rate) = NeutralRate
                    Else
                        Rates(Index - 1, RateDetail.Rate) = Rates(Index - 1, RateDetail.Rate) / BaseRate
                    End If
                Next
            End If
            
            Set DataCollection = Nothing
            
            ' Round the call time down to the start of the update interval.
            ThisCall = CDate(Fix(Now * 24 * 60 / UpdatePause) / (24 * 60 / UpdatePause))
            ' Record requested base currency and hour of retrieval.
            LastCode = IsoBase
            LastCall = ThisCall
        End If
    End If
    
    ExchangeRatesFxr = Rates

End Function

' Retrieve the current exchange rates from "open exchange rates" for one base currency.
' The rates are returned as an array and cached until the next update.
' The rates are updated from once per hour down to once per minute.
'
' Default base currency is USD.
' For the free plan, exchange rates for other base currencies are
' calculated from USD by triangular calculation.
'
' Source:
'   https://openexchangerates.org/
'   https://docs.openexchangerates.org/
'
' Note:
'   Exchange rates are classed as indicative rates and are accurate enough to display price estimations.
'   The rates are unsuitable for forex trading or processing cross currency settlements.
'
' Example:
'   Dim Rates As Variant
'   Rates = ExchangeRatesOxr()
'   Rates(12, 0) -> 2018-09-20 12:00:00 ' Publishing date.
'   Rates(12, 1) -> "BDT"               ' Currency code.
'   Rates(12, 2) -> 84.064038           ' Exchange rate.
'
' 2018-09-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ExchangeRatesOxr( _
    Optional ByVal IsoBase As String) _
    As Variant
    
    ' Operational constants.
    '
    ' API endpoint.
    Const ServiceUrl    As String = "https://openexchangerates.org/api/latest.json"
    ' Update interval: 60, 30, or 5 minutes.
    Const UpdatePause   As Integer = 60
    
    ' Function constants.
    '
    ' Default base currency code.
    Const DefaultBase   As String = USDollarCode
    ' Node names in retrieved collection.
    Const RootNodeName  As String = "root"
    Const TimeNodeName  As String = "timestamp"
    Const RateNodeName  As String = "rates"
    
    Static Rates()      As Variant
    Static LastCode     As String
    Static LastCall     As Date
    
    Dim DataCollection  As Collection
    
    Dim Parameter()     As String
    Dim Parameters()    As String
    Dim UrlParts(1)     As String
    
    Dim RateCount       As Integer
    Dim RateItem        As Variant
    Dim BaseRate        As Double
    Dim Index           As Integer
    Dim Url             As String
    Dim ResponseText    As String
    Dim ValueDate       As Date
    Dim ThisCall        As Date
    
    If IsoBase = "" Then
        IsoBase = DefaultBase
    End If
    
    If LastCode = IsoBase And DateDiff("n", LastCall, Now) < UpdatePause Then
        ' Return cached rates.
    Else
        ' Retrieve updated rates.
        
        ' Define parameter array.
        ' Redim for two dimensions: name, value.
        ReDim Parameter(0 To 1, 0 To 1)
        ' Parameter names.
        Parameter(0, ParameterDetail.Name) = "app_id"
        Parameter(1, ParameterDetail.Name) = "base"
        ' Parameter values.
        Parameter(0, ParameterDetail.Value) = OxrApiId
        Parameter(1, ParameterDetail.Value) = IsoBase
        
        ' Assemble parameters.
        ReDim Parameters(LBound(Parameter, 1) To UBound(Parameter, 1))
        For Index = LBound(Parameters) To UBound(Parameters)
            Parameters(Index) = Parameter(Index, 0) & "=" & Parameter(Index, 1)
        Next
        
        ' Assemble URL.
        UrlParts(0) = ServiceUrl
        UrlParts(1) = Join(Parameters, "&")
        Url = Join(UrlParts, "?")
        ' Uncomment for debugging.
        ' Debug.Print Url
        
        ' Define default result array.
        ' Redim for three dimensions: date, code, rate.
        ReDim Rates(0, 0 To 2)
        Rates(0, RateDetail.Date) = NoValueDate
        Rates(0, RateDetail.Code) = NeutralCode
        Rates(0, RateDetail.Rate) = NeutralRate
        
        If RetrieveDataResponse(Url, ResponseText) = True Then
            Set DataCollection = CollectJson(ResponseText)
        Else
            ' Check error codes, and requery if possible.
            Select Case Left(ResponseText, 3)
                Case HttpStatus.Unauthorized
                    ' Invalid app_id.
                Case HttpStatus.Forbidden
                    ' Free license and base <> USD.
                    ' Rebuld Url to use base = USD.
                    Parameter(1, 1) = DefaultBase
                    ' Reassemble parameters.
                    For Index = LBound(Parameters) To UBound(Parameters)
                        Parameters(Index) = Parameter(Index, 0) & "=" & Parameter(Index, 1)
                    Next
                    
                    ' Reassemble URL.
                    UrlParts(0) = ServiceUrl
                    UrlParts(1) = Join(Parameters, "&")
                    Url = Join(UrlParts, "?")
                    
                    ' Try once more to retrieve the rates.
                    If RetrieveDataResponse(Url, ResponseText) = True Then
                        Set DataCollection = CollectJson(ResponseText)
                    End If
            End Select
        End If
    
        If Not DataCollection Is Nothing Then
            ' Rates were retrieved.
            ' Get the UTC value date and time for the rates.
            ValueDate = DateUnix(DataCollection(RootNodeName)(CollectionItem.Data)(TimeNodeName)(CollectionItem.Data))
            ' Get count of rates.
            RateCount = DataCollection(RootNodeName)(CollectionItem.Data)(RateNodeName)(CollectionItem.Data).Count
            ' Redim for three dimensions: date, code, rate.
            ReDim Rates(RateCount - 1, 0 To 2)
            BaseRate = NeutralRate
    
            ' Fill the array from the collection items.
            For Index = 1 To RateCount
                ' A retrieved rate item is an array.
                RateItem = DataCollection(RootNodeName)(CollectionItem.Data)(RateNodeName)(CollectionItem.Data)(Index)
                Rates(Index - 1, RateDetail.Date) = ValueDate
                Rates(Index - 1, RateDetail.Code) = RateItem(CollectionItem.Name)
                Rates(Index - 1, RateDetail.Rate) = RateItem(CollectionItem.Data)
                If RateItem(CollectionItem.Name) = IsoBase And RateItem(CollectionItem.Data) <> NeutralRate Then
                    ' Prepare triangular calculation.
                    BaseRate = RateItem(CollectionItem.Data)
                End If
            Next
            If BaseRate <> NeutralRate Then
                For Index = 1 To RateCount
                    ' Perform triangular calculation of the exchange rates.
                    If Rates(Index - 1, RateDetail.Code) = IsoBase Then
                        Rates(Index - 1, RateDetail.Rate) = NeutralRate
                    Else
                        Rates(Index - 1, RateDetail.Rate) = Rates(Index - 1, RateDetail.Rate) / BaseRate
                    End If
                Next
            End If
            
            Set DataCollection = Nothing
            
            ' Round the call time down to the start of the update interval.
            ThisCall = CDate(Fix(Now * 24 * 60 / UpdatePause) / (24 * 60 / UpdatePause))
            ' Record requested base currency and hour of retrieval.
            LastCode = IsoBase
            LastCall = ThisCall
        End If
    End If
    
    ExchangeRatesOxr = Rates

End Function

' Retrieve the current exchange rates from "XE" for one base currency.
' The rates are returned as an array and cached until the next update.
' The rates are updated from once per day down to once per minute.
'
' Default base currency is USD.
' For the free plan, exchange rates for other base currencies are
' calculated from USD by triangular calculation.
'
' Source:
'   https://www.xe.com/
'   https://www.xe.com/xecurrencydata/
'
' Note:
'   Exchange rates are live mid-market rates, which are not available to
'   consumers and are for informational purposes only.
'
' Example:
'   Dim Rates As Variant
'   Rates = ExchangeRatesXec()
'   Rates(12, 0) -> 2018-10-12 00:00:00 ' Publishing date.
'   Rates(12, 1) -> "BDT"               ' Currency code.
'   Rates(12, 2) -> 83.7886823907       ' Exchange rate.
'
' 2018-10-12. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ExchangeRatesXec( _
    Optional ByVal IsoBase As String) _
    As Variant
    
    ' Operational constants.
    '
    ' API endpoint.
    Const ServiceUrl    As String = "https://xecdapi.xe.com/v1/convert_from/"
    ' Update interval: 60, 30, or 5 minutes.
    Const UpdatePause   As Integer = 60
    
    ' Function constants.
    '
    ' Default base currency code.
    Const DefaultBase   As String = USDollarCode
    ' Node names in retrieved collection.
    Const RootNodeName  As String = "root"
    Const TimeNodeName  As String = "timestamp"
    Const RateNodeName  As String = "to"
    Const CodeNodeName  As String = "quotecurrency"
    Const ValueNodeName As String = "mid"
    
    Static Rates()      As Variant
    Static LastCode     As String
    Static LastCall     As Date
    
    Dim DataCollection  As Collection
    
    Dim Parameter()     As String
    Dim Parameters()    As String
    Dim UrlParts(1)     As String
    
    Dim UserName        As String
    Dim Password        As String
    
    Dim RateCount       As Integer
    Dim RateItem        As Variant
    Dim BaseRate        As Double
    Dim Index           As Integer
    Dim Url             As String
    Dim ResponseText    As String
    Dim ValueDate       As Date
    Dim ThisCall        As Date
    
    If IsoBase = "" Then
        IsoBase = DefaultBase
    End If
    
    If LastCode = IsoBase And DateDiff("n", LastCall, UtcNow) < UpdatePause Then
        ' Return cached rates.
    Else
        ' Retrieve updated rates.
        
        ' Define parameter array.
        ' Redim for two dimensions: name, value.
        ReDim Parameter(0 To 1, 0 To 1)
        ' Parameter names.
        Parameter(0, ParameterDetail.Name) = "from"
        Parameter(1, ParameterDetail.Name) = "to"
        ' Parameter values.
        Parameter(0, ParameterDetail.Value) = IsoBase
        Parameter(1, ParameterDetail.Value) = "*"
        
        ' Assemble parameters.
        ReDim Parameters(LBound(Parameter, 1) To UBound(Parameter, 1))
        For Index = LBound(Parameters) To UBound(Parameters)
            Parameters(Index) = Parameter(Index, 0) & "=" & Parameter(Index, 1)
        Next
        
        ' Assemble URL.
        UrlParts(0) = ServiceUrl
        UrlParts(1) = Join(Parameters, "&")
        Url = Join(UrlParts, "?")
        ' Uncomment for debugging.
        ' Debug.Print Url
        
        ' Credentials.
        UserName = XeAccount
        Password = XeApiId
        
        ' Define default result array.
        ' Redim for three dimensions: date, code, rate.
        ReDim Rates(0, 0 To 2)
        Rates(0, RateDetail.Date) = NoValueDate
        Rates(0, RateDetail.Code) = NeutralCode
        Rates(0, RateDetail.Rate) = NeutralRate
                
        If RetrieveDataResponse(Url, ResponseText, , UserName, Password) = True Then
            Set DataCollection = CollectJson(ResponseText)
        Else
            ' Check error codes.
            Select Case Left(ResponseText, 3)
                Case HttpStatus.Forbidden
                    ' Invalid credentials.
            End Select
            ' No rates were received.
            Set DataCollection = Nothing
        End If
    
        If Not DataCollection Is Nothing Then
            ' Rates were retrieved.
            ' Get the UTC value date and time for the rates.
            ValueDate = DateIso8601(DataCollection(RootNodeName)(CollectionItem.Data)(TimeNodeName)(CollectionItem.Data))
            ' Get count of rates.
            RateCount = DataCollection(RootNodeName)(CollectionItem.Data)(RateNodeName)(CollectionItem.Data).Count
            ' Redim for three dimensions: date, code, rate.
            ReDim Rates(RateCount - 1, 0 To 2)
            BaseRate = NeutralRate
    
            ' Fill the array from the collection items.
            For Index = 1 To RateCount
                ' A retrieved rate item is yet a collection with an array.
                RateItem = DataCollection(RootNodeName)(CollectionItem.Data)(RateNodeName)(CollectionItem.Data)(Index)
                Rates(Index - 1, RateDetail.Date) = ValueDate
                Rates(Index - 1, RateDetail.Code) = RateItem(CollectionItem.Data)(CodeNodeName)(CollectionItem.Data)
                Rates(Index - 1, RateDetail.Rate) = RateItem(CollectionItem.Data)(ValueNodeName)(CollectionItem.Data)
            Next
            
            Set DataCollection = Nothing
            
            ' Round the call time down to the start of the update interval.
            ThisCall = CDate(Fix(Now * 24 * 60 / UpdatePause) / (24 * 60 / UpdatePause))
            ' Record requested base currency and hour of retrieval.
            LastCode = IsoBase
            LastCall = ThisCall
        End If
    End If
    
    ExchangeRatesXec = Rates

End Function

' Returns True if the passed Index is a key of Collection.
'
' 2018-09-22. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsCollectionItem( _
    ByRef Collection As VBA.Collection, _
    ByVal Index As Variant) _
    As Boolean
    
    Const InvalidProcedureArgument  As Long = 5
    
    Dim Item    As Variant
    Dim Result  As Boolean
    
    On Error GoTo Err_IsCollectionItem
    
    If Not Collection Is Nothing Then
        Item = Collection.Item(Index)
        Result = Not Collection Is Nothing
    End If
    
    IsCollectionItem = Result
    
Exit_IsCollectionItem:
    Exit Function
    
Err_IsCollectionItem:
    Select Case Err.Number
        Case InvalidProcedureArgument
            ' Key is not present in Collection.
        Case Else
            ' Other error. Ignore.
    End Select
    Resume Exit_IsCollectionItem
    
End Function

