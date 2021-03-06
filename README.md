# VBA.CurrencyExchange #
![Help](https://raw.githubusercontent.com/GustavBrock/VBA.CurrencyExchange/master/images/EE%20CurrencyExchange.png)

## Currency Exchange Rates ##

### Services ###
Exchange rates can be obtained from many sources, some free, some paid.

Hardly two of these serve the same purpose or are targeted the same users. This means, that some research typically is necessary to pick the service that will fit a given scenario and demand.

The services, that this project addresses, are:

1. The European Central Bank
2. The Danish National Bank
3. The Central Bank of the Russian Federation
4. Currency Converter API
5. Currencylayer API
6. ExchangeRate API
7. Fixer
8. Open Exchange Rates
9. php.mk (National Bank of the Republic of North Macedonia)
10. XE

All services support the currencies commonly used in international trade; for more exotic currencies, you may be limited in the choice of service.

For free, a few services provide exchange rates from any base currency, some provide exchange rates based on one currency only, some only one or a few currencies based on any currency, and one provides exchange rates *to* one currency only (Euro, The European Central Bank). One service, XE, offers *no free plan* at all.

For those services that - for free or by design - offer only one base currency, exchange rates can still be retrieved for other base currencies, though not directly (of course) but, automatically, by triangular calculation against the supported base currency.

URLs for the services and (where needed) their documentation can be found in the in-line documentation.


#### Important:
> The exchange rates published by the services are what is called *mid-market rates*.
> 
> This means, that they cannot be used for real transactions; for such, you must refer to the actual buying and selling rates of your bank or broker. 


### Functions ###
Like the services differ in offerings, so do the various APIs or download options, though only three basic techniques are used:

* addressing an API, delivering data as Json
* reading an XML document
* parsing an HTML document (web scraping, data extracting)

However, no two services - even using the same basic technique - offer the same data format; thus a custom function is required for each service.

The main functions offered are named:

	ExchangeRatesXyz

where **Xyz** is a three-letter abbreviation of the service name.

Each of these functions returns an array with the rates, and also attempts to cache the download for two reasons:

- to speed up reading the rates multiple times
- to save the usage of and the load on the service

The functions are supplemented with a set of matching functions for converting an amount from one currency to another. These are named in a similar way:

	CurrencyConvertXyz

These functions each utilises the output from the corresponding *ExchangeRatesXyz* function. Further, they cache the conversion factor for a set of currencies to speed up the calculation of many amounts between the same two currencies.

All functions support the *neutral currency code* **XXX** for an exchange rate of **1**.


### Code ###
Where relevant, all functions support both early and late binding. Code has been tested with both 32-bit and 64-bit *Microsoft Access and Excel 2019* and *365*.

It requires the Json modules from the project [VBA.CVRAPI](https://github.com/CactusData/VBA.CVRAPI).

### Documentation ###
Full documentation can be found here:

![EE Logo](https://raw.githubusercontent.com/GustavBrock/VBA.CurrencyExchange/master/images/EE%20Logo.png) 

[Exchange Rates and Currency Conversion in VBA](https://www.experts-exchange.com/articles/33199/Exchange-Rates-and-Currency-Conversion-in-VBA.html)

Included is a Microsoft Access example application and a Microsoft Excel example workbook.

<hr>

*If you wish to support my work or need extended support or advice, feel free to:*

<p>

[<img src="https://raw.githubusercontent.com/GustavBrock/VBA.CurrencyExchange/master/images/BuyMeACoffee.png">](https://www.buymeacoffee.com/gustav/)