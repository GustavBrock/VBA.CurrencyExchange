# VBA.CurrencyExchange #
![Help](https://raw.githubusercontent.com/GustavBrock/VBA.CurrencyExchange/master/images/EE%20CurrencyExchange.png)

## Currency Exchange Rates ##

### Services ###
Exchange rates can be obtained from many sources, some free, some paid.

Hardly two of these serve the same purpose or are targeted the same users. This means, that some research typically is necessary to pick the service that will fit a given scenario and demand.

The services, that this project addresses, are:

1. The European Central Bank
2. The Danish National Bank
3. Currency Converter API
4. Currencylayer API
5. ExchangeRate API
6. Fixer
7. Open Exchange Rates

All services support the currencies commonly used in international trade; for more exotic currencies, you may be limited in the choice of service.

For free, a few services provides exchange rates from any base currency, some provide exchange rates based on one currency only, some only one or a few currencies based on any currency, and one provides exchange rates *to* one currency only (Euro, The European Central Bank).

For those services that - for free or by design - offer only one base currency, exchange rates can still be retrieved for other base currencies, though not directly (of course) but, automatically, by triangular calculation against the supported base currency.


#### Important:
> The exchange rates published by the services are what is called *mid-market rates*.
> 
> This means, that they cannot be used for real transactions; for such, you must refer to the actual buying and selling rates of your bank or broker. 


### Functions ###
Like the services differ in offerings, so do the various APIs or download options, though only two basic techniques are used:

* an API, delivering data as Json
* An XML document

However, no two services - even using the same basic technique - offer the same data format; thus a custom function is required for each service.

The main functions offered are named:

	ExchangeRatesXyz

where **Xys** is a three-letter abbreviation of the service name.

These functions each downloads and returns an array with the rates, and also attempts to cache the download for two reasons:

- to speed up reading the rates multiple times
- to save the usage of and the load on the service

These functions are supplemented with a set of matching functions for converting an amount from one currency to another. These are named in a similar way:

	CurrencyConvertXyz

These functions each utilises the output from the corresponding *ExchangeRatesXyz* function. Further, they cache the conversion factor for a set of currencies to speed up the calculation of many amounts between the same two currencies.

All functions support the *neutral currency code* **XXX** for an exchange rate of **1**.


### Code ###
Where relevant, all functions support both early and late binding. Code has been tested with both 32-bit and 64-bit *Microsoft Access 2016* and *365*.

It requires the Json modules from the project [VBA.CVRAPI](https://github.com/CactusData/VBA.CVRAPI).

### Documentation ###
Full documentation can be found here:

![EE Logo](https://raw.githubusercontent.com/GustavBrock/VBA.CurrencyExchange/master/images/EE%20Logo.png) 

[Exchange Rates and Currency Conversion in VBA](https://www.experts-exchange.com/articles/33199/Exchange-Rates-and-Currency-Conversion-in-VBA.html)

Included is a Microsoft Access example application.