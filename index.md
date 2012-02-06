---
layout: default
title: FinAnSu
---

#summary A table of contents for FinAnSu functionality

  * Getting Started
    * [Introduction]
    * [Installation]
  * Web Import Functions
    * [Using Array Formulas](ArrayFormulas)
    * [Real-Time Quotes](Quotes)
    * [Historical Quotes](QuoteHistory)
    * [Web Data](WebData)
  * Options Math Functions
    * [Black-Scholes](BlackScholes)
    * [Greeks]
    * [American/Bermudan Options](AmericanBermudan)
    * [Forward Rate Agreements](FRAs)
  * Yield Curve Functions
    * [Interpolation]
  * Other Functions
    * [Sorting]
    * [Color Conversion](Colors)
    * [Distribution Functions](Distributions)
    * [Version Information](Version)
  * Formatting Macros
    * [Currencies]
    * [Layout]
  * Credits & Licensing
    * [Credits]
    * [License]

#summary An introduction to FinAnSu.

Purpose
-------

!FinAnSu aims to provide user-friendly tools for use in financial applications. The add-in is in its development stages, but it currently offers:

  * *Live, streaming [web import](WebData) capabilities* (Excel 2002+), including custom functions for easily [importing security prices](Quotes) from Bloomberg.com, Google Finance and Yahoo! Finance
  * Functions that can [parse CSVs](WebData#ImportCSV) on the web, including custom functions to easily import stock quote data from [Google Finance](Quotes#GoogleHistory) and instrument data from the [Fed H.15 release](Quotes#H15History) (yields for Fed Funds, commercial paper, Treasuries, interest-rate swaps, etc.).
  * Basic options pricing, including [Black-Scholes](BlackScholes), [the options Greeks](Greeks) and [a few more complex options models](AmericanBermudan)
  * A bit on [FRAs] (forward rate agreements)
  * A continuously compounded rate [interpolator](Interpolation) (need to add more: linear, cubic, cubic spline, etc.)
  * A Federal Reserve holiday calculator (need to add more for different calendars, as well as roll date conventions)
  * A handful of tools for converting between discount factors and forward rates
  * Function to [automatically sort](Sorting) ranges that contain rows and columns of data
  * Some formatting macros for [currencies](Currencies) and [a few layout options](Layout) that aren't easily accessible in Excel

  http://finansu.googlecode.com/hg/img/quote.gif

----

Frequently Asked Questions
--------------------------

#### How do I `[`do such-and-such`]` ?

  First, try using the function wizard in Excel. All functions have descriptions of what values they return and accept. That is, type something like `=Quote(` in a cell in Excel, then hit the _f,,x,,_ key just above the worksheet (or press `Ctrl+A`).

  If that's confusing, check the documentation here and the [Examples worksheet](http://finansu.googlecode.com/hg/FinAnSu/Examples.xls).

  If none of that helps or if you have any questions or suggestions for clarity, go ahead and email me. My contact information is [below](#Contact_Information).

How long will you support this?
-------------------------------

  Indefinitely. I developed this add-in in my spare time to address a real professional need, and I use it daily. That said, there are two reasons for a loss of functionality in the quote import functions specifically, neither of which I control:

  * *Temporary:* Bloomberg, Google or Yahoo! change the setup of their website such that I have to change how I parse data from them. This is extremely rare; I anticipate it happening maybe once every few years. When it does, it may take a _brief_ while before I notice and update the program.

  * *Permanent:* Bloomberg, Google or Yahoo! stop publishing financial data publicly. That said, if you change the `source` parameter for functions like [Quote](Quotes#Quote) or [QuoteHistory](QuoteHistory#QuoteHistory), very often you can find an alternate source. Of course, losing _all three_ of them would be a devastating [Black Swan](http://en.wikipedia.org/wiki/Black_swan_theory). You know, one of those "once-in-a-million-years" events (i.e. a truly unpredictable misfortune for which we assume a ridiculously optimistic probability of avoidance).

  There may also be some disruptions in different Excel or .NET versions, but I _think_ they will be minimal.

#### Why is [Quote](Quotes#Quote) or [QuoteHistory](Quotes#QuoteHistory) only returning one value?

  See [the section on array formulas](ArrayFormulas).

#### Can I use this at work, on other computer, etc.?

  Hopefully. I'm unfamiliar with the access restrictions at different companies, but in general if you meet the [minimum requirements](#Requirements) you should be fine. If there are real access limitations, feel free to [inform me](#Contact_Information), but I don't know how much I can do about it.

  Also, this application does _not_ transmit any usage data to me or even connect to any servers owned by me. Feedback is always appreciated, but I'm not collecting it behind the scenes.

#### What do I do if I notice an error?

  Either [email me](#Contact_Information) or [enter a new issue](http://code.google.com/p/finansu/issues/entry). There are some very real errors out there, and in the long-run the fix is _usually_ better solved on my end.

----

Requirements
------------

  * [Microsoft Excel for Windows](http://office.microsoft.com/excel/)
  * [Microsoft .NET Framework 4](http://www.microsoft.com/downloads/details.aspx?FamilyID=9cfb2d51-5ff4-4491-b0e5-b386f32c0992)
  * The ability to install add-ins. There shouldn't be any problems unless macro security is set to high or you have a _very_ restrictive policy regarding writing to your `%AppData%` folder.

----

Installation
============

  1. Close Excel.
  2. Download the [latest zip file from the Downloads tab](https://code.google.com/p/finansu/downloads/list).
  3. Unzip the files to a safe location.
  4. There are two ways to install !FinAnSu: the automatic way that uses the installation batch file, or a simple copy-and-paste of the .xll file into the appropriate directory.
     * *Automatic:* Run `install.bat` (closing Excel if prompted).
        http://finansu.googlecode.com/hg/img/run_install_bat.png
     * *Manual:* Copy `FinAnSu.xll` to `%AppData%\Microsoft\AddIns`. _(Note: You can type that into the `Start Menu > Run...` window or the location bar of any Folder Explorer window.)_
        http://finansu.googlecode.com/hg/img/gotoappdata.png
  5. If this is your first install, open Excel and do the following:
     * *Excel 2007 and later*
       * Click `Office Button > Excel Options... > Add-Ins`.
       * In the `Manage:` drop-down at the bottom, click `Excel Add-ins` and `Go...` .
       * Place a checkmark next to !FinAnSu.
     * *Excel 2003 and earlier*
       * Go to `Tools > Add-Ins`.
       * Place a checkmark next to !FinAnSu.
  6. If you want to check the installation, open the `Examples.xls` spreadsheet. If the functions are returning values, you're all set!

  #summary How to write array formulas.

Many import functions require the use of array formulas. Microsoft has a good, long-form [explanation of array formulas](http://office.microsoft.com/en-us/excel-help/introducing-array-formulas-in-excel-HA001087290.aspx) that I recommend consulting if you'd like more information. Many (most?) Excel users never receive exposure to array formulas, but they are immensely useful for complex calculations.

What They Are
-------------

Instead of returning a single value, some complex Excel functions are capable of returning _multiple values_ to _multiple cells_. This helps save a lot of processor time as you only calculate the results once, and !FinAnSu uses it for several functions. For example, instead of pulling data from Google Finance to calculate one date, then doing the same and calculating the opening price, then repeating for every day you specify, !FinAnSu downloads the necessary information once, parses it, and outputs all results to a range of cells.

Using these formulas requires a slightly different input method. First, you select the cells you want the formula to apply to, then you type the formula, and then you hit `Ctrl+Shift+Enter`.

An Example
----------

Below is an example using [GoogleHistory](Quotes#GoogleHistory).

  # *Select a range of cells*, for example cells `A1` through `F5`.
  # *Type your formula.* In this case you would type something like `=GoogleHistory("GOOG")`.
  # *Hold down `Ctrl+Shift` and hit `Enter`* _(note: without `Ctrl+Shift+Enter` it will just be a normal formula that for `GoogleHistory` just returns the most recent business day)_.
  # This should return the most recent five business days of price data for [GOOG](http://www.google.com/finance?q=GOOG), which includes date, open, high, low, close and volume. You can verify the numbers by looking at http://www.google.com/finance/historical?q=NASDAQ:GOOG. Also, the formula in your formula bar should look like `{=GoogleHistory("GOOG")}` (note the curly braces).

http://finansu.googlecode.com/hg/img/array_formula.gif

#summary FinAnSu's web import functionality for security quotes.

Note: Before proceeding, please read the section on [array formulas](ArrayFormulas) if you are unfamiliar with their usage.

`Quote`
-------

Returns the current quote for a security ID from Bloomberg, Google or Yahoo!

```
=Quote(security_id, source, params, live_updating, frequency, show_headers)

// Returns the current price of WFC from Bloomberg, updated every 15 seconds
// All functions are effectively identical
=LiveQuote("WFC")
=Quote("WFC", , , true)
=Quote("WFC", "b", , true, 15)
=BloombergQuote("WFC", , true, 15)
=LiveQuote("WFC", "b", , 15)

// Returns the current price, change and % change for WFC, including headers (static)
=Quote("WFC", "b", "px%", , , true)
=Quote("WFC", "b", QuoteParams(true, true, true), , , true)
```

  * `security_id` is the security ID from the quote service.
  * `source` is the name or abbreviation of the quote service (`"b"`, `"Bloomberg"`, `"g"`, `"Google"`, `"y"`, `"Yahoo"`, etc.). Defaults to `"Bloomberg"`.
  * `params` is a list of which values to return. Accepts any combination of `"px%dtbahlv"` (price, change, % change, date, time, bid, ask, high, low, volume). Bid/ask are not available through Google. Use [=QuoteParams()](#QuoteParams) for help if necessary. Defaults to price.
  * `live_updating` is whether you want this function to return continuously stream live quotes to the cell. Defaults to `false`.
  * `frequency` is the number of seconds between update requests (if live_updating is `true`). Defaults to `15` seconds.
  * `show_headers` is whether to display the headers for each column. Defaults to `false`.

http://www.brymck.com/images/finansu_live_quote.gif

`QuoteParams`
-------------

Builds a text string for use in [=Quote()](#Quote) designating which values you would like returned.

```
=QuoteParams(price, change, pct_change, date, time, bid, ask, open, high, low, volume)

// Returns the appropriate text string for use in =Quote(),
// such that values for price, change and % change will be returned
=QuoteParams(true, true, true)
```

  * `price` is whether you wish to return the current price or value.
  * `change` is whether you wish to return the day's change.
  * `pct_change` is whether you wish to return the day's percentage change.
  * `date` is whether you wish to return the latest trade date.
  * `time` is whether you wish to return the latest trade time.
  * `bid` is whether you wish to return the current bid price.
  * `ask` is whether you wish to return the current ask price.
  * `open` is whether you wish to return the day's opening price.
  * `high` is whether you wish to return the day's high price.
  * `low` is whether you wish to return the day's low price.
  * `volume` is whether you wish to return the day's closing price.

`LiveQuote`
-----------

Same as [Quote](#Quote) with the `live_updating` argument equal to `true`.

`BloombergQuote`
----------------

Same as [Quote](#Quote) with the `source` argument equal to `"Bloomberg"`.

`GoogleQuote`
-------------

Same as [Quote](#Quote) with the `source` argument equal to `"Google"`.

`YahooQuote`
------------

Same as [Quote](#Quote) with the `source` argument equal to `"Yahoo"`.

`FullTicker`
------------

Returns !FinAnSu's interpretation of an abbreviated security ID. Mostly for debugging purposes.

```
=FullTicker(security_id, source, force_interpret)
=FullTicker("WFC", "b", false) // returns "WFC"
```

  * `security_id` is the security ID from the quote service.
  * `source` is the name or abbreviation of the quote service (`"b"`, `"Bloomberg"`, `"g"`, `"Google"`, `"y"`, `"Yahoo"`, etc.). Defaults to `"Bloomberg"`.
  * `force_interpret` forces !FinAnSu to guess at a suffix if none exists if `force_interpret` is set to `true` (may result in errors).

`ShortenSource`
---------------

Returns FinAnSu's intepretation of an abbreviated source name. Mostly for debugging purposes.

```
=ShortenSource(source)
=ShortenSource("bloomberg") // returns "b"
```

  * `source` is the name or abbreviation of the quote service.

<table>
  <thead>
    <tr>
      <th>Input</th>
      <th>Output</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>(blank)</td>
      <td>"b"</td>
    </tr>
    <tr>
      <td>"b"</td>
      <td>"b"</td>
    </tr>
    <tr>
      <td>"bb"</td>
      <td>"b"</td>
    </tr>
    <tr>
      <td>"bberg"</td>
      <td>"b"</td>
    </tr>
    <tr>
      <td>"bloomberg"</td>
      <td>"b"</td>
    </tr>
    <tr>
      <td>"g"</td>
      <td>"g"</td>
    </tr>
    <tr>
      <td>"goog"</td>
      <td>"g"</td>
    </tr>
    <tr>
      <td>"google"</td>
      <td>"g"</td>
    </tr>
    <tr>
      <td>"y"</td>
      <td>"y"</td>
    </tr>
    <tr>
      <td>"yhoo"</td>
      <td>"y"</td>
    </tr>
    <tr>
      <td>"yahoo"</td>
      <td>"y"</td>
    </tr>
    <tr>
      <td>"yahoo!"</td>
      <td>"y"</td>
    </tr>
  </tbody>
</table>

#summary FinAnSu's web import functionality for quote history.

Note: Before proceeding, please read the section on [array formulas](ArrayFormulas) if you are unfamiliar with their usage.

`QuoteHistory`
--------------

Returns the historical date, open, high, low, close, volume and adjusted price for a security ID from the selected quotes provider.

```
=QuoteHistory(security_id, source, start_date, end_date, period, names, show_headers, date_order)

// Returns date, OHLC, volume and adjusted close for WFC from Yahoo! Finance
// for the past year
=QuoteHistory("WFC")

// Returns monthly date and adjusted close for WFC from Yahoo! Finance
// for the past year
=QuoteHistory("WFC", "y" , , , "m", "da", false, false)
=YahooHistory("WFC", , , "m", "da", false, false)

// Returns the date, close and volume for WFC from Google Finance
// for each day in 2010, in chronological order, with headers
=QuoteHistory("WFC", "g", DATE(2010, 1, 1), DATE(2010, 12, 31), "d", "dcv", true, true)
=GoogleHistory("WFC", DATE(2010, 1, 1), DATE(2010, 12, 31), "d", "dcv", true, true)
```

  * `security_id` is the security ID from the quote service.
  * `source` is the name or abbreviation of the quote service (g, Google, y, Yahoo, etc.). Defaults to Yahoo!.
  * `start_date` is the date from which to start retrieving history. Defaults to the most recent close.
  * `end_date` is the date at which to stop retrieving history. Defaults to one year ago.
  * `period` is a text flag representing whether you want daily (`"d"`), weekly (`"w"`), monthly (`"m"`) or yearly (`"y"`) quotes. Monthly and yearly quotes are available only through Yahoo!. Defaults to daily. Defaults to `"d"`.
  * `names` is a list of which values to return. Accepts any combination of `"dohlcva"` (date, open, high, low, close, volume, adj price). Adj price is available only through Yahoo!. Use [=QuoteHistoryParams()](#QuoteHistoryParams) for help if necessary. Defaults to all.
  * `show_headers` is whether to display the headers for each column. Defaults to `false`.
  * `date_order` is whether to sort dates in ascending chronological order. Defaults to `false`.

`QuoteHistoryParams`
--------------------

Builds a text string for use in [=QuoteHistory()](#QuoteHistory) designating which values you would like returned.

```
=QuoteHistoryParams(date, open, high, low, close, volume, adj_close)

// Returns the appropriate text string for use in =QuoteHistory(),
// such that values for date, close and volume will be returned
=QuoteHistoryParams(true, , , , true, true)
```

  * `date` is whether you wish to return the date.
  * `open` is whether you wish to return the day's opening price.
  * `high` is whether you wish to return the day's high price.
  * `low` is whether you wish to return the day's lowest price.
  * `close` is whether you wish to return the day's closing price.
  * `volume` is whether you wish to return the day's volume.
  * `adj_close` is whether you wish to return the day's closing price.


`YahooHistory`
--------------

Same as [QuoteHistory](#QuoteHistory) with the `source` argument equal to `"yahoo"`.

`GoogleHistory`
---------------

Same as [QuoteHistory](#QuoteHistory) with the `source` argument equal to `"google"`. Note that Google does not contain easily accessible data for things like indexes (such as the S&P 500). If you require such information, I recommend using [YahooHistory](#YahooHistory) instead.

`H15History`
------------

Returns information from the Fed's [http://www.federalreserve.gov/releases/h15/update/ H.15 Statistical Release].

```
=H15History(instrument_id, frequency)
=H15History("AAA_NA", "m") // returns a list of dates and yields for Aaa corporate bonds
```

  * `instrument_id` is the instrument ID. Go to [and click on one of the data links. The ID is in the URL in the form `H15_[id](http://www.federalreserve.gov/releases/h15/data.htm]).txt`.
  * `frequency` is business day (`"b"`), daily (`"d"`), weekly Wednesday (`"ww"`), weekly Thursday (`"wt"`), weekly Friday (`"wf"`), bi-weekly (`"bw"`), monthly (`"m"`) or annual (`"a"`). Not all frequencies are available for all instruments. Defaults to business day (`"b"`)."

#summary Retrieving generic web data.

Note: Before proceeding, please read the section on [array formulas](ArrayFormulas) if you are unfamiliar with their usage.

`Import`
--------

Returns a horizontal array of values based on a URL and regular expression.

```
=Import(url, pattern, max_length, live_updating, frequency)

// Returns the title of the top story from bloomberg.com
=Import("http://www.bloomberg.com/", "story_link[^>]+>(.*?)<", 1, false) 
```

  * `url` is the full URL of the target website.
  * `pattern` is a [regular expression pattern](http://www.regular-expressions.info/quickstart.html) where the first backreference (in parentheses) is the value you wish to retrieve.
  * `max_length` is the maximum length of the results array.
  * `live_updating` is whether you want this function to return continuously stream live quotes to the cell.
  * `frequency` is the number of seconds between update requests (if live_updating is `true`). Defaults to `15` seconds.

`GetWebData`
------------

Same as [Import](#Import) with the `live_updating` argument equal to `false`.

`ImportCSV`
-----------

Returns an array of values from a CSV.

```
=ImportCSV(url, start_line, reverse, formats)

// Returns a list of market sectors and security types
=ImportCSV("http://bsym.bloomberg.com/sym/pages/security_type.csv", 1, false, {"string", "string"})
```

  * `url` is the URL of the target CSV file.
  * `start_line` is the first line of the CSV to begin parsing (starting with `0`).
  * `reverse` is whether to reverse the results."
  * `formats` is an array of formats: use `"double"`, `"string"` or a date format.

#summary Black-Scholes European put/call valuation.

`BlackScholes`
--------------

Returns the Black-Scholes European call/put valuation.

```
=BlackScholes(call_put_flag, stock_price, strike_price, time_to_expiry, risk-free rate, dividend_yield, volatility)
=BlackScholes("c", 60, 65, 0.25, 0.08, 0, 0.3) // returns 2.1334
```

  * `call_put_flag` is whether the instrument is a call (`"c"`) or a put (`"p"`).
  * `stock_price` is the current value of the underlying stock.
  * `strike_price` is the option's strike price.
  * `time_to_expiry` is the time to maturity in years.
  * `risk-free rate` is the risk-free rate through expiry.
  * `dividend_yield` is the annual dividend yield.
  * `volatility` is the implied volatility at expiry.

`GBlackScholes`
---------------

Returns the Black-Scholes European call/put valuation.

```
=GBlackScholes(call_put_flag, stock_price, strike_price, time_to_expiry, risk-free rate, cost_of_carry, volatility)
```

  * `call_put_flag` is whether the instrument is a call (`"c"`) or a put (`"p"`).
  * `stock_price` is the current value of the underlying stock.
  * `strike_price` is the option's strike price.
  * `time_to_expiry` is the time to maturity in years.
  * `risk-free rate` is the risk-free rate through expiry.
  * `cost_of_carry` is the annualized cost of carry.
  * `volatility` is the implied volatility at expiry.

`ImpliedVolatility`
-------------------

Returns the Black-Scholes implied volatility using the Newton-Raphson method.

```
=ImpliedVolatility(call_put_flag, stock_price, strike_price, time_to_expiry, risk-free rate, dividend_yield, price)
=ImpliedVolatility("c", 60, 65, 0.25, 0.08, 0, 2.1334) // returns 0.3
```

  * `call_put_flag` is whether the instrument is a call (`"c"`) or a put (`"p"`).
  * `stock_price` is the current value of the underlying stock.
  * `strike_price` is the option's strike price.
  * `time_to_expiry` is the time to maturity in years.
  * `risk-free rate` is the risk-free rate through expiry.
  * `dividend_yield` is the annual dividend yield.
  * `price` is the Black-Scholes European put/call valuation.

`Black76`
---------

Returns the Black-76 valuation for options on futures and forwards.

```
=Black76(call_put_flag, forward, strike_price, time_to_expiry, risk-free_rate, volatility)
=Black76("c", 100, 98, 1, 0.05, 0.1) // returns 4.7829
```

  * `call_put_flag` is whether the instrument is a call (`"c"`) or a put (`"p"`).
  * `forward` is the current forward value.
  * `strike_price` is the option's strike price.
  * `time_to_expiry` is the time to maturity in years.
  * `risk-free rate` is the risk-free rate through expiry.
  * `volatility` is the implied volatility at expiry.


`Swaption`
----------

Returns the Black-76 European payer/receiver swaption valuation.

```
=Swaption(pay_rec_flag, tenor, periods, swap_rate, strike_rate, time_to_expiry, risk-free_rate, volatility)
```

  * `pay_rec_flag` is whether the instrument is a payer (`"p"`) or a receiver (`"r"`).
  * `tenor` is the tenor of the swap in years.
  * `periods` is the number of compoundings per year.
  * `swap_rate` is the current underlying swap rate.
  * `strike_rate` is the option's strike rate.
  * `time_to_expiry` is the time to maturity in years.
  * `risk-free_rate` is the risk-free rate through expiry.
  * `volatility` is the implied volatility at expiry.

#summary One-sentence summary of this page.

Greeks
------

Returns the options Greek for a particular sensitivity. _(Note: All functions for the Greeks share a common set of arguments, regardless of whether those inputs are used in a particular Greek's calculation.)_

```
=BSDelta(call_put_flag, stock_price, strike_price, time_to_expiry, risk_free_rate, dividend_yield, volatility)
=BSDelta("c", 60, 65, 0.25, 8%, 0%, 30%) // returns 0.37
```

  * `call_put_flag` is whether the instrument is a call (`"c"`) or a put (`"p"`).
  * `stock_price` is the current value of the underlying stock.
  * `strike_price` is the option's strike price.
  * `time_to_expiry` is the time to maturity in years.
  * `risk-free_rate` is the risk-free rate through expiry.
  * `dividend_yield` is the annual dividend yield.
  * `volatility` is the implied volatility at expiry.

<table>
  <thead>
    <tr>
      <th>*Function*</th>
      <th>*Sensitivity of `__`*</th>
      <th>*to changes in `__`*</td>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>=BSDelta()</td>
      <td>option price</td>
      <td>underlying price</td>
    </tr>
    <tr>
      <td>=Vega()</td>
      <td>option price</td>
      <td>volatility</td>
    </tr>
    <tr>
      <td>=Theta()</td>
      <td>option price</td>
      <td>passage of time</td>
    </tr>
    <tr>
      <td>=Rho()</td>
      <td>option price</td>
      <td>risk-free rate</td>
    </tr>
    <tr>
      <td>=Gamma()</td>
      <td>option price</td>
      <td>delta</td>
    </tr>
    <tr>
      <td>=Vanna()</td>
      <td>delta</td>
      <td>volatility</td>
    </tr>
    <tr>
      <td>=Charm()</td>
      <td>delta</td>
      <td>passage of time</td>
    </tr>
    <tr>
      <td>=Speed()</td>
      <td>gamma</td>
      <td>underlying price</td>
    </tr>
    <tr>
      <td>=Zomma()</td>
      <td>gamma</td>
      <td>volatility</td>
    </tr>
    <tr>
      <td>=Color()</td>
      <td>gamma</td>
      <td>passage of time</td>
    </tr>
    <tr>
      <td>=DvegaDtime()</td>
      <td>vega</td>
      <td>passage of time</td>
    </tr>
    <tr>
      <td>=Vomma()</td>
      <td>vega</td>
      <td>volatility</td>
    </tr>
    <tr>
      <td>=DualDelta()</td>
      <td>option price</td>
      <td>strike price</td>
    </tr>
    <tr>
      <td>=DualGamma()</td>
      <td>delta</td>
      <td>strike price</td>
    </tr>
  </tbody>
</table>

#summary Options valuation for American and Bermudan options.

`American`
----------

Returns the Barone-Adesi-Whaley approximation for an American option.

```
=American(call_put_flag, stock_price, strike_price, time_to_expiry, risk-free rate, dividend_yield, volatility)
```

  * `call_put_flag` is whether the instrument is a call (`"c"`) or a put (`"p"`).
  * `stock_price` is the current value of the underlying stock.
  * `strike_price` is the option's strike price.
  * `time_to_expiry` is the time to maturity in years.
  * `risk-free rate` is the risk-free rate through expiry.
  * `dividend_yield` is the annual dividend yield.
  * `volatility` is the implied volatility at expiry.

`BermudanBinomial`
------------------

Returns the binomial valuation for a Bermudan option.

```
=BermudanBinomial(call_put_flag, stock_price, strike_price, time_to_expiry, risk-free rate,
                  dividend_yield, volatility, potential_exercise_times, iterations)
```

  * `call_put_flag` is whether the instrument is a call (`"c"`) or a put (`"p"`).
  * `stock_price` is the current value of the underlying stock.
  * `strike_price` is the option's strike price.
  * `time_to_expiry` is the time to maturity in years.
  * `risk-free rate` is the risk-free rate through expiry.
  * `dividend_yield` is the annual dividend yield.
  * `volatility` is the implied volatility at expiry.
  * `potential_exercise_times` is a range of potential exercise times in years.
  * `iterations` is the number of calculations performed to increase precision. Defaults to `500`.

  #summary Forward rate agreements valuation

`FRA`
-----

Returns the theoretical forward rate between tenor A and tenor B.

```
=FRA(rate_a, ttm_a, rate_a, rate_b, basis)
```

  * `rate_a` is the rate through point A.
  * `ttm_a` is time in days to point A.
  * `rate_b` is the rate through point B.
  * `ttm_b` is time in days to point B.
  * `basis` is the basis in days (`360`, `365`, etc.).

`FRAFromFXLong`
---------------

For short-term contracts with a maturity of less than one year from now, returns the theoretical long forward rate given a currency forward and FX rates.

```
=FRAFromFXLong(fx_spot, fx_swap_long, fx_swap_short, foreign_fra,
               start_days, end_days, domestic_basis, foreign_basis)
```

  * `fx_spot` is the foreign exchange spot rate.
  * `fx_swap_long` is the long foreign exchange swap.
  * `fx_swap_short` is the short foreign exchange swap.
  * `foreign_fra` is the foreign forward rate agreement.
  * `start_days` is the time in days to the start of the FRA.
  * `end_days` is the time in days to the end of the FRA.
  * `domestic_basis` is the domestic basis in days (`360`, `365`, etc.).
  * `foreign_basis` is the foreign basis in days (`360`, `365`, etc.).

`FRAFromFXShort`
----------------

For short-term contracts with a maturity of less than one year from now, returns the theoretical short forward rate given a currency forward and FX rates.

```
=FRAFromFXShort(fx_spot, fx_swap_long, fx_swap_short, foreign_fra,
                start_days, end_days, domestic_basis, foreign_basis)
```

  * `fx_spot` is the foreign exchange spot rate.
  * `fx_swap_long` is the long foreign exchange swap.
  * `fx_swap_short` is the short foreign exchange swap.
  * `foreign_fra` is the foreign forward rate agreement.
  * `start_days` is the time in days to the start of the FRA.
  * `end_days` is the time in days to the end of the FRA.
  * `domestic_basis` is the domestic basis in days (`360`, `365`, etc.).
  * `foreign_basis` is the foreign basis in days (`360`, `365`, etc.).

ContinuousInterpolation

Returns the interpolated rate given two sets of continuously compounded rates and tenors.

=ContinuousInterpolation(tenors, rates, tenor_interp)
=ContinuousInterpolation({0, 1, 2, 3}, {0.25, 0.5, 1, 2}, 1.5) // returns 0.8333
tenors is a range of tenors in years.
rates is a range of continuously compounded rates.
tenor_interp is the tenor for which you want an interpolated rate.
CubicSpline
Interpolates a rate based on a cubic spline.

=CubicSpline(interp_dates, known_dates, known_rates)
interp_dates is a range of unknown dates for which to find the interpolated rates.
known_dates is a range of known dates or term points.
known_rates is a range of known rates.
Stub
Calculates the spot stub.

=Stub(valuation_date, overnight_start, overnight_value, term_start, term_ends, term_values)
valuation_date is the target valuation date.
overnight_start is the overnight value date (for LIBOR, usually 1 business day after trade date).
overnight_value is the overnight value.
term_start is the term valuation date (for LIBOR, usually 2 business days after trade date).
term_ends is either 1 or 2 term LIBOR expiration dates.
term_values is either 1 or 2 term LIBOR values.

#summary Sorting information in Excel.

Note: Before proceeding, please read the section on [array formulas](ArrayFormulas) if you are unfamiliar with their usage.

`AutoSort`
----------

Automatically sorts an array in Excel when the range it refers to is updated.

```
=AutoSort(range, index, sort_vertical, sort_ascending)
=AutoSort(A1:B3, 2) // returns an array of values based on cells A1:B3, sorted by the second column
```

  * `range` is a list of cells to monitor for changes in value.
  * `index` is the index of the row of column you wish to sort on. Defaults to `1`.
  * `sort_vertical` is whether you wish to sort based on vertical data. Defaults to `TRUE`.
  * `sort_ascending` is whether you wish to sort in ascending or alphabetical order. Defaults to `TRUE`.

_(Note: After filling out the formula as shown below, press `Ctrl+Shift+Enter` instead of `Enter`. This will create an array formula that applies to all selected cells.)_

http://finansu.googlecode.com/hg/img/autosort.gif

#summary Some color-related functions.

These functions are designed to aid in the development of themes, just because I got tired of converting hexadecimal colors using online tools. It supports the conversion of [hexadecimal](http://en.wikipedia.org/wiki/Web_colors), [RGB](http://en.wikipedia.org/wiki/RGB_color_model) and [HSV](http://en.wikipedia.org/wiki/HSL_and_HSV) colors between each other.

`RGBToHex`
----------

Converts an RGB color to hexadecimal format.

```
=RGBToHex(red, green, blue)
=RGBToHex(142, 186, 229) // returns "#8ebae5"
```

  * `red` is the level of red in the color (0‒255).
  * `green` is the level of green in the color (0‒255).
  * `blue` is the level of blue in the color (0‒255).

`RGBToHSV`
----------

Converts an RGB color to HSV format.

```
=RGBToHSV(red, green, blue, flag)
=RGBToHSV(142, 186, 229)      // returns {209.7, 0.38, 0.90}
=RGBToHSV(142, 186, 229, "h") // returns 209.7
```

  * `red` is the level of red in the color (0‒255).
  * `green` is the level of green in the color (0‒255).
  * `blue` is the level of blue in the color (0‒255).
  * `flag` is text describing which HSV value you are requesting. Defaults to a horizontal array containing all three.

`HexToRGB`
----------

Converts a hexadecimal color to RGB format.

```
=HexToRGB(hex, flag)
=HexToRGB("#8ebae5")      // returns {142, 186, 229}
=HexToRGB("#8ebae5", "r") // returns 142
```

  * `hex` is the color in three- or six-digit hexadecimal format.
  * `flag` is text describing which RGB value you are requesting. Defaults to a horizontal array containing all three.

`HexToHSV`
----------

Converts a hexadecimal color to HSV format.

```
=HexToHSV(hex, flag)
=HexToHSV("#8ebae5")        // returns {209.7, 0.38, 0.90}
=HexToHSV("#8ebae5", "sat") // returns 0.38
```

  * `hex` is the color in three- or six-digit hexadecimal format.
  * `flag` is text describing which RGB value you are requesting. Defaults to a horizontal array containing all three.

`HSVToRGB`
----------

Converts an HSV color to RGB format.

```
=HSVToRGB(hue, saturation, value, flag)
=HSVToRGB(209.6, 0.379, 0.899)          // returns {142, 186, 229}
=HSVToRGB(209.6, 0.379, 0.899, "green") // returns 186
```

  * `hue` is the level of hue in the color (0‒100%).
  * `saturation` is the level of saturation in the color (0‒100%).
  * `value` is the level of value (or brightness) in the color (0‒360).
  * `flag` is text describing which RGB value you are requesting. Defaults to a horizontal array containing all three.

`HSVToHex`
----------

Converts an HSV color to hexadecimal format.

```
=HSVToRGB(hue, saturation, value)
=HSVToRGB(209.6, 0.379, 0.899) // returns "#8ebae5"
```

  * `hue` is the level of hue in the color (0‒100%).
  * `saturation` is the level of saturation in the color (0‒100%).
  * `value` is the level of value (or brightness) in the color (0‒360).

#summary Distribution and Density Functions

Retrieving Security Quotes
--------------------------

### `BND`

Returns the bivariate normal distribution function.

```
=BND(x, y, rho)
```

### `CND`

Returns the standard normal cumulative distribution (has a mean of zero and a standard deviation of one).

```
=CND(z)
=CND(0.5) // returns 0.6915
```

  * `z` is the value for which you want the distribution.

### `CNDEV`

Returns the inverse cumulative normal distribution function.

```
=CNDEV(U)
=CNDEV(0.5) // returns 0
```

  * `U` is the value for which you want the distribution.

### `ND`

Returns the normal distribution function.

```
=ND(x)
=ND(0.5) // returns 0.3521
```

  * `x` is the value for which you want the distribution.

### `PDF`

Returns the probability density function.

```
=PDF(z)
=PDF(0.5) // returns 0.3521
```

* `z` is the value for which you want the distribution.

#summary Functions for retrieving version information.

`LatestVersion`
---------------

Retrieves the version number for the latest version of !FinAnSu available on the [project home page](http://code.google.com/p/finansu/).

```
=LatestVersion() // returns a version number such as "1.0.1"
```

`CurrentVersion`
----------------

Returns the currently installed version number for !FinAnSu.

```
=CurrentVersion() // returns a version number such as "1.0.1"
```

`UpdateAvailable`
-----------------

Returns whether an update exists for !FinAnSu.

```
=UpdateAvailable() // returns TRUE or FALSE
```

#summary Currency Formats

Currency Formats
----------------

The `Currency` drop-down in the `Formatting` group of the Excel Ribbon will format cells in accounting format with the appropriate currency symbol and decimal places.

<table>
  <thead>
    <tr>
      <th>Label</th>
      <th>Symbol</th>
      <th>Decimals</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>Whole number</td>
      <td></td>
      <td>0</td>
    </tr>
    <tr>
      <td>Two decimal places</td>
      <td></td>
      <td>2</td>
    </tr>
    <tr>
      <td>Four decimal places</td>
      <td></td>
      <td>4</td>
    </tr>
    <tr>
      <td>Dollars</td>
      <td>$</td>
      <td>0</td>
    </tr>
    <tr>
      <td>Dollars and cents</td>
      <td>$</td>
      <td>2</td>
    </tr>
    <tr>
      <td>Euros</td>
      <td>€</td>
      <td>0</td>
    </tr>
    <tr>
      <td>Euros and cents</td>
      <td>€</td>
      <td>2</td>
    </tr>
    <tr>
      <td>Yen</td>
      <td>¥</td>
      <td>0</td>
    </tr>
    <tr>
      <td>Yen and sen</td>
      <td>¥</td>
      <td>2</td>
    </tr>
    <tr>
      <td>Pounds</td>
      <td>£</td>
      <td>0</td>
    </tr>
    <tr>
      <td>Pounds and pence</td>
      <td>£</td>
      <td>2</td>
    </tr>
  </tbody>
</table>

#summary Layout formats.

`Accounting Underline`
----------------------

Formats a cell with a single accounting underline.

`Center Across Selection`
-------------------------

Centers a cell across the selected range.

#summary Credits for FinAnSu.

ExcelDNA
--------

This project depends on [Excel-Dna](http://exceldna.codeplex.com/), a very nifty tool which lets you utilize modern programming languages in Excel without the headache of [VSTO](http://en.wikipedia.org/wiki/Visual_Studio_Tools_for_Office).

#summary FinAnSu's license information.

MIT License
-----------

A [copy of this license](http://code.google.com/p/finansu/source/browse/license.txt) is included in the source code.

  Copyright (C) 2011 by Bryan !McKelvey

  Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

  The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
