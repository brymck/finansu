---
layout: default
title: Web Import
description: FinAnSu's web import functions for security quotes and other data
---

Real-Time Security Quotes
=========================

_Note: Before proceeding, please read the section on [array
formulas](faq#array_formulas) if you are unfamiliar with their usage._

<a name="quote">Quote</a>
-------------------------

Returns the current quote for a security ID from Bloomberg, Google or Yahoo!

{% highlight vbnet %}
=Quote(security_id, source, params, live_updating, frequency, show_headers)

' Returns the current price of WFC from Bloomberg, updated every 15 seconds
' All functions are effectively identical
=LiveQuote("WFC")
=Quote("WFC", , , true)
=Quote("WFC", "b", , true, 15)
=BloombergQuote("WFC", , true, 15)
=LiveQuote("WFC", "b", , 15)

' Returns the current price, change and % change for WFC, including headers (static)
=Quote("WFC", "b", "px%", , , true)
=Quote("WFC", "b", QuoteParams(true, true, true), , , true)
{% endhighlight %}

  * `security_id` is the security ID from the quote service.
  * `source` is the name or abbreviation of the quote service (`"b"`, `"Bloomberg"`, `"g"`, `"Google"`, `"y"`, `"Yahoo"`, etc.). Defaults to `"Bloomberg"`.
  * `params` is a list of which values to return. Accepts any combination of `"px%dtbahlv"` (price, change, % change, date, time, bid, ask, high, low, volume). Bid/ask are not available through Google. Use [=QuoteParams()](#quote_params) for help if necessary. Defaults to price.
  * `live_updating` is whether you want this function to return continuously stream live quotes to the cell. Defaults to `false`.
  * `frequency` is the number of seconds between update requests (if live_updating is `true`). Defaults to `15` seconds.
  * `show_headers` is whether to display the headers for each column. Defaults to `false`.

http://www.brymck.com/images/finansu_live_quote.gif

<a name="quote_params">QuoteParams</a>
--------------------------------------

Builds a text string for use in [=Quote()](#quote) designating which values you would like returned.

{% highlight vbnet %}
=QuoteParams(price, change, pct_change, date, time, bid, ask, open, high, low, volume)

' Returns the appropriate text string for use in =Quote(), such that values for
' price, change and % change will be returned
=QuoteParams(true, true, true)
{% endhighlight %}

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

LiveQuote
---------

Same as [Quote](#quote) with the `live_updating` argument equal to `true`.

BloombergQuote
--------------

Same as [Quote](#quote) with the `source` argument equal to `"Bloomberg"`.

GoogleQuote
-----------

Same as [Quote](#quote) with the `source` argument equal to `"Google"`.

YahooQuote
----------

Same as [Quote](#quote) with the `source` argument equal to `"Yahoo"`.

FullTicker
----------

Returns FinAnSu's interpretation of an abbreviated security ID. Mostly for debugging purposes.

{% highlight vbnet %}
=FullTicker(security_id, source, force_interpret)

' Returns "WFC"
=FullTicker("WFC", "b", false)
{% endhighlight %}

  * `security_id` is the security ID from the quote service.
  * `source` is the name or abbreviation of the quote service (`"b"`, `"Bloomberg"`, `"g"`, `"Google"`, `"y"`, `"Yahoo"`, etc.). Defaults to `"Bloomberg"`.
  * `force_interpret` forces FinAnSu to guess at a suffix if none exists if `force_interpret` is set to `true` (may result in errors).

ShortenSource
-------------

Returns FinAnSu's intepretation of an abbreviated source name. Mostly for debugging purposes.

{% highlight vbnet %}
=ShortenSource(source)

' Returns "b"
=ShortenSource("bloomberg")
{% endhighlight %}

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

Quote History
=============

<a name="quote_history">QuoteHistory</a>
----------------------------------------

Returns the historical date, open, high, low, close, volume and adjusted price
for a security ID from the selected quotes provider.

{% highlight vbnet %}
=QuoteHistory(security_id, source, start_date, end_date, period, names, show_headers, date_order)

' Returns date, OHLC, volume and adjusted close for WFC from Yahoo! Finance for
' the past year
=QuoteHistory("WFC")

' Returns monthly date and adjusted close for WFC from Yahoo! Finance for the
' past year
=QuoteHistory("WFC", "y" , , , "m", "da", false, false)
=YahooHistory("WFC", , , "m", "da", false, false)

' Returns the date, close and volume for WFC from Google Finance for each day in
' 2010, in chronological order, with headers
=QuoteHistory("WFC", "g", DATE(2010, 1, 1), DATE(2010, 12, 31), "d", "dcv", true, true)
=GoogleHistory("WFC", DATE(2010, 1, 1), DATE(2010, 12, 31), "d", "dcv", true, true)
{% endhighlight %}

  * `security_id` is the security ID from the quote service.
  * `source` is the name or abbreviation of the quote service (g, Google, y,
    Yahoo, etc.). Defaults to Yahoo!.
  * `start_date` is the date from which to start retrieving history. Defaults to
    the most recent close.
  * `end_date` is the date at which to stop retrieving history. Defaults to one
    year ago.
  * `period` is a text flag representing whether you want daily (`"d"`), weekly
    (`"w"`), monthly (`"m"`) or yearly (`"y"`) quotes. Monthly and yearly quotes are
    available only through Yahoo!. Defaults to daily. Defaults to `"d"`.
  * `names` is a list of which values to return. Accepts any combination of
    `"dohlcva"` (date, open, high, low, close, volume, adj
    price). Adj price is available only through Yahoo!. Use
    [=QuoteHistoryParams()](#quote_history_params) for help if necessary. Defaults
    to all.
  * `show_headers` is whether to display the headers for each column. Defaults
    to `false`.
  * `date_order` is whether to sort dates in ascending chronological order.
    Defaults to `false`.

<a name="quote_history_params">QuoteHistoryParams</a>
-----------------------------------------------------

Builds a text string for use in [=QuoteHistory()](#quote_history) designating
which values you would like returned.

{% highlight vbnet %}
=QuoteHistoryParams(date, open, high, low, close, volume, adj_close)

' Returns the appropriate text string for use in =QuoteHistory(), such that
' values for date, close and volume will be returned
=QuoteHistoryParams(true, , , , true, true)
{% endhighlight %}

  * `date` is whether you wish to return the date.
  * `open` is whether you wish to return the day's opening price.
  * `high` is whether you wish to return the day's high price.
  * `low` is whether you wish to return the day's lowest price.
  * `close` is whether you wish to return the day's closing price.
  * `volume` is whether you wish to return the day's volume.
  * `adj_close` is whether you wish to return the day's closing price.


<a name="yahoo_history">YahooHistory</a>
----------------------------------------

Same as [QuoteHistory](#quote_history) with the `source` argument equal to
`"yahoo"`.

<a name="google_history">GoogleHistory</a>
------------------------------------------

Same as [QuoteHistory](#quote_history) with the `source` argument equal to
`"google"`. Note that Google does not contain easily accessible data for things
like indexes (such as the S&P 500). If you require such information, I recommend
using [yahoo_history](#yahoo_history) instead.

<a name="h15_history">H15History</a>
------------------------------------

Returns information from the Fed's
[http://www.federalreserve.gov/releases/h15/update/ H.15 Statistical Release].

{% highlight vbnet %}
=H15History(instrument_id, frequency)

' Returns a list of dates and yields for Aaa corporate bonds
=H15History("AAA_NA", "m")
{% endhighlight %}

  * `instrument_id` is the instrument ID. Go to [and click
    on one of the data links. The ID is in the URL in the form
    `H15_[id](http://www.federalreserve.gov/releases/h15/data.htm]).txt`.
  * `frequency` is business day (`"b"`), daily (`"d"`), weekly Wednesday
    (`"ww"`), weekly Thursday (`"wt"`), weekly Friday (`"wf"`), bi-weekly
    (`"bw"`), monthly (`"m"`) or annual (`"a"`). Not all frequencies are
    available for all instruments. Defaults to business day (`"b"`)."

Generic Web Importing
=====================

FinAnSu is not limited to security quotes or even finance. It allows
you to import any data you consider relevant by simply exposing the
function it uses to pull security quotes to the user. Although it's
not as user-friendly, if you know the URL and how to use [regular
expressions](http://www.regular-expressions.info/reference.html)*, you can use
it for anything.

_\* Yes, I'm aware that regex is an ugly tool for the job. However, it's the
easiest to maintain and provides some performance benefits over the HTML/XML
parsers I considered._

<a name="import">Import</a>
---------------------------

Returns a horizontal array of values based on a URL and regular expression.

{% highlight vbnet %}
=Import(url, pattern, max_length, live_updating, frequency)

' Returns the title of the top story from bloomberg.com
=Import("http://www.bloomberg.com/", "story_link[^>]+>(.*?)<", 1, false) 
{% endhighlight %}

  * `url` is the full URL of the target website.
  * `pattern` is a [regular expression
    pattern](http://www.regular-expressions.info/quickstart.html) where the
    first backreference (in parentheses) is the value you wish to retrieve.
  * `max_length` is the maximum length of the results array.
  * `live_updating` is whether you want this function to return continuously
    stream live quotes to the cell.
  * `frequency` is the number of seconds between update requests (if
    live_updating is `true`). Defaults to `15` seconds.

GetWebData
----------

Same as [Import](#import) with the `live_updating` argument equal to `false`.

<a name="import_csv">ImportCSV</a>
----------------------------------

Returns an array of values from a CSV.

{% highlight vbnet %}
=ImportCSV(url, start_line, reverse, formats)

' Returns a list of market sectors and security types
=ImportCSV("http://bsym.bloomberg.com/sym/pages/security_type.csv", 1, false, {"string", "string"})
{% endhighlight %}

  * `url` is the URL of the target CSV file.
  * `start_line` is the first line of the CSV to begin parsing (starting with
    `0`).
  * `reverse` is whether to reverse the results."
  * `formats` is an array of formats: use `"double"`, `"string"` or a date
    format.
