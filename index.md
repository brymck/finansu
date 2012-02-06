---
layout: default
title: FinAnSu
description: An introduction to FinAnSu
---

What Does It Do?
================

FinAnSu aims to provide user-friendly tools for use in financial applications.
The add-in is in its development stages, but it currently offers:

  * *Live, streaming [web import](WebData) capabilities* (Excel 2002+),
    including custom functions for easily [importing security prices](Quotes)
    from Bloomberg.com, Google Finance and Yahoo! Finance
  * Functions that can [parse CSVs](WebData#ImportCSV) on the web,
    including custom functions to easily import stock quote data from [Google
    Finance](Quotes#GoogleHistory) and instrument data from the [Fed H.15
    release](Quotes#H15History) (yields for Fed Funds, commercial paper,
    Treasuries, interest-rate swaps, etc.).
  * Basic options pricing, including [Black-Scholes](BlackScholes), [the options
    Greeks](Greeks) and [a few more complex options models](AmericanBermudan)
  * A bit on [FRAs] (forward rate agreements)
  * A continuously compounded rate [interpolator](Interpolation) (need to add
    more: linear, cubic, cubic spline, etc.)
  * A Federal Reserve holiday calculator (need to add more for different
    calendars, as well as roll date conventions)
  * A handful of tools for converting between discount factors and forward rates
  * Function to [automatically sort](Sorting) ranges that contain rows and
    columns of data
  * Some formatting macros for [currencies](Currencies) and [a few layout
    options](Layout) that aren't easily accessible in Excel

![FinAnSu in action](img/quote.gif)

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

Accounting Underline
--------------------

Formats a cell with a single accounting underline.

Center Across Selection
-----------------------

Centers a cell across the selected range.
