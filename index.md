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

Getting Started
===============

Follow the directions on the [install page](installation.html) to get FinAnSu
up and running.
