---
layout: default
title: FinAnSu
description: An introduction to FinAnSu
---

What Does It Do?
================

FinAnSu aims to provide user-friendly tools for use in financial applications.
The add-in is in its development stages, but it currently offers:

  * *Live, streaming [web import](web) capabilities* (Excel 2002+), including
    custom functions for easily [importing security prices](web#quote) from
    Bloomberg.com, Google Finance and Yahoo! Finance
  * Functions that can [parse CSVs](web#import_csv) on the web, including
    custom functions to easily import stock quote data from [Google
    Finance](web#google_history) and instrument data from the [Fed H.15
    release](web#h15_history) (yields for Fed Funds, commercial paper, Treasuries,
    interest-rate swaps, etc.).
  * Basic options pricing, including [Black-Scholes](options#black_scholes),
    [the options Greeks](options#greeks) and [a few more complex options
    models](options#complex)
  * A bit on [FRAs](yield_curve#fra) (forward rate agreements)
  * A continuously compounded rate [interpolator](yield_curve#interpolation)
    (need to add more: linear, cubic, cubic spline, etc.)
  * A Federal Reserve holiday calculator (need to add more for different
    calendars, as well as roll date conventions)
  * A handful of tools for converting between discount factors and forward rates
  * Function to [automatically sort](other#sorting) ranges that contain rows and
    columns of data
  * Some formatting macros for [currencies](macro#currencies) and [a few layout
    options](macro#layout) that aren't easily accessible in Excel

![FinAnSu in action](img/quote.gif)

Getting Started
===============

Follow the directions on the [install page](install) to get FinAnSu up and
running.
