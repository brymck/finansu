FinAnSu
=======

<table>
  <tr>
    <th>Homepage</th>
    <td><a href="http://brymck.herokuapp.com">brymck.herokuapp.com</a></td>
  </tr>
  <tr>
    <th>Documentation</th>
    <td><a href="http://brymck.github.com/finansu/">brymck.github.com/finansu/</a></td>
  </tr>
  <tr>
    <th>Author</th>
    <td>Bryan McKelvey</td>
  </tr>
  <tr>
    <th>Copyright</th>
    <td>(c) 2010-2012 Bryan McKelvey</td>
  </tr>
  <tr>
    <th>License</th>
    <td>MIT</td>
  </tr>
</table>

FinAnSu aims to provide user-friendly tools for use in financial applications.
The add-in is in its development stages, but it currently offers:

  * <b>Live, streaming <a href="http://brymck.github.com/finansu/web">web
    import</a> capabilities</b> (Excel 2002+), including
    custom functions for easily [importing security
    prices](http://brymck.github.com/finansu/web#quote) from Bloomberg.com,
    Google Finance and Yahoo! Finance
  * Functions that can [parse
    CSVs](http://brymck.github.com/finansu/web#import_csv) on the web,
    including custom functions to easily import stock quote data from
    [Google Finance](http://brymck.github.com/finansu/web#google_history)
    and instrument data from the [Fed H.15
    release](http://brymck.github.com/finansu/web#h15_history) (yields for Fed
    Funds, commercial paper, Treasuries, interest-rate swaps, etc.).
  * Basic options pricing, including
    [Black-Scholes](http://brymck.github.com/finansu/options#black_scholes),
    [the options Greeks](http://brymck.github.com/finansu/options#greeks)
    and [a few more complex options
    models](http://brymck.github.com/finansu/options#complex)
  * A bit on [FRAs](http://brymck.github.com/finansu/yield_curve#fra) (forward
    rate agreements)
  * A continuously compounded rate
    [interpolator](http://brymck.github.com/finansu/yield_curve#interpolation)
    (need to add more: linear, cubic, cubic spline, etc.)
  * A Federal Reserve holiday calculator (need to add more for different
    calendars, as well as roll date conventions)
  * A handful of tools for converting between discount factors and forward rates
  * Function to [automatically
    sort](http://brymck.github.com/finansu/other#sorting) ranges that contain rows
    and columns of data
  * Some formatting macros for
    [currencies](http://brymck.github.com/finansu/macro#currencies) and [a few
    layout options](http://brymck.github.com/finansu/macro#layout) that aren't
    easily accessible in Excel

Requirements
------------

Currently, this works with **32- and 64-bit versions of Windows and Excel**
(Office 2003+ for best results).

The add-in also requires [.NET 4](http://www.microsoft.com/net).

Installation
------------

1. Download the current version of FinAnSu on my [GitHub downloads
   page](https://github.com/brymck/finansu/downloads). Most of you are probably
   using a 32-bit version of Office, as it's the default installation even on
   64-bit versions of Windows.
2. Unzip it to a temporary directory
3. Run the `install.bat` script

Development
-----------

This project has the following dependencies (that don't come bundled with the repo):

1. [Ruby for Windows](http://rubyinstaller.org/downloads/)  
   Once it's installed, make sure you have Bundler (`gem install bundler` in a
   command prompt). Then run `bundle update` in the top directory of the repo
   (the same folder as this readme).
2. [7za](http://www.7-zip.org/download.html), the command line version of 7-Zip,
   installed somewhere in your `%Path%`.
3. [Git](http://help.github.com/set-up-git-redirect)
4. And anything like Windows, Office or .NET 4 if you want to do any testing or
   compilation
5. [NUnit](http://www.nunit.org/index.php?p=download) if you plan to run any
   tests. And if you do, make certain that `nunit-console.exe` can be found on
   the system `%Path%`.
6. Optionally, the [File
   utility](http://gnuwin32.sourceforge.net/packages/file.htm) from GnuWin32,   
   again somewhere on your `%Path%`.                                            

In Action
---------

![FinAnSu in action](http://brymck.github.com/finansu/img/quote.gif)
