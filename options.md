---
layout: default
title: Options
description: summary Black-Scholes European put/call valuation.
---

BlackScholes
------------

Returns the Black-Scholes European call/put valuation.

{% highlight vbnet %}
=BlackScholes(call_put_flag, stock_price, strike_price, time_to_expiry, risk-free rate, dividend_yield, volatility)

' Returns 2.1334
=BlackScholes("c", 60, 65, 0.25, 0.08, 0, 0.3)
{% endhighlight %}

  * `call_put_flag` is whether the instrument is a call (`"c"`) or a put (`"p"`).
  * `stock_price` is the current value of the underlying stock.
  * `strike_price` is the option's strike price.
  * `time_to_expiry` is the time to maturity in years.
  * `risk-free rate` is the risk-free rate through expiry.
  * `dividend_yield` is the annual dividend yield.
  * `volatility` is the implied volatility at expiry.

GBlackScholes
-------------

Returns the Black-Scholes European call/put valuation.

{% highlight vbnet %}
=GBlackScholes(call_put_flag, stock_price, strike_price, time_to_expiry, risk-free rate, cost_of_carry, volatility)
{% endhighlight %}

  * `call_put_flag` is whether the instrument is a call (`"c"`) or a put (`"p"`).
  * `stock_price` is the current value of the underlying stock.
  * `strike_price` is the option's strike price.
  * `time_to_expiry` is the time to maturity in years.
  * `risk-free rate` is the risk-free rate through expiry.
  * `cost_of_carry` is the annualized cost of carry.
  * `volatility` is the implied volatility at expiry.

ImpliedVolatility
-----------------

Returns the Black-Scholes implied volatility using the Newton-Raphson method.

{% highlight vbnet %}
=ImpliedVolatility(call_put_flag, stock_price, strike_price, time_to_expiry, risk-free rate, dividend_yield, price)

' Returns 0.3
=ImpliedVolatility("c", 60, 65, 0.25, 0.08, 0, 2.1334)
{% endhighlight %}

  * `call_put_flag` is whether the instrument is a call (`"c"`) or a put (`"p"`).
  * `stock_price` is the current value of the underlying stock.
  * `strike_price` is the option's strike price.
  * `time_to_expiry` is the time to maturity in years.
  * `risk-free rate` is the risk-free rate through expiry.
  * `dividend_yield` is the annual dividend yield.
  * `price` is the Black-Scholes European put/call valuation.

Black76
-------

Returns the Black-76 valuation for options on futures and forwards.

{% highlight vbnet %}
=Black76(call_put_flag, forward, strike_price, time_to_expiry, risk-free_rate, volatility)

' Returns 4.7829
=Black76("c", 100, 98, 1, 0.05, 0.1)
{% endhighlight %}

  * `call_put_flag` is whether the instrument is a call (`"c"`) or a put (`"p"`).
  * `forward` is the current forward value.
  * `strike_price` is the option's strike price.
  * `time_to_expiry` is the time to maturity in years.
  * `risk-free rate` is the risk-free rate through expiry.
  * `volatility` is the implied volatility at expiry.


Swaption
--------

Returns the Black-76 European payer/receiver swaption valuation.

{% highlight vbnet %}
=Swaption(pay_rec_flag, tenor, periods, swap_rate, strike_rate, time_to_expiry, risk-free_rate, volatility)
{% endhighlight %}

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

{% highlight vbnet %}
=BSDelta(call_put_flag, stock_price, strike_price, time_to_expiry, risk_free_rate, dividend_yield, volatility)

' Returns 0.37
=BSDelta("c", 60, 65, 0.25, 8%, 0%, 30%)
{% endhighlight %}

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
      <th>Function</th>
      <th>Sensitivity of __</th>
      <th>to changes in __</td>
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

American
--------

Returns the Barone-Adesi-Whaley approximation for an American option.

{% highlight vbnet %}
=American(call_put_flag, stock_price, strike_price, time_to_expiry, risk-free rate, dividend_yield, volatility)
{% endhighlight %}

  * `call_put_flag` is whether the instrument is a call (`"c"`) or a put (`"p"`).
  * `stock_price` is the current value of the underlying stock.
  * `strike_price` is the option's strike price.
  * `time_to_expiry` is the time to maturity in years.
  * `risk-free rate` is the risk-free rate through expiry.
  * `dividend_yield` is the annual dividend yield.
  * `volatility` is the implied volatility at expiry.

BermudanBinomial
----------------

Returns the binomial valuation for a Bermudan option.

{% highlight vbnet %}
=BermudanBinomial(call_put_flag, stock_price, strike_price, time_to_expiry, risk-free rate,
                  dividend_yield, volatility, potential_exercise_times, iterations)
{% endhighlight %}

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

FRA
---

Returns the theoretical forward rate between tenor A and tenor B.

{% highlight vbnet %}
=FRA(rate_a, ttm_a, rate_a, rate_b, basis)
{% endhighlight %}

  * `rate_a` is the rate through point A.
  * `ttm_a` is time in days to point A.
  * `rate_b` is the rate through point B.
  * `ttm_b` is time in days to point B.
  * `basis` is the basis in days (`360`, `365`, etc.).

FRAFromFXLong
-------------

For short-term contracts with a maturity of less than one year from now, returns the theoretical long forward rate given a currency forward and FX rates.

{% highlight vbnet %}
=FRAFromFXLong(fx_spot, fx_swap_long, fx_swap_short, foreign_fra,
               start_days, end_days, domestic_basis, foreign_basis)
{% endhighlight %}

  * `fx_spot` is the foreign exchange spot rate.
  * `fx_swap_long` is the long foreign exchange swap.
  * `fx_swap_short` is the short foreign exchange swap.
  * `foreign_fra` is the foreign forward rate agreement.
  * `start_days` is the time in days to the start of the FRA.
  * `end_days` is the time in days to the end of the FRA.
  * `domestic_basis` is the domestic basis in days (`360`, `365`, etc.).
  * `foreign_basis` is the foreign basis in days (`360`, `365`, etc.).

FRAFromFXShort
--------------

For short-term contracts with a maturity of less than one year from now, returns the theoretical short forward rate given a currency forward and FX rates.

{% highlight vbnet %}
=FRAFromFXShort(fx_spot, fx_swap_long, fx_swap_short, foreign_fra,
                start_days, end_days, domestic_basis, foreign_basis)
{% endhighlight %}

  * `fx_spot` is the foreign exchange spot rate.
  * `fx_swap_long` is the long foreign exchange swap.
  * `fx_swap_short` is the short foreign exchange swap.
  * `foreign_fra` is the foreign forward rate agreement.
  * `start_days` is the time in days to the start of the FRA.
  * `end_days` is the time in days to the end of the FRA.
  * `domestic_basis` is the domestic basis in days (`360`, `365`, etc.).
  * `foreign_basis` is the foreign basis in days (`360`, `365`, etc.).
