---
layout: default
title: Yield Curves
description: FinAnSu yield curve interpolation and manipulation functions
---

<a name="fra">Forward Rate Agreements</a>
=========================================

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

For short-term contracts with a maturity of less than one year from now, returns
the theoretical long forward rate given a currency forward and FX rates.

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

For short-term contracts with a maturity of less than one year from now, returns
the theoretical short forward rate given a currency forward and FX rates.

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

<a name="interpolation">Yield Curve Interpolation</a>
=====================================================

ContinuousInterpolation
-----------------------

Returns the interpolated rate given two sets of continuously compounded rates
and tenors.

{% highlight vbnet %}
=ContinuousInterpolation(tenors, rates, tenor_interp)
=ContinuousInterpolation({0, 1, 2, 3}, {0.25, 0.5, 1, 2}, 1.5) // returns 0.8333
{% endhighlight %}

  * `tenors` is a range of tenors in years.
  * `rates` is a range of continuously compounded rates.
  * `tenor_interp` is the tenor for which you want an interpolated rate.

CubicSpline
-----------

Interpolates a rate based on a cubic spline.

{% highlight vbnet %}
=CubicSpline(interp_dates, known_dates, known_rates)
{% endhighlight %}

  * `interp_dates` is a range of unknown dates for which to find the
    interpolated rates.
  * `known_dates` is a range of known dates or term points.
  * `known_rates` is a range of known rates.

Stub
----

Calculates the spot stub.

{% highlight vbnet %}
=Stub(valuation_date, overnight_start, overnight_value, term_start, term_ends, term_values)
{% endhighlight %}

  * `valuation_date` is the target valuation date.
  * `overnight_start` is the overnight value date (for LIBOR, usually 1 business
    day after trade date).
  * `overnight_value` is the overnight value.
  * `term_start` is the term valuation date (for LIBOR, usually 2 business days
    after trade date).
  * `term_ends` is either 1 or 2 term LIBOR expiration dates.
  * `term_values` is either 1 or 2 term LIBOR values.
