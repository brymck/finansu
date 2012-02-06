---
layout: default
title: Yield Curves
description: FinAnSu yield curve interpolation and manipulation functions
---

Yield Curves
============

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
