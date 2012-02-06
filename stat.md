---
layout: default
title: Statistics
description: Statistical functions available in FinAnSu
---

Statistical Functions
=====================

BND
---

Returns the bivariate normal distribution function.

{% highlight vbnet %}
=BND(x, y, rho)
{% endhighlight %}

CND
---

Returns the standard normal cumulative distribution (has a mean of zero and a
standard deviation of one).

{% highlight vbnet %}
=CND(z)
=CND(0.5) // returns 0.6915
{% endhighlight %}

  * `z` is the value for which you want the distribution.

CNDEV
-----

Returns the inverse cumulative normal distribution function.

{% highlight vbnet %}
=CNDEV(U)
=CNDEV(0.5) // returns 0
{% endhighlight %}

  * `U` is the value for which you want the distribution.

ND
--

Returns the normal distribution function.

{% highlight vbnet %}
=ND(x)
=ND(0.5) // returns 0.3521
{% endhighlight %}

  * `x` is the value for which you want the distribution.

PDF
---

Returns the probability density function.

{% highlight vbnet %}
=PDF(z)
=PDF(0.5) // returns 0.3521
{% endhighlight %}

* `z` is the value for which you want the distribution.
