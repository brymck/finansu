FinAnSu Documentation
=====================

This branch serves as the documentation for FinAnSu.

Dependencies
------------

You must have [Ruby](http://www.ruby-lang.org/en/downloads/) and then Bundler
(`gem install bundler`) installed. Unlike the main add-in, these pages are not
OS-dependent (in fact, I'm doing everything in OSX).

    git clone git://github.com/brymck/finansu.git
    git checkout -b gh-pages remotes/origin/gh-pages
    bundle update
    rake build
