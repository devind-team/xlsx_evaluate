Excel number precision
----------------------

Excel number precision is a complex discussion.

It has been discussed in a [Wikipedia page](https://en.wikipedia.org/wiki/Numeric_precision_in_Microsoft_Excel>)

The fundamentals come down to floating point numbers and a contention between
how they are represented in memory Vs how they are stored on disk Vs how they
are presented on screen. A [Microsoft article](https://www.microsoft.com/en-us/microsoft-365/blog/2008/04/10/understanding-floating-point-precision-aka-why-does-excel-give-me-seemingly-wrong-answers/)
explains the contention.

This project is attempting to take care while reading numbers from the Excel
file to try and remove a variety of representation errors.

Further work will be required to keep numbers in-line with Excel throughout
different transformations.

From what I can determine this requires a low-level implementation of a
numeric datatype (C or C++, Cython??) to replicate its behaviour. Python
built-in numeric types don't replicate behaviours appropriately.


Unit testing Excel formulas directly from the workbook.
-------------------------------------------------------

If you are interested in unit testing formulas in your workbook, you can use
[FlyingKoala](https://github.com/bradbase/flyingkoala). An example on how can
be found
[here](https://github.com/bradbase/flyingkoala/tree/master/flyingkoala/unit_testing_formulas).