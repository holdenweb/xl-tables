# xltables

Thank you for your interest in this project.

The `xltables` module attempts to extract tabular information from Excel spreadsheets and make it available
as usable computational data in the form of Pandas DataFrames (which can in turn be transformed into many
other formats).

It has been built initially to transform one particular spreadsheet, historical data on UK doctors' salaries.
This spreasheet is included as a part of the sample data, and was extracted from
[this file](http://www.hscic.gov.uk/catalogue/PUB12625/gpearnextime.xls),
whose description (and the link to the data file) can be found
[here](http://www.hscic.gov.uk/searchcatalogue?productid=13317&q=title%3a%22GP+Earnings+and+Expenses%22&sort=Relevance&size=10&page=1#top)

The first announcement of this repository followed a 25-minute presentation to the
PyData London group on 4th AUgust 2015.
Not mentioned in the presentation was the real point of the project, the module
that can extract tables of the same format from any workbook opened
with `openpyxl`. It has a couple of features that I didn't have time to discuss in
the presentation.

I was interested principally in determining how difficult it would be to extract the data,
since many "open data" policies similarly result in the production of data in computationally
intractable forms. This stands in the way of progress in areas like data-based journalism and in
fact-based decision making.

Since it turned out to involve a significant effort, this made me realize that it might be
useful to others as some sort of starting point and explanation of some of the necessary
techniques as a starting point for their own projects. I have therefore left most of my
exploratory meanderings in the `Invitations.ipynb` Jupyter notebook in hopes that they may
guide others more reliably along a still not yet well-worn path.

## Architecture

At the time of writing xltables is a single module containing a
single function, `extract_tables`, which takes the Pandas representation
of a single worksheet and returns a list of DataFrames, each one
representing a table on that worksheet.

When run as a main program it loads the test workbook, reads in the tables from each sheet and outputs each table as an HTML file.
Of course HTML is not a suitable data transfer medium, so it would be a good idea to add demonstrations of how other forms can
be produced to allow readers to realize the potential of the technology.

There is a lot of work that can be done to improve the range of
table patterns that the routines can recognize.
I shall be happy for readers to post issues to identify the (many) current weaknesses of the system.
This will help set future directions for the project, which is sorely in need of additional effort.
The thing this project needs more than anything else is input from other people, hence its appearance on Github.
Please try the software. If you would like it to do more, or experience problems,
please [post an issue](https://github.com/holdenweb/xl-tables/issues) - look for the
green "New Issue" button - and I'll respond.

