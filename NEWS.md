# WordR 0.3
* Removing dependancy on ReporteRs package (keeping it just in addFlexTables function for backward compatibility)
* changing addPlots function (thanks D.Gohel) to use officer not ReporteRs
* adding new function body_add_flextables() which should replace addFlexTables()
* adding flextable into imports

# WordR 0.2.3

* changing the output path to tempdir() in examples and tests, as the CRAN tests were failing

# WordR 0.2.2

* breaking change: different format of inline code notation in docx files!
* adding WordR help page
* improving/changing vignette
* removing setting Java home dir from tests

Thanks goes to David Gohel, maintainer of `officer` package for his comments.
