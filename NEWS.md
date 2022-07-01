# WordR 0.3.4
* resubmission to CRAN
* fixed problem with writing file in example of renderAll()

# WordR 0.3.3
* adding function renderAll which combines rendering of inline code, plots and flextables
* rendering functions now accept officer::rdocx object as input
* rendering functions now can output officer::rdocx object
* addPlots now accepts plots as objects (where possible e.g. gglots, not base plots); and width and height parameters can be set up individually for each plot.

# WordR 0.3.2
* redoing internals of renderInlineCode - officer::slip_in_text was deprecated, so I moved it here to keep the functionality (consented by David Gohel - thanks!).
* adding Import: xml2

# WordR 0.3.1
* Removing Suggest:ReporteRs
* Removing function addFlexTables (as ReporteRs package is removed from CRAN)

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
