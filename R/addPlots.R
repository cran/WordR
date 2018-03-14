#' Read Word document with bookmarks and create other Word documnet with rendered plots in place.
#'
#' This function is basically a loop wrapper around \code{\link[ReporteRs]{addPlot}} function.
#'
#' @param docxIn String of length one; path to Word file with bookmarks.
#' @param docxOut String of length one; path for output Word file.
#' @param Plots Named list of function creating plots (see \code{\link[ReporteRs]{addPlot}}) for details) to be inserted into the Word file
#' @param debug Boolean of length one; If \code{True} then \code{\link[base]{browser}()} is called at the beginning of the function
#' @param ... Parameters to be sent to other methods (mainly \code{\link[ReporteRs]{addPlot}})
#' @return  Path to the rendered Word file if the operation was successfull.
#' @examples
#'library(ggplot2)
#'Plots<-list(plot1=function()plot(hp~wt,data=mtcars,col=cyl),
#'  plot2=function()print(ggplot(mtcars,aes(x=wt,y=hp,col=as.factor(cyl)))+geom_point()))
#'addPlots(
#'  paste(examplePath(),'templates/templatePlots.docx',sep = ''),
#'  paste(tempdir(),'/resultPlots.docx',sep = ''),
#'  Plots,height=4)
#'
addPlots <- function(docxIn, docxOut, Plots = list(), debug = F, ...) {
    if (debug) {
        browser()
    }

    doc <- ReporteRs::docx(template = docxIn)

    bks <- gsub("^p_", "", grep("^p_", ReporteRs::list_bookmarks(doc), value = T))
    for (bk in bks) {
        if (!bk %in% names(Plots)) {
            stop(paste("Plot rendering: Plot", bk, "not in the Plots list"))
        }
        doc <- ReporteRs::addPlot(doc, Plots[[bk]], bookmark = paste0("p_", bk), par.properties = ReporteRs::parProperties(text.align = "center"), ...)
    }

    ReporteRs::writeDoc(doc, docxOut)
    return(docxOut)
}
