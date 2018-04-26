#' WARNING:
#' Function WordR::addFlextables() will be deprecated.
#' Pls change your scripts which creates the input tables to using flextable package (and not ReporteRs)
#' and use WordR::body_add_flextables() instead.
#'
#' Read Word document with bookmarks and create other Word document with rendered tables in place.
#'
#' This function is basically a loop wrapper around \code{\link[ReporteRs]{addFlexTable}} function.
#'
#' @param docxIn String of length one; path to Word file with bookmarks.
#' @param docxOut String of length one; path for output Word file.
#' @param FlexTables Named list of FlexTables; Tables to be inserted into the Word file
#' @param debug Boolean of length one; If \code{True} then \code{\link[base]{browser}()} is called at the beginning of the function
#' @param ... Parameters to be sent to other methods (mainly \code{\link[ReporteRs]{addFlexTable}})
#' @return  Path to the rendered Word file if the operation was successfull.
#' @examples
#' #library(ReporteRs)
#' #ft_mtcars <- vanilla.table(mtcars)
#' #ft_iris <- vanilla.table(iris)
#' #FT <- list(ft_mtcars=ft_mtcars,ft_iris=ft_iris)
#' #addFlexTables(
#' #   paste(examplePath(),'templates/templateFT.docx',sep = ''),
#' #   paste(tempdir(),'/resultFT.docx',sep = ''),
#' #   FT)
#'
addFlexTables <- function(docxIn, docxOut, FlexTables = list(), debug = F, ...) {
    if (debug) {
        browser()
    }
    warning("Function WordR::addFlextables() will be deprecated.
            \n Pls change your scripts which creates the input tables to using flextable package (and not ReporteRs)
            \n and use WordR::body_add_flextables() instead.")
    doc <- ReporteRs::docx(template = docxIn)

    bks <- gsub("^t_", "", grep("^t_", ReporteRs::list_bookmarks(doc), value = T))
    for (bk in bks) {
        if (!bk %in% names(FlexTables)) {
            stop(paste("Table rendering: Table", bk, "not in the FlexTables list"))
        }
        doc <- ReporteRs::addFlexTable(doc, FlexTables[[bk]], bookmark = paste0("t_", bk), par.properties = ReporteRs::parProperties(text.align = "center"), ...)
    }

    ReporteRs::writeDoc(doc, docxOut)
    return(docxOut)
}
