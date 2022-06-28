#' Read Word document and apply functions \code{\link[WordR]{body_add_flextables}},\code{\link[WordR]{renderInlineCode}},\code{\link[WordR]{addPlots}}.
#'
#' @param docxIn String of length one; path to Word file with bookmarks OR officer::rdocx object
#' @param docxOut String of length one; path for output Word file or NA
#' @param flextables Named list of flextables; Tables to be inserted into the Word file
#' @param Plots Named list of functions (or plots - for plots which are stored as objects e.g., ggplot) creating plots to be inserted into the Word file
#' @param width width of the plots in output in inches, either of length 1 (and will be recycled), or same length as Plots
#' @param height height of the plots in output in inches, either of length 1 (and will be recycled), or same length as Plots
#' @param debug Boolean of length one; If \code{True} then \code{\link[base]{browser}()} is called at the beginning of the function
#' @return  Path to the rendered Word file if the operation was successfull OR officer::rdocx object if docxOut is NA
#'
#' @import officer
#' @import dplyr
#'
#' @export
#' @examples
#' library(flextable)
#' library(ggplot2)
#' ft_mtcars <- flextable(mtcars)
#' FT <- list(mtcars=ft_mtcars)
#' Plots <- list(mtcars1 = ggplot(mtcars, aes(x = wt, y = hp, col = as.factor(cyl))) + geom_point())
#' renderAll(docxIn=paste0(examplePath(), "templates/templateAll.docx"),
#'           docxOut=paste0(examplePath(), "results/templateAll2.docx"),
#'           flextables=FT,
#'           Plots=Plots)
#'
renderAll <- function(docxIn, docxOut, flextables= NULL, Plots= NULL, height =6, width = 6, debug = F) {
  if (debug) {
    browser()
  }

  if("rdocx" %in% class(docxIn)){
    doc<-docxIn
  } else
  {
    doc <- officer::read_docx(path = docxIn)
  }

  doc<-renderInlineCode(doc,NA)
  if(!is.null(flextables)){
    doc<-body_add_flextables(doc, NA, flextables=flextables)
  }
  if(!is.null(Plots)){
    doc<-addPlots(doc, NA, Plots=Plots, height = height, width=width)
  }

  if(is.na(docxOut)){
    return(doc)
  }

  print(doc, target = docxOut)
  return(docxOut)
}
