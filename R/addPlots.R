#' Read Word document with bookmarks and create other Word document with rendered plots in place.
#'
#' This function takes a list of functions creating plots and insert into Word document on places with correspondning bookmarks.
#'
#' @param docxIn String of length one; path to Word file with bookmarks OR officer::rdocx object
#' @param docxOut String of length one; path for output Word file or NA
#' @param Plots Named list of functions (or plots - for plots which are stored as objects e.g., ggplot) creating plots to be inserted into the Word file
#' @param width width of the plots in output in inches, either of length 1 (and will be recycled), or same length as Plots
#' @param height height of the plots in output in inches, either of length 1 (and will be recycled), or same length as Plots
#' @param res see res parameter in \code{\link[grDevices]{png}}
#' @param style see style parameter in \code{\link[officer]{body_add_img}}
#' @param debug Boolean of length one; If \code{True} then \code{\link[base]{browser}()} is called at the beginning of the function
#' @param ... Parameters to be sent to other methods (\code{\link[grDevices]{png}})
#' @return  Path to the rendered Word file if the operation was successfull OR officer::rdocx object if docxOut is NA
#'
#' @import officer
#'
#' @export
#' @examples
#'library(ggplot2)
#'Plots<-list(plot1=function()plot(hp~wt,data=mtcars,col=cyl),
#'  plot2=function()print(ggplot(mtcars,aes(x=wt,y=hp,col=as.factor(cyl)))+geom_point()))
#'addPlots(
#'  paste(examplePath(),'templates/templatePlots.docx',sep = ''),
#'  paste(tempdir(),'/resultPlots.docx',sep = ''),
#'  Plots,height=4)
#'
addPlots <- function(docxIn, docxOut, Plots = list(), width = 6, height = 6, res= 300, style= NULL, debug = F, ...) {
  if (debug) {
    browser()
  }
  if("rdocx" %in% class(docxIn)){
    doc<-docxIn
  } else
  {
    doc <- officer::read_docx(path = docxIn)
  }

  if(length(width)==1){
    width<-rep(width,length(Plots))
  }

  if(length(height)==1){
    height<-rep(height,length(Plots))
  }


  bks <- gsub("^p_", "", grep("^p_", officer::docx_bookmarks(doc), value = T))
  for (bk in bks) {
    if (!(bk %in% names(Plots))) {
      warning(paste("Plot rendering: Plot", bk, "not in the Plots list"))
    } else
    {
      ind<-which(bk==names(Plots))
      doc <- officer::cursor_bookmark(doc, paste0("p_", bk))
      pngfile <- tempfile(fileext = ".png")
      grDevices::png(filename = pngfile, width = width[ind], height = height[ind], units = "in", res = res, ...)
      if("function" %in% class(Plots[[bk]])){
        Plots[[bk]]()
      } else
      {
        print(Plots[[bk]])
      }

      grDevices::dev.off()

      doc <- officer::body_add_img(doc, pngfile, style = style, width = width[ind], height = height[ind] ,pos = "on")
    }
  }

  if(is.na(docxOut)){
    return(doc)
  }

  print(doc, target = docxOut)
  return(docxOut)
}
