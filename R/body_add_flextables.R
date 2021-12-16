#' Read Word document with bookmarks and create other Word document with rendered tables in place.
#'
#' This function is basically a loop wrapper around \code{\link[flextable]{body_add_flextable}} function.
#'
#' @param docxIn String of length one; path to Word file with bookmarks.
#' @param docxOut String of length one; path for output Word file.
#' @param flextables Named list of flextables; Tables to be inserted into the Word file
#' @param debug Boolean of length one; If \code{True} then \code{\link[base]{browser}()} is called at the beginning of the function
#' @param ... Parameters to be sent to other methods (mainly \code{\link[flextable]{body_add_flextable}})
#' @return  Path to the rendered Word file if the operation was successfull.
#' @export
#' @examples
#' library(flextable)
#' ft_mtcars <- flextable(mtcars)
#' ft_iris <- flextable(iris)
#' FT <- list(ft_mtcars=ft_mtcars,ft_iris=ft_iris)
#' body_add_flextables(
#'    paste(examplePath(),'templates/templateFT.docx',sep = ''),
#'    paste(tempdir(),'/resultFT.docx',sep = ''),
#'    FT)


body_add_flextables<-function(docxIn, docxOut, flextables,debug = F,...) {
  if (debug) {
    browser()
  }

  doc<-officer::read_docx(docxIn)
  bks <- gsub("^t_", "", grep("^t_", officer::docx_bookmarks(doc), value = T))
  for (bk in bks) {
    if (!(bk %in% names(flextables))) {
      warning(paste("Table rendering: Table", bk, "not in the flextables list"))
    } else
    {
    doc<- officer::cursor_bookmark(doc,paste0("t_",bk))
    doc<- flextable::body_add_flextable(doc, flextables[[bk]],...)
    }
  }

  print(doc, target=docxOut)
  return(docxOut)
}
