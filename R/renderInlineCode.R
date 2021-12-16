#' Read Word document with R code blocks, evaluate them and writes the result into another Word document.
#'
#' @param docxIn String of length one; path to Word file with R code blocks.
#' @param docxOut String of length one; path for output Word file.
#' @param debug Boolean of length one; If \code{True} then \code{\link[base]{browser}()} is called at the beginning of the function
#' @return Path to the rendered Word file if the operation was successfull.
#'
#' @import officer
#' @import dplyr
#'
#' @export
#' @examples
#' renderInlineCode(
#'   paste(examplePath(),'templates/template1.docx',sep = ''),
#'   paste(tempdir(),'/result1.docx',sep = ''))
#'
renderInlineCode <- function(docxIn, docxOut, debug = F) {
    if (debug) {
        browser()
    }

  doc<-officer::read_docx(docxIn)
  smm<-officer::docx_summary(doc)

  styles<-officer::styles_info(doc)

  regx<-"^[ ]*`r[ ](.*)`$"
  smm$expr<-ifelse(grepl(regx,smm$text),sub(regx,"\\1",smm$text),NA)

  smm$values <- sapply(smm$expr, FUN = function(x) {
    eval(parse(text = x))
  })

  smm<-smm[!is.na(smm$expr),,drop=F]

  i<-3
  for(i in seq_len(nrow(smm))){
    stylei<-switch(ifelse(is.na(smm$style_name[i]), "a", "b"), a = NULL, b = styles$style_name[styles$style_id==paste0(styles$style_id[styles$style_name==smm$style_name[i]&styles$style_type=="paragraph"],"Char")])
    doc <- officer::cursor_reach(doc, keyword = paste0("\\Q",smm$text[i],"\\E")) %>% officer::body_remove() %>%
      officer::cursor_backward() %>% slip_in_text2(smm$values[i],pos="after", style = stylei)
  }

    print(doc, target = docxOut)
    return(docxOut)
}
