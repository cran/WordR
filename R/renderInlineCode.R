#' Read Word document with R code blocks, evaluate them and writes the result into another Word document.
#'
#' @param docxIn String of length one; path to Word file with R code blocks.
#' @param docxOut String of length one; path for output Word file.
#' @param defaultStyle String of length one or NA; Default style for inserted text. It must be one of the styles in the document.
#' List of the styles can be obtained by running \code{\link[officer]{styles_info}(read_docx(docxIn)) \%>\% filter(style_type==character) \%>\% select(style_name)}
#' @param debug Boolean of length one; If \code{True} then \code{\link[base]{browser}()} is called at the beginning of the function
#' @return Path to the rendered Word file if the operation was successfull.
#' @examples
#' library(ReporteRs)
#' renderInlineCode(
#'   paste(examplePath(),'templates/template1.docx',sep = ''),
#'   paste(examplePath(),'results/result1.docx',sep = ''))
#'
renderInlineCode <- function(docxIn, docxOut, defaultStyle = NA, debug = F) {
    if (debug) {
        browser()
    }
    doc <- ReporteRs::docx(template = docxIn)

    extractedText <- ReporteRs::text_extract(doc)

    extractedTextCollapsed <- paste(extractedText, collapse = "", sep = "")

    starts <- gregexpr("#r[[:digit:]]* ", extractedTextCollapsed, fixed = F)[[1]]
    ends <- gregexpr("r#", extractedTextCollapsed, fixed = T)[[1]]

    expressionsPars <- matrix(c(starts, ends), ncol = 2) %>% apply(1, function(x) {
        if (x[1] >= x[2]) {
            stop("Wrong sequence of inline R code separator (ending found before beginning)!")
        }
        substr(extractedTextCollapsed, 2 + x[1], x[2] - 1)
    })

    rexp <- "^(\\w+)\\s?(.*)$"
    expressions <- data.frame(ID = sub(rexp, "\\1", expressionsPars), expression2 = sub(rexp, "\\2", expressionsPars), stringsAsFactors = F)
    expressions$expression <- sub("@[{].*[}]$", "\\1", expressions$expression2)
    expressions$style <- sub("(.*)[}]", "\\1", sub("^.*@[{]", "", expressions$expression2))
    expressions$style <- ifelse(expressions$style == expressions$expression2, defaultStyle[1], expressions$style)
    if (any(duplicated(expressions$ID))) {
        stop(paste("Duplicate inline code IDs:", paste(expressions$ID[duplicated(expressions$ID)], sep = ",", collapse = ",")))
    }

    values <- sapply(expressions$expression, FUN = function(x) {
        eval(parse(text = x))
    })

    docA <- officer::read_docx(docxIn)

    for (i in seq_len(nrow(expressions))) {
        ids <- paste0("#r", expressions$ID[i])
        nfound <- grep(paste0("^", ids, " "), extractedText) %>% length
        if (nfound == 0) {
            stop(paste("R Inline code chunk ID:", ids, "not found. Pls reidentificate!(retype the \"#rXXXXX[space]\")"))
        }
        if (nfound > 1) {
            stop(paste("R Inline code chunk ID:", ids, "found in multiple places. Pls check!"))
        }
        docA <- officer::cursor_reach(docA, keyword = ids) %>% officer::body_remove() %>% officer::cursor_backward() %>% officer::slip_in_text(values[i], pos = "before",
            style = switch(ifelse(is.na(expressions$style[i]), "a", "b"), a = NULL, b = expressions$style[i]))
    }

    print(docA, target = docxOut)
    return(docxOut)
}
