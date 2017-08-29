#' Returns path to examples folder
#'
#' @return Returns path to examples folder
#' @examples
#' examplePath()

examplePath <- function() {
    p <- paste(find.package("WordR"), "inst", sep = "/")
    paste(find.package("WordR"), ifelse(dir.exists(p), "/inst", ""), "/examples/", sep = "")
}
