#' Returns path to examples folder
#'
#' @return Returns path to examples folder
#' @export
#' @examples
#' examplePath()

examplePath <- function() {
    p <- paste(find.package("WordR"), "inst", sep = "/")
    paste(find.package("WordR"), ifelse(dir.exists(p), "/inst", ""), "/examples/", sep = "")
}

# @description append text into a paragraph of an rdocx object.
# This function is copy of deprecated one in officer
# @param x an rdocx object
# @param str text
# @param style text style
# @param pos where to add the new element relative to the cursor,
# "after" or "before".
# @param hyperlink turn the text into an external hyperlink
slip_in_text2 <- function( x, str, style = NULL, pos = "after", hyperlink = NULL ){

  if( is.null(style) )
    style <- x$default_styles$character

  style_id <- get_style_id(x, style=style, type = "character")

  wr_ns_yes <- "<w:r xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">"

  if( is.null(hyperlink) ) {
    xml_elt <- paste0( wr_ns_yes,
                       "<w:rPr><w:rStyle w:val=\"%s\"/></w:rPr>",
                       "<w:t xml:space=\"preserve\">%s</w:t></w:r>")
    xml_elt <- sprintf(xml_elt, style_id, htmlEscapeCopy(str))
  } else {
    hyperlink_id <- paste0("rId", x$doc_obj$relationship()$get_next_id())
    x$doc_obj$relationship()$add(
      id = hyperlink_id,
      type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
      target = hyperlink,
      target_mode = "External" )

    xml_elt <- paste0( "<w:hyperlink r:id=\"%s\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">",
                       "<w:r><w:rPr><w:rStyle w:val=\"%s\"/></w:rPr>",
                       "<w:t xml:space=\"preserve\">%s</w:t></w:r></w:hyperlink>")
    xml_elt <- sprintf(xml_elt, hyperlink_id, style_id, htmlEscapeCopy(str))
  }

  slip_in_xml(x = x, str = xml_elt, pos = pos)
}

# @export
# @title add a wml string into a Word document
# @description The function add a wml string into
# the document after, before or on a cursor location.
#
# This function is copy of deprecated one in officer # @param x an rdocx object
# @param str a wml string
# @param pos where to add the new element relative to the cursor,
# "after" or "before".
# @keywords internal
#' @import xml2

slip_in_xml <- function(x, str, pos){

  xml_elt <- xml2::as_xml_document(str)
  pos_list <- c("after", "before")

  if( !pos %in% pos_list )
    stop("unknown pos ", shQuote(pos, type = "sh"), ", it should be ",
         paste( shQuote(pos_list, type = "sh"), collapse = " or ") )

  cursor_elt <- docx_current_block_xml(x)
  pos <- ifelse(pos=="after", length(xml2::xml_children(cursor_elt)), 1)
  xml2::xml_add_child(.x = cursor_elt, .value = xml_elt, .where = pos )
  x
}

# This function is copy of one in officer
htmlEscapeCopy <- local({

  .htmlSpecials <- list(
    `&` = '&amp;',
    `<` = '&lt;',
    `>` = '&gt;'
  )
  .htmlSpecialsPattern <- paste(names(.htmlSpecials), collapse='|')
  .htmlSpecialsAttrib <- c(
    .htmlSpecials,
    `'` = '&#39;',
    `"` = '&quot;',
    `\r` = '&#13;',
    `\n` = '&#10;'
  )
  .htmlSpecialsPatternAttrib <- paste(names(.htmlSpecialsAttrib), collapse='|')
  function(text, attribute=FALSE) {
    pattern <- if(attribute)
      .htmlSpecialsPatternAttrib
    else
      .htmlSpecialsPattern
    text <- enc2utf8(as.character(text))
    # Short circuit in the common case that there's nothing to escape
    if (!any(grepl(pattern, text, useBytes = TRUE)))
      return(text)
    specials <- if(attribute)
      .htmlSpecialsAttrib
    else
      .htmlSpecials
    for (chr in names(specials)) {
      text <- gsub(chr, specials[[chr]], text, fixed = TRUE, useBytes = TRUE)
    }
    Encoding(text) <- "UTF-8"
    return(text)
  }
})

# This function is copy of one in officer
get_style_id <- function(x, style, type ){
  ref <- styles_info(x, type = type)

  if(!style %in% ref$style_name){
    t_ <- shQuote(ref$style_name, type = "sh")
    t_ <- paste(t_, collapse = ", ")
    t_ <- paste0("c(", t_, ")")
    stop("could not match any style named ", shQuote(style, type = "sh"), " in ", t_, call. = FALSE)
  }
  ref$style_id[ref$style_name == style]
}
