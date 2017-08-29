library(officer)
library(dplyr)

############################## listing available styles
paste(find.package("WordR"), "examples/templates/template1.docx", sep = "/") %>% read_docx() %>% styles_info() %>%
  filter(style_type == "character") %>% select(style_name)

############################## rendering inline code example
renderInlineCode(paste(examplePath(), "templates/template1.docx", sep = ""),
                 paste(examplePath(), "results/result1.docx", sep = ""))

############################## adding Flex Tables example
library(ReporteRs)
ft_mtcars <- vanilla.table(mtcars)
ft_iris <- vanilla.table(iris)
FT <- list(ft_mtcars = ft_mtcars, ft_iris = ft_iris)
addFlexTables(paste(examplePath(), "templates/templateFT.docx", sep = ""), paste(examplePath(), "results/resultFT.docx", sep = ""), FT)

############################## adding Plots example

Plots <- list(plot1 = function() plot(hp ~ wt, data = mtcars, col = cyl), plot2 = function() print(ggplot(mtcars, aes(x = wt, y = hp, col = as.factor(cyl))) + geom_point()))
addPlots(paste(examplePath(), "templates/templatePlots.docx", sep = ""),
         paste(examplePath(), "results/resultPlots.docx", sep = ""), Plots, height = 4)

############################## combining all together
library(ReporteRs)
library(ggplot2)
# prepare outputs
ft_mtcars <- vanilla.table(mtcars)
FT <- list(mtcars = ft_mtcars)
Plots <- list(mtcars1 = function() print(ggplot(mtcars, aes(x = wt, y = hp, col = as.factor(cyl))) + geom_point()))

# render docx
renderInlineCode(paste(examplePath(), "templates/templateAll.docx", sep = ""),
                 paste(examplePath(), "results/templateAll.docx", sep = ""))

addFlexTables(paste(examplePath(), "results/templateAll.docx", sep = ""),
              paste(examplePath(), "results/templateAll.docx", sep = ""), FT)

addPlots(paste(examplePath(), "results/templateAll.docx", sep = ""),
         paste(examplePath(), "results/templateAll.docx", sep = ""), Plots, height = 4)
