############################## rendering inline code example
renderInlineCode(paste(examplePath(), "templates/template1.docx", sep = ""),
                 paste(examplePath(), "results/result1.docx", sep = ""))

############################## adding flextables example
library(flextable)
ft_mtcars <- flextable(mtcars)
ft_iris <- flextable(iris)
FT <- list(ft_mtcars=ft_mtcars,ft_iris=ft_iris)
body_add_flextables(
  paste(examplePath(),'templates/templateFT.docx',sep = ''),
  paste(examplePath(),'/resultFT.docx',sep = ''),
  FT)

############################## adding Plots example
library(ggplot2)
Plots <- list(plot1 = function() plot(hp ~ wt, data = mtcars, col = cyl), plot2 = function() print(ggplot(mtcars, aes(x = wt, y = hp, col = as.factor(cyl))) + geom_point()))
addPlots(paste(examplePath(), "templates/templatePlots.docx", sep = ""),
         paste(examplePath(), "results/resultPlots.docx", sep = ""), Plots, height = 4)

############################## combining all together
library(flextable)
library(ggplot2)
# prepare outputs
ft_mtcars <- flextable(mtcars)
FT <- list(mtcars=ft_mtcars)
Plots <- list(mtcars1 = function() print(ggplot(mtcars, aes(x = wt, y = hp, col = as.factor(cyl))) + geom_point()))

# render docx
renderInlineCode(paste(examplePath(), "templates/templateAll.docx", sep = ""),
                 paste(examplePath(), "results/templateAll.docx", sep = ""))

body_add_flextables(paste(examplePath(), "results/templateAll.docx", sep = ""),
              paste(examplePath(), "results/templateAll.docx", sep = ""), FT)

addPlots(paste(examplePath(), "results/templateAll.docx", sep = ""),
         paste(examplePath(), "results/templateAll.docx", sep = ""), Plots, height = 4)
