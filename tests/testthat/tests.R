
test_that("Rendering inline code", {
  expect_equal(renderInlineCode(
    paste(examplePath(),'templates/template1.docx',sep = '/'),
    paste(examplePath(),'results/result1.docx',sep = '/'),debug=F),
    paste(examplePath(),'results/result1.docx',sep = '/'))
})

test_that("addFlexTables", {
  expect_equal({library(ReporteRs);
  ft_mtcars <- vanilla.table(mtcars);
  ft_iris <- vanilla.table(iris);
  FT<-list(ft_mtcars=ft_mtcars,ft_iris=ft_iris);
  addFlexTables(
    paste(examplePath(),'templates/templateFT.docx',sep = '/'),
    paste(examplePath(),'results/resultFT.docx',sep = '/'),FT,debug=F)},
  paste(examplePath(),'results/resultFT.docx',sep = '/'))
})

test_that("addPlots", {
  expect_equal({library(ggplot2);Plots<-list(plot1=function()plot(hp~wt,data=mtcars,col=cyl),plot2=function()print(ggplot(mtcars,aes(x=wt,y=hp,col=as.factor(cyl)))+geom_point()));
  addPlots(
    paste(examplePath(),'templates/templatePlots.docx',sep = ''),
    paste(examplePath(),'results/resultPlots.docx',sep = ''),
    Plots,debug=F,height=4)},
  paste(examplePath(),'results/resultPlots.docx',sep = ''))
})
