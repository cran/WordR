
test_that("Rendering inline code", {
  expect_equal(renderInlineCode(
    paste(examplePath(),'templates/template1.docx',sep = '/'),
    paste(tempdir(),'/result1.docx',sep = '/'),debug=F),
    paste(tempdir(),'/result1.docx',sep = '/'))
})

test_that("body_add_flextables", {
  expect_equal({
    library(flextable)
    ft_mtcars <- flextable(mtcars)
    ft_iris <- flextable(iris)
    FT <- list(ft_mtcars=ft_mtcars,ft_iris=ft_iris)
    body_add_flextables(
      paste(examplePath(),'templates/templateFT.docx',sep = ''),
      paste(tempdir(),'/resultFT.docx',sep = ''),
      FT)
  },
  paste(tempdir(),'/resultFT.docx',sep = ''))
})

test_that("addPlots", {
  expect_equal({library(ggplot2);Plots<-list(plot1=function()plot(hp~wt,data=mtcars,col=cyl),plot2=function()print(ggplot(mtcars,aes(x=wt,y=hp,col=as.factor(cyl)))+geom_point()));
  addPlots(
    paste(examplePath(),'templates/templatePlots.docx',sep = ''),
    paste(tempdir(),'/resultPlots.docx',sep = ''),
    Plots,debug=F,height=4)},
  paste(tempdir(),'/resultPlots.docx',sep = ''))
})

test_that("addPlots2", {
  library(ggplot2)
  outpath<-paste(tempdir(),'/resultPlots2.docx',sep = '')
  Plots<-list(plot1=function(){plot(hp~wt,data=mtcars,col=cyl)},
              plot2=(ggplot(mtcars,aes(x=wt,y=hp,col=as.factor(cyl)))+geom_point()))
  resFile<-addPlots(
    paste(examplePath(),'templates/templatePlots.docx',sep = ''),
    outpath,
    Plots,debug=F,height=4, width=c(3,6))

  expect_equal(resFile, outpath)
})



test_that("addPlots3", {
  library(ggplot2)
  outpath<-paste(tempdir(),.Platform$file.sep,'resultPlots3.docx',sep = '')
  Plots<-list(plot1=function(){plot(hp~wt,data=mtcars,col=cyl)},
              plot2=(ggplot(mtcars,aes(x=wt,y=hp,col=as.factor(cyl)))+geom_point()))
  resFile<-addPlots(docxIn= officer::read_docx(path =paste(examplePath(),'templates/templatePlots.docx',sep = '')),
                    docxOut=NA,
                    Plots,debug=F,height=4, width=c(3,6))

  resFilePath<-print(resFile, target = outpath)

  expect_equal(basename(resFilePath), 'resultPlots3.docx')
  expect_equal(class(resFile), "rdocx")
})

test_that("renderAll", {
  library(flextable)
  library(ggplot2)
  outpath<-paste(tempdir(),'/resultAll.docx',sep = '')
  ft_mtcars <- flextable(mtcars)
  FT <- list(mtcars=ft_mtcars)
  Plots <- list(mtcars1 = ggplot(mtcars, aes(x = wt, y = hp, col = as.factor(cyl))) + geom_point())
  resFile <- renderAll(docxIn = paste0(examplePath(), "templates/templateAll.docx"),
                       docxOut = outpath,
                       flextables = FT,
                       Plots = Plots)
  expect_equal(resFile, outpath)
})


