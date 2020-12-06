# R function to view data in an Excel spreadsheet

```{r setup, include=FALSE}
knitr::opts_chunk$set(echo = TRUE)
```

Sometimes it becomes a mess when we have to view our data in the RStudio pre-defined viwer. However, the view function
comes to help us with this task by generating automatically an Excel Spreadsheet (already with the filter option!) that 
can help us to organize our data in the best possible way. Let's see the code:

```{r, echo=TRUE, message = FALSE }
library(rJava)
library(XLConnect)
```

As you can see, we are using the XLConnect package, which requires Java to be installed.
You are free to change it by other alternatives (e.g writexl)


```{r, echo=TRUE}
view <- function(data, autofilter=TRUE) {
  # data: data frame
  # autofilter: whether to apply a filter to make sorting and filtering easier
  open_command <- switch(Sys.info()[['sysname']],
                         Windows= 'open',
                         Linux  = 'xdg-open',
                         Darwin = 'open')
  require(XLConnect)
  temp_file <- paste0(tempfile(), '.xlsx')
  wb <- loadWorkbook(temp_file, create = TRUE)
  createSheet(wb, name = "temp")
  writeWorksheet(wb, data, sheet = "temp", startRow = 1, startCol = 1)
  if (autofilter) setAutoFilter(wb, 'temp', aref('A1', dim(data)))
  saveWorkbook(wb, )
  system(paste(open_command, temp_file))
}

# we use the pre-loaded cars database to check our function...

#view(cars) # This generates an excel spreadsheet in such a way that you can visualize 
# your data better than with the RStudio viewer

#View(cars)  # The V(capital V)iew function initialize the RStudio viewer

```


