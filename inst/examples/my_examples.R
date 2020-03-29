# ### some additional arguments with mtcars ----
# library(excelR)
#
# # for now we need to manipulate a data.frame with rownames
# mtcars2 <-
#   cbind(name = rownames(mtcars), mtcars, stringsAsFactors = FALSE)
# mtcars2$name <- rownames(mtcars)
# # change rownames so jsonlite will not convert to a column
# rownames(mtcars2) <- seq_len(nrow(mtcars2))
#
# excelTable(
#   data = mtcars2,
#   colHeaders = toupper(colnames(mtcars2)),
#   # upper case the column names
#   fullscreen = TRUE,
#   # fill screen with table
#   columnDrag = FALSE,
#   # allow dragging (reordering) of columns
#   rowDrag = FALSE,
#   # disallow dragging (reordering) of rows
#   wordWrap = TRUE,
#   # wrap text in a cell if is longer than cell width,
#   defaultColWidth = 300
# )



library(excelR)

data = data.frame(
  Make = c('Honda', 'Honda', 'Hyundai'),
  Car = c('Civic', 'City', 'Polo'),
  MyDate = c(as.Date('2000-1-1'), as.Date('2010-2-1'), as.Date('2009-3-1'))
)

columns = data.frame(
  title = c('Make', 'Car', 'My Date'),
  width = c(600, 600, 600),
  type = c('text', 'dropdown', 'calendar'),
  source = I(list(
    0,
    c('Civic', 'City',  'Polo', 'Creta', 'Santro'),
    0
  ))
)

row_height <- data.frame(
  row_index = c(0,2),
  row_value = c(100,150)
)


excelTable(
  data = data,
  columns = columns,
  rowHeight = row_height,
  fullscreen = TRUE,
  wordWrap = TRUE,
  showToolbar = TRUE,
  columnDrag = FALSE,
  allowInsertColumn = FALSE,
  allowRenameColumn = FALSE,
  allowDeleteColumn = FALSE,
  defaultColWidth = 300,
  toolBar = TRUE
)
