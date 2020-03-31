library(shiny)
library(excelR)

app <- shinyApp(
  ui = fluidPage(
    actionButton("change", "change"),
    excelOutput("table", height = "700px")
  )
  ,

  server = function(input, output) {
    state_store <- reactiveValues(
      df =
        data.frame(
          ID = c(1, 2, 3),
          Make = c('Honda', 'Honda', 'Hyundai'),
          Car = c('Civic', 'City', 'Polo'),
          MyDate = c(
            as.Date('2000-1-1'),
            as.Date('2010-2-1'),
            as.Date('2009-3-1')
          ),
          State = c('server', 'server', 'server')
        ),
      style = list(),
      comments = list(A1 = 'comments: A1', B2 = 'comments: B2'),
      meta_data = list(
        B1 = list(customKey = list(state = "valid"))
      )
      # meta_data = list()
    )

    next_table <-
      data.frame(
        ID = c(1, 2, 3),
        Make = c('one', 'two', 'three'),
        Car = c('Civic', 'City', 'Polo'),
        MyDate = c(
          as.Date('2000-1-1'),
          as.Date('2010-2-1'),
          as.Date('2009-3-1')
        ),
        State = c('server', 'server', 'server')
      )


    output$table <- renderExcel({
      columns <- data.frame(
        title = c('ID', 'Make', 'Car', 'My Date', 'State'),
        width = c(100, 600, 600, 600, 600),
        type = c('readonly', 'text', 'dropdown', 'calendar', 'hidden'),
        source = I(list(
          0,
          0,
          c('Civic', 'City',  'Polo', 'Creta', 'Santro'),
          0, 0
        ))
      )

      # row_height <- data.frame(row_index = c(0, 2),
      #                          row_value = c(100, 150))

      excel_table <- excelTable(
        data = state_store$df,
        columns = columns,
        style = state_store$style,
        wordWrap = TRUE,
        showToolbar = TRUE,
        columnDrag = TRUE,
        allowInsertColumn = FALSE,
        allowRenameColumn = FALSE,
        allowDeleteColumn = FALSE,
        defaultColWidth = 300,
        toolBar = TRUE,
        tableHeight = '700px' ,
        comments = state_store$comments,
        metaData = state_store$meta_data
      )

      excel_table
    })

    observeEvent(input$table, {
      print(input$table$changeEventInfo)

      table_data <- excel_to_R(input$table)

      style <- input$table$style
      comments <- input$table$comments
      meta_data <- input$table$metaData

      if (!is.null(table_data)) {
        if (isTruthy(input$table$changeEventInfo$value)) {
          # If came from the server, then change state to changed else is a client record
          if (table_data[input$table$changeEventInfo$rowId,]$State == "server") {
            table_data[input$table$changeEventInfo$rowId,]$State = "changed"
          }
        } else{
          #  Came from the client not the server
          new_rows <- table_data[which(table_data$State == ''),]
          if (nrow(new_rows) > 0) {
            table_data[which(table_data$State == ''),]$State = "new"
          }
        }

        if (isTruthy(input$table$changeEventInfo$cellName)) {
          style[[input$table$changeEventInfo$cellName]] = 'background-color:orange; color:green;'
          comments[[input$table$changeEventInfo$cellName]] = 'changed'
        }
        # print(input$table$style)
        # print(table_data)
        # print(style)

        print(input$table)
        # print(table_data)

        state_store$df <- table_data
        state_store$style <- style
        state_store$comments <- comments
        state_store$meta_data <- meta_data
      }
    })

    observeEvent(input$change, {
      state_store$df <- next_table
      print("change")
    })

  }
)

# runApp(app)
