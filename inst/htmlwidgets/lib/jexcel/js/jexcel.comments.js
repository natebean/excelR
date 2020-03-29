if (HTMLWidgets.shinyMode) {

    // This function is used to set comments in the table
    Shiny.addCustomMessageHandler("excelR:setComments", function (message) {

        var el = document.getElementById(message[0]);
        if (el) {
            el.excel.setComments(message[1], message[2]);
        }
    });

    // This function is used to get comments  from table
    Shiny.addCustomMessageHandler("excelR:getComments", function (message) {

        var el = document.getElementById(message[0]);
        if (el) {
            var comments = message[1] ? el.excel.getComments(message[1]) : el.excel.getComments(null);
            Shiny.setInputValue(message[0],
                {
                    comments
                });
        }
    });
} 