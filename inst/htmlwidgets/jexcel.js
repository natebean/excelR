(function () {
  HTMLWidgets.widget({

    name: "jexcel",

    type: "output",

    factory: function (el, width, height) {
      var elementId = el.id;
      var container = document.getElementById(elementId);
      var excel = null;

      return {
        renderValue: function (paramsFromR) {
          // get params that need adjustment or used to adjust other params
          var rowHeight = getValueOrDefault(paramsFromR, "rowHeight", undefined);
          var showToolbar = getValueOrDefault(paramsFromR, "showToolbar", false);
          var dateFormat = getValueOrDefault(paramsFromR, "dataFormat", "MM/DD/YYYY");
          var autoWidth = getValueOrDefault(paramsFromR, "autoWidth", true);
          var autoFill = getValueOrDefault(paramsFromR, "autoFill", false);
          var getSelectedData = getValueOrDefault(paramsFromR, "getSelectedData", false);

          var imageColIndex = undefined;

          var paramsToJexcel = {};

          paramsFromR = removeParms(paramsFromR);
          paramsToJexcel = extend({}, paramsFromR);
          paramsToJexcel = columnTypeHandler(paramsFromR, paramsToJexcel, dateFormat);
          paramsToJexcel.rows = rowsHandler(rowHeight);

          //Lets add the image url in the update table function
          if (imageColIndex) {
            paramsToJexcel.updateTable = function (instance, cell, col, row, val, id) {
              if (col == imageColIndex && "data:image" != val.substr(0, 10)) {
                cell.innerHTML = '<img src="' + val + '" style="width:100px;height:100px">';
              }
            }
          }

          paramsToJexcel.tableOverflow = true;
          paramsToJexcel.onchange = this.onChange;
          paramsToJexcel.onbeforechange = this.onBeforeChange;
          paramsToJexcel.oninsertrow = this.onChange;
          paramsToJexcel.ondeleterow = this.onChange;
          paramsToJexcel.oninsertcolumn = this.onChange;
          paramsToJexcel.ondeletecolumn = this.onChange;
          paramsToJexcel.onsort = this.onChange;
          paramsToJexcel.onmoverow = this.onChange;
          paramsToJexcel.onchangeheader = this.onChangeHeader;

          if (getSelectedData) {
            paramsToJexcel.onselection = this.onSelection;
          }

          // TODO:  make this configurable
          if (showToolbar) {
            paramsToJexcel.toolbar = [
              { type: 'i', content: 'add', onclick: function () { excel.insertRow(1); } },
            ]
          }

          // Snapshot selection before removing instance
          var selection = excel ? excel.selectedCell : undefined;

          // removing instance, is there a better way? 
          if (excel !== null) {
            while (container.firstChild) {
              container.removeChild(container.firstChild);
            }
          }

          excel = jexcel(container, paramsToJexcel);

          if (autoWidth) {
            // TODO:  due thru instance
            excel.table.setAttribute("style", "width: auto; height: auto; white-space: normal;")
          }

          if (!autoWidth && autoFill) {
            // TODO:  due thru instance
            excel.table.setAttribute("style", "width: 100%; height: 100%; white-space: normal;")
            // TODO:  due thru instance
            container.getElementsByClassName("jexcel_content")[0].setAttribute("style", "height:100%")
          }

          // Add cursor back to where it was at
          if (selection) {
            excel.updateSelectionFromCoords(selection[0], selection[1], selection[2], selection[3]);
          }

          container.excel = excel;

        },

        resize: function (width, height) {

        },

        onChange: function (obj, cell, x, y, value) {

          if (HTMLWidgets.shinyMode) {
            // this is the jexcel object that called this function
            var changedData = getOnChangeData(this.data, this.columns, this.colHeaders);
            var metaData = metaDataTransform(obj.jexcel.getMeta());
            Shiny.setInputValue(obj.id,
              {
                data: changedData.data,
                colHeaders: changedData.colHeaders,
                colType: changedData.colType,
                forSelectedVals: false,
                changeEventInfo: {
                  columnId: parseInt(x) + 1, // 0 index to 1 for R
                  rowId: parseInt(y) + 1, // 0 index to 1 for R
                  value: value
                },
                metaData: metaData
              })
          }
        },

        onBeforeChange: function (instance, cell, x, y, value) {
          var cellName = jexcel.getColumnNameFromId([x, y]);
          // Could be used for batch tracking
          instance.jexcel.setMeta(cellName, "changelog", "changed");
        },

        onChangeHeader: function (obj, column, oldValue, newValue) {

          if (HTMLWidgets.shinyMode) {
            // this is the jexcel object that called this function
            var changedData = getOnChangeData(this.data, this.columns, this.colHeaders);

            var newColHeader = changedData.colHeaders;
            newColHeader[parseInt(column)] = newValue;

            Shiny.setInputValue(obj.id,
              {
                data: changedData.data,
                colHeaders: newColHeader,
                colType: changedData.colType,
                forSelectedVals: false,
              })
          }
        },

        onSelection: function (obj, borderLeft, borderTop, borderRight, borderBottom, origin) {

          if (HTMLWidgets.shinyMode) {
            // TODO: move to own function
            // Get arrays between top to bottom, this will return the array of array for selected data
            // Returning data that was selected
            var data = this.data.reduce(function (acc, value, index) {
              if (index >= borderTop && index <= borderBottom) {
                var val = value.reduce(function (innerAcc, innerValue, innerIndex) {
                  if (innerIndex >= borderLeft && innerIndex <= borderRight) {
                    innerAcc.push(innerValue);
                  }
                  return innerAcc;
                }, [])
                acc.push(val);
              }
              return acc;
            }, [])

            var fullData = getOnChangeData(this.data, this.columns, this.colHeaders);

            Shiny.setInputValue(obj.id,
              {
                fullData: fullData,
                selectedData: data,
                selectedDataBoundary: {
                  borderLeft,
                  borderTop,
                  borderRight,
                  borderBottom
                },
                forSelectedVals: true
              })

          }
        }
      };
    }
  });


  function getOnChangeData(data, columns, colHeaders) {

    var colType = columns.map(function (column) {
      return column.type;
    })

    // If no headers create them.
    if (colHeaders.every(function (val) { return (val === '') })) {
      var colHeaders = columns.map(function (column) { return column.title })
    }

    return { data: data, colHeaders: colHeaders, colType: colType }
  }


  function removeParms(params) {
    delete params.dateFormat;
    delete params.rowHeight;
    delete params.autoWidth;
    delete params.getSelectedData;
    delete params.otherParams;
    return params;
  }

  function columnTypeHandler(paramsFromR, paramsToJexcel, dateFormat) {
    var columns = paramsFromR.columns;
    if (columns) {
      paramsToJexcel.columns = paramsFromR.columns.map(function (column, index) {
        // If the date format is not default we'll need to pass it properly to jexcel table
        if (column.type === "calendar") {
          column.options = { format: dateFormat };
        }
        // If image url is specified, we'll need to pass it to jexcel table only
        // in updateTable function,so here we'll first find the column index. This
        if (column.type === "image") {
          imageColIndex = index;
        }

        if (column.type === "readonly") {
          column.type = 'text';
          column.readOnly = true;
        }

        return column;
      });
    }
    return paramsToJexcel;
  }

  function rowsHandler(rowHeight) {
    var rows = (function () {
      if (rowHeight) {
        const rows = {};
        rowHeight.map(function (data) {
          return rows[data[0]] = {
            height: `${data[1]}px`
          }
        });
        return rows;
      }
      return {};
    })();
    return rows;
  }

  function metaDataTransform(metaData) {

    if (!metaData) return null;

    var transformed = {};

    // Assuming shape of metaData =
    // {A1: {obj}, B2: {obj}}
    // Key transform from 'A1' to '1-1' for example
    // 0 index to 1 index for R
    Object.keys(metaData).forEach((key) => {
      const newKeyArray = jexcel.getIdFromColumnName(key, true);
      const columnId = parseInt(newKeyArray[0]) + 1;
      const rowId = parseInt(newKeyArray[1]) + 1;
      const newKey = columnId + '-' + rowId;
      transformed[newKey] = metaData[key];
    });

    return transformed;
  }

})();
