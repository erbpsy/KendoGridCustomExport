 function excelExport() {
        var exportAll = $('.selectrow').is(":checked");
        var grid = $("#grid");
        var rows = [{
            cells: [
                { value: "column1" },
                { value: "column2" },
                { value: "column3" },
                { value: "column4" },
                { value: "column5" },
                { value: "column6" },
                { value: "column7" }
            ]
        }];
        if (exportAll) {
            var dataDource = grid.getKendoGrid();
            var trs = $("#grid").find('tr');
            for (var i = 0; i < trs.length; i++) {
                if ($(trs[i]).find(":checkbox").is(":checked")) {
                    var dataItem = dataDource.dataItem(trs[i]);
                    rows.push({
                        cells: [
                            { value: dataItem.column1 },
                            { value: dataItem.column2 },
                            { value: dataItem.column3 },
                            { value: dataItem.column4 },
                            { value: dataItem.column5 },
                            { value: dataItem.column6 },
                            { value: dataItem.column7 }
                        ]
                    })
                }
            }
        }
        else {
            var dataSource = grid.data("kendoGrid").dataSource;
            var trs = grid.find('tr');
            for (var i = 1; i < dataSource._data.length; i++) {
                    var dataItem = dataSource._data[i];
                    rows.push({
                        cells: [
                            { value: dataItem.column1 },
                            { value: dataItem.column2 },
                            { value: dataItem.column3 },
                            { value: dataItem.column4 },
                            { value: dataItem.column5 },
                            { value: dataItem.column6 },
                            { value: dataItem.column7 }
                        ]
                    })
            }
        }



        var workbook = new kendo.ooxml.Workbook({
            sheets: [
                {
                    columns: [
                        { autoWidth: true },
                        { autoWidth: true },
                        { autoWidth: true },
                        { autoWidth: true },
                        { autoWidth: true },
                        { autoWidth: true },
                        { autoWidth: true },
                        { autoWidth: true },
                        { autoWidth: true }
                    ],
                    title: "Exported Data Result",
                    rows: rows
                }
            ]
        });
        kendo.saveAs({ dataURI: workbook.toDataURL(), fileName: "ExportedData" });
    }
