(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // 每次加载新页面时都必须运行初始化函数。
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // 初始化通知机制并隐藏它
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            // 如果未使用 Excel 2016，请使用回退逻辑。
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $("#template-description").text("此示例将显示电子表格中选定单元格的值。");
                $('#button-text').text("显示!");
                $('#button-desc').text("显示所选内容");

                //$('#highlight-button').click(displaySelectedCells);
                return;
            }

            //$("#template-description").text("此示例将突出显示电子表格中选定单元格的最高值。");
            //$('#createBarChart-text').text("createBarChart");
            //$('#button-desc').text("突出显示最大数字。");

            //loadSampleData();

            // 为突出显示按钮添加单击事件处理程序。
            //$('#highlight-button').click(hightlightHighestValue);

            $('#createBarChart-text').text("createBarChart");
            $('#createBarChart').click(CreateBarChart);

            $('#playBarChart-text').text("playBarChart");
            $('#playBarChart').click(PlayBarChart);

            $('#createColumnChart-text').text("createColumnChart");
            $('#createColumnChart').click(CreateColumnChart);

            $('#playColumnChart-text').text("playColumnChart");
            $('#playColumnChart').click(PlayColumnChart);

            $('#sampleData-text').text("SampleData");
            $('#sampleData').click(loadSampleData);

        });
    };

    /**
     * create for bar chart
     */
    async function CreateBarChart() {
        try {
            await CreateBarOrColumnChart(barChartFlag);
        } catch (error) {
            console.error(error);
        }
    }

    /**
     * play for bar chart
     */
    async function PlayBarChart() {
        try {
            await PlayBarOrColumnChart(barChartFlag);
        } catch (error) {
            console.error(error);
        }
    }

    /**
     * create for column chart
     */
    async function CreateColumnChart() {
        try {
            await CreateBarOrColumnChart(columnChartFlag);
        } catch (error) {
            console.error(error);
        }
    }

    /**
     * play for column chart
     */
    async function PlayColumnChart() {
        try {
            await PlayBarOrColumnChart(columnChartFlag);
        } catch (error) {
            console.error(error);
        }
    }

    let activeTableId;
    let inputPointItems;
    let orientation;

    let pointItemsCount;

    //for original table
    let totalColumnCount;
    let totalRowCount;

    //--------------------------------
    // Parameters. Modify it if needed.
    const chartWidth = 600,
        chartHeight = 400,
        chartLeft = 150,
        chartTop = 50;

    const splitIncreasement = 2;

    const colorList = [
        "#afc97a",
        "#cd7371",
        "#729aca",
        "#b65708",
        "#276a7c",
        "#4d3b62",
        "#5f7530",
        "#772c2a",
        "#2c4d75",
        "#f79646",
        "#4bacc6",
        "#8064a2",
        "#9bbb59",
        "#c0504d",
        "#4f81bd"
    ];
    const fontSize_Title = 28,
        fontSize_CategoryName = 13,
        fontSize_AxisValue = 11,
        fontSize_DataLabel = 13;

    // Internal used const. DO NOT CHANGE
    //for barChart and columnChart and tool
    const toolSheetName = "toolSheet9527"; //+UUID
    const toolTableName = "toolTable9527"; //+UUID

    //for line chart
    let lineChartName = "LineChartName9527";
    let linePointSetLabel;
    let linePointUnsetLabel;
    // let seriesUpdate;

    const barChartName = "BarChartName9527";
    const barChartFlag = 1;

    const columnChartName = "ColumnChartName9527";
    const columnChartFlag = 2;

    let ToolTableColumnIndex = {
        category: 0,
        value: 1,
        color: 2,
        map: 3
    }

    let ToolTableColumnName = {
        category: "categoryColumn",
        //value
        color: "colorColumn",
        map: "mapColumn"
    }



    /**
     * create for bar or column chart
     */
    async function CreateBarOrColumnChart(flag) {
        try {
            await Excel.run(async context => {
                // Find selected table
                const activeRange = context.workbook.getSelectedRange();
                let dataTables = activeRange.getTables(false);
                dataTables.load("items");
                await context.sync();

                // Get active table
                let dataTable = dataTables.items[0];
                let dataSheet = context.workbook.worksheets.getActiveWorksheet();
                activeTableId = dataTable.id; //id can not be loaded
                let table = dataSheet.tables.getItem(activeTableId);
                await context.sync();

                let wholeRange = table.getRange();
                wholeRange.load("rowCount");
                wholeRange.load("columnCount");
                await context.sync();
                totalColumnCount = wholeRange.columnCount;
                totalRowCount = wholeRange.rowCount;

                //create toolTable
                //delete the old chart and sheet
                let toolSheet;
                toolSheet = context.workbook.worksheets.getItemOrNullObject(toolSheetName);
                toolSheet.load();
                await context.sync();
                let lastBarChart;
                let lastColumnChart;

                if (JSON.stringify(toolSheet) !== "{}") {
                    lastBarChart = dataSheet.charts.getItemOrNullObject(barChartName);
                    lastColumnChart = dataSheet.charts.getItemOrNullObject(columnChartName);
                    //chart delete
                    lastBarChart.load();
                    lastColumnChart.load();
                    await context.sync();
                    if (JSON.stringify(lastBarChart) !== "{}") {
                        lastBarChart.delete();
                    }
                    if (JSON.stringify(lastColumnChart) !== "{}") {
                        lastColumnChart.delete();
                    }
                    toolSheet.delete();
                }
                toolSheet = context.workbook.worksheets.add(toolSheetName);

                hiddenSheet(toolSheet);
                await context.sync();

                let toolRange = toolSheet.getCell(0, 0).getAbsoluteResizedRange(totalRowCount, 4);
                let toolTable = toolSheet.tables.add(toolRange, true);
                toolTable.set({
                    name: toolTableName
                });
                //set columnName
                toolTable.columns
                    .getItemAt(ToolTableColumnIndex.category)
                    .set({ name: ToolTableColumnName.category });
                toolTable.columns
                    .getItemAt(ToolTableColumnIndex.color)
                    .set({ name: ToolTableColumnName.color });
                toolTable.columns
                    .getItemAt(ToolTableColumnIndex.map)
                    .set({ name: ToolTableColumnName.map });

                let categoryBodyRange = toolTable.columns.getItem(ToolTableColumnName.category).getDataBodyRange();
                let curIteratedRange = toolTable.columns.getItemAt(ToolTableColumnIndex.value).getRange();
                let curIteratedBodyRange = toolTable.columns
                    .getItemAt(ToolTableColumnIndex.value)
                    .getDataBodyRange();
                let colorBodyRange = toolTable.columns.getItem(ToolTableColumnName.color).getDataBodyRange();
                let mapBodyRange = toolTable.columns.getItem(ToolTableColumnName.map).getDataBodyRange();

                //copy Range
                categoryBodyRange.copyFrom(table.columns.getItemAt(0).getDataBodyRange());
                curIteratedRange.copyFrom(table.columns.getItemAt(1).getRange()); //copy headers too

                colorBodyRange.load("values");
                await context.sync();
                let tmpColorArr = [];
                for (let i = 0; i < totalRowCount - 1; ++i) {
                    tmpColorArr.push([colorList[i % colorList.length]]);
                }
                colorBodyRange.values = tmpColorArr;

                mapBodyRange.load("values");
                await context.sync();
                let tmpMapArr = [];
                for (let i = 1; i < totalRowCount; ++i) {
                    tmpMapArr.push([i]);
                }
                mapBodyRange.values = tmpMapArr;

                //input
                let inputElement = document.getElementById("PointItems");
                inputPointItems = inputElement.value;
                let optionElement = document.getElementById("orientation");
                orientation = Number(optionElement.value);

                //get input and target items
                pointItemsCount = formatInput(inputPointItems, totalRowCount);

                let targetIteratedBodyRange = getPartialRange(
                    curIteratedBodyRange,
                    pointItemsCount,
                    orientation
                );
                let targetCategoryRange = getPartialRange(
                    categoryBodyRange,
                    pointItemsCount,
                    orientation
                );

                // Create Chart
                toolTable.sort.apply([{ key: 1, ascending: true }], true); //toolTable only does ascending sort
                let chart;
                if (flag === barChartFlag) {
                    chart = dataSheet.charts.add(Excel.ChartType.barClustered, targetIteratedBodyRange);
                    chart.set({
                        name: barChartName,
                        height: chartHeight,
                        width: chartWidth,
                        left: chartLeft,
                        top: chartTop
                    });
                } else {
                    chart = dataSheet.charts.add(Excel.ChartType.columnClustered, targetIteratedBodyRange);
                    chart.set({
                        name: columnChartName,
                        height: chartHeight,
                        width: chartWidth,
                        left: chartLeft,
                        top: chartTop
                    });
                }

                let curheaderRange = curIteratedRange.getCell(0, 0);
                curheaderRange.load("text");
                await context.sync();
                // Set chart tile and style
                chart.title.text = curheaderRange.text[0][0];
                chart.title.format.font.set({ size: fontSize_Title });
                chart.legend.set({ visible: false });

                // Set Axis
                let categoryAxis = chart.axes.getItem(Excel.ChartAxisType.category);
                categoryAxis.setCategoryNames(targetCategoryRange);
                categoryAxis.set({ visible: true });
                categoryAxis.format.font.set({ size: fontSize_CategoryName });
                let valueAxis = chart.axes.getItem(Excel.ChartAxisType.value);
                valueAxis.format.font.set({ size: fontSize_AxisValue });

                let series = chart.series.getItemAt(0);
                series.set({ hasDataLabels: true, gapWidth: 30 });
                series.dataLabels.set({ showCategoryName: false, numberFormat: "#,##0" });
                series.dataLabels.format.font.set({ size: fontSize_DataLabel });
                series.points.load();
                await context.sync();

                colorBodyRange.load("values");
                await context.sync();
                let sortedColorArr = colorBodyRange.values;

                // Set data points color
                for (let i = 0; i < series.points.count; i++) {
                    if (orientation === 1) {
                        series.points
                            .getItemAt(i)
                            .format.fill.setSolidColor(
                                sortedColorArr[totalRowCount - pointItemsCount - 1 + i][0]
                            );
                    } else {
                        series.points.getItemAt(i).format.fill.setSolidColor(sortedColorArr[i][0]);
                    }
                }
                series.points.load();

                await context.sync();
            });
        } catch (error) {
            console.error(error);
        }
    }

    /**
     * play bar or column chart
     */
    async function PlayBarOrColumnChart(flag) {
        try {
            await Excel.run(async context => {
                console.log(inputPointItems);
                console.log(orientation);
                console.log(pointItemsCount);

                let dataSheet = context.workbook.worksheets.getActiveWorksheet();
                let table = dataSheet.tables.getItem(activeTableId);

                //get toolTable
                let toolSheet = context.workbook.worksheets.getItem(toolSheetName);
                // let toolTable = dataSheet.tables.getItem(toolTableName);
                let toolTable = toolSheet.tables.getItem(toolTableName);

                let categoryBodyRange = toolTable.columns
                    .getItemAt(ToolTableColumnIndex.category)
                    .getDataBodyRange();
                let curIteratedHeaderRange = toolTable.columns
                    .getItemAt(ToolTableColumnIndex.value)
                    .getHeaderRowRange();
                let curIteratedBodyRange = toolTable.columns
                    .getItemAt(ToolTableColumnIndex.value)
                    .getDataBodyRange();
                let mapBodyRange = toolTable.columns.getItem(ToolTableColumnName.map).getDataBodyRange();
                let colorBodyRange = toolTable.columns.getItem(ToolTableColumnName.color).getDataBodyRange();

                let chart;
                if (flag == barChartFlag) {
                    chart = dataSheet.charts.getItem(barChartName);
                } else {
                    chart = dataSheet.charts.getItem(columnChartName);
                }
                //todo splitIncreasement input

                categoryBodyRange.load("values");
                curIteratedBodyRange.load("values");
                mapBodyRange.load("values");
                colorBodyRange.load("values");
                await context.sync();

                //initial countryArr
                let countryArray = [];
                for (let i = 0; i < totalRowCount - 1; ++i) {
                    let curCategory = categoryBodyRange.values[i][0];
                    let curValue = curIteratedBodyRange.values[i][0];
                    let curMap = mapBodyRange.values[i][0];
                    let curColor = colorBodyRange.values[i][0];

                    let curCountry = new Country(curCategory, curValue, curMap, 0, curColor);
                    countryArray.push(curCountry);
                }

                // paly
                for (let i = 2; i < totalColumnCount; ++i) {
                    let nextIteratedHeaderRange = table.columns.getItemAt(i).getHeaderRowRange(); //from table
                    let nextIteratedRange = table.columns.getItemAt(i).getRange();

                    nextIteratedRange.load("values");
                    curIteratedBodyRange.load("values");
                    curIteratedHeaderRange.load("text");
                    mapBodyRange.load("values");
                    await context.sync();

                    let nextArr = mapTargetRangeValue(mapBodyRange, nextIteratedRange);
                    // Calculate increase based on current value and next value
                    let increaseData = calculateIncrease(curIteratedBodyRange.values, nextArr, splitIncreasement);

                    for (let j = 0; j < totalRowCount - 1; ++j) {
                        countryArray[j].setIncreasement(increaseData[j]);
                    }

                    for (let step = 1; step <= splitIncreasement; step++) {
                        if (step === splitIncreasement) {
                            mapBodyRange.load("values");
                            await context.sync();
                            //The mapRange here is the one that was ordered in the previous 'else', and you'll have to take it again because countryArr already sorted.
                            nextArr = mapTargetRangeValue(mapBodyRange, nextIteratedRange);

                            for (let j = 0; j < totalRowCount - 1; ++j) {
                                countryArray[j].setValue(nextArr[j][0]);
                            }
                            //set title
                            curIteratedHeaderRange.copyFrom(nextIteratedHeaderRange);
                        } else {
                            // Add increase amount
                            for (let j = 0; j < totalRowCount - 1; ++j) {
                                countryArray[j].updateIncrease();
                            }
                        }

                        //sort
                        countryArray.sort((a, b) => a.value - b.value); //countryArray only does ascending sort

                        //set some value to excel Range
                        let categoryArray = [];
                        let valueArray = [];
                        let mapArray = [];
                        let colorArray = [];
                        for (let j = 0; j < totalRowCount - 1; ++j) {
                            categoryArray.push([countryArray[j].name]);
                            valueArray.push([countryArray[j].value]); //the chart will use this column
                            mapArray.push([countryArray[j].mapColumn]); //this column will be used to map row's number
                            colorArray.push([countryArray[j].color]);
                        }
                        categoryBodyRange.values = categoryArray;
                        curIteratedBodyRange.values = valueArray;
                        mapBodyRange.values = mapArray;
                        colorBodyRange.values = colorArray;
                        await context.sync();

                        // Set data points color
                        let series = chart.series.getItemAt(0);
                        series.load("points");
                        colorBodyRange.load("values");
                        await context.sync();
                        let tmpColorArr = colorBodyRange.values;
                        for (let k = 0; k < series.points.count; k++) {
                            if (orientation === 1) {
                                series.points
                                    .getItemAt(k)
                                    .format.fill.setSolidColor(
                                        tmpColorArr[totalRowCount - pointItemsCount - 1 + k][0]
                                    );
                            } else {
                                series.points.getItemAt(k).format.fill.setSolidColor(tmpColorArr[k][0]);
                            }
                        }
                        series.points.load();
                        await context.sync();
                    }

                    curIteratedHeaderRange.load("text");
                    await context.sync();
                    chart.title.text = curIteratedHeaderRange.text[0][0];
                    await context.sync();
                }

                await context.sync();
            });
        } catch (error) {
            console.error(error);
        }
    }

    function formatInput(input, rowCount) {
        let pointItemsCount = Number(input);
        if (isNaN(pointItemsCount) || pointItemsCount <= 0 || pointItemsCount > rowCount || String(input).indexOf(".") >= 0) {
            console.log("please input a integer");
            pointItemsCount = rowCount - 1;
        }
        return pointItemsCount;
    }

    /**
     * @param originalRange : BodyRange
     * @param pointItemsCount : itemscount that u want
     */
    function getPartialRange(originalRange, pointItemsCount, orientation) {
        let partialRange;
        if (orientation === 1) {
            //for top n
            partialRange = originalRange
                .getCell(totalRowCount - pointItemsCount - 1, 0)
                .getAbsoluteResizedRange(pointItemsCount, 1);
        } else {
            partialRange = originalRange.getCell(0, 0).getAbsoluteResizedRange(pointItemsCount, 1);
        }

        return partialRange;
    }

    /**
     * @param originalRange : BodyRange
     * @param pointItemsCount : itemscount that u want
     */
    function getLinePartialRange(originalCell, pointItemsCount, columnsCount) {
        let partialRange;
        partialRange = originalCell.getAbsoluteResizedRange(pointItemsCount + 1, columnsCount);
        return partialRange;
    }

    // To calculate the increase for each step between next data list and current data list
    //function calculateIncrease(current: Array<Array<number>>, next: Array<Array<number>>, steps: number) {
    function calculateIncrease(current, next, steps) {
        if (current.length != next.length) {
            console.error("Error! current data length:" + current.length + ", next data length" + next.length + ".");
        }

        let result = [];
        for (let i = 0; i < current.length; i++) {
            let increasement = (next[i][0] - current[i][0]) / steps;
            result[i] = increasement;
        }

        return result;
    }

    function mapTargetRangeValue(mapRange, targetRange) {
        let targetArr = [];
        let mapArr = mapRange.values;
        for (let j = 0; j < mapArr.length; ++j) {
            let mapIndex = mapArr[j][0];
            let mapVal = targetRange.values[mapIndex][0];
            targetArr.push([mapVal]);
        }
        return targetArr;
    }

    function hiddenSheet(sheet) {
        sheet.set({ visibility: "Hidden" });
        // sheet.set({ visibility: "Visible"});
    }

    function sleep(sleepTime) {
        var start = new Date().getTime();
        for (var i = 0; i < 1e7; i++) {
            if (new Date().getTime() - start > sleepTime) {
                break;
            }
        }
    }

    function loadSampleData() {

        let category = [
            ['category'], ['c1'], ['c2'], ['c3'], ['c4'], ['c5'], ['c6'], ['c7'], ['c8'], ['c9'], ['c10']
        ];

        let timeLine = [['t1', 't2', 't3', 't4', 't5', 't6', 't7', 't8', 't9', 't10']];

        var values = [
            [1128, 1731, 2084, 2546, 3144, 3927, 4680, 6092, 7424, 9172],
            [69, 89, 106, 126, 161, 233, 338, 445, 572, 717],
            [74, 74, 150, 179, 248, 365, 430, 613, 909, 1648],
            [117, 150, 196, 240, 349, 565, 684, 939, 1112, 1139],
            [593, 1501, 2336, 2922, 3513, 4747, 5823, 6566, 7161, 8042],
            [100, 130, 204, 257, 377, 577, 716, 949, 1209, 1606],
            [24, 27, 46, 66, 81, 210, 228, 281, 337, 476],
            [35, 40, 51, 85, 90, 163, 206, 273, 319, 373],
            [3736, 4335, 5186, 5621, 6088, 6593, 7041, 7313, 7478, 7513],
            [7, 10, 24, 38, 82, 128, 128, 265, 321, 382]
        ];

        // 针对 Excel 对象模型运行批处理操作
        Excel.run(function (ctx) {
            // 为活动工作表创建代理对象
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            // 将向电子表格写入示例数据的命令插入队列
            //sheet.getRange("B3:D5").values = values;

            sheet.getRange("A1:A11").values = category;
            sheet.getRange("B1:K1").values = timeLine;
            sheet.getRange("B2:K11").values = values;

            sheet.tables.add("A1:K11", true);
            // 运行排队的命令，并返回承诺表示任务完成
            return ctx.sync();
        })
            .catch(errorHandler);
    }

    //function hightlightHighestValue() {
    //    // 针对 Excel 对象模型运行批处理操作
    //    Excel.run(function (ctx) {
    //        // 创建选定范围的代理对象，并加载其属性
    //        var sourceRange = ctx.workbook.getSelectedRange().load("values, rowCount, columnCount");

    //        // 运行排队的命令，并返回承诺表示任务完成
    //        return ctx.sync()
    //            .then(function () {
    //                var highestRow = 0;
    //                var highestCol = 0;
    //                var highestValue = sourceRange.values[0][0];

    //                // 找到要突出显示的单元格
    //                for (var i = 0; i < sourceRange.rowCount; i++) {
    //                    for (var j = 0; j < sourceRange.columnCount; j++) {
    //                        if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] > highestValue) {
    //                            highestRow = i;
    //                            highestCol = j;
    //                            highestValue = sourceRange.values[i][j];
    //                        }
    //                    }
    //                }

    //                cellToHighlight = sourceRange.getCell(highestRow, highestCol);
    //                sourceRange.worksheet.getUsedRange().format.fill.clear();
    //                sourceRange.worksheet.getUsedRange().format.font.bold = false;

    //                // 突出显示该单元格
    //                cellToHighlight.format.fill.color = "orange";
    //                cellToHighlight.format.font.bold = true;
    //            })
    //            .then(ctx.sync);
    //    })
    //    .catch(errorHandler);
    //}

    //function displaySelectedCells() {
    //    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
    //        function (result) {
    //            if (result.status === Office.AsyncResultStatus.Succeeded) {
    //                showNotification('选定的文本为:', '"' + result.value + '"');
    //            } else {
    //                showNotification('错误', result.error.message);
    //            }
    //        });
    //}

    // 处理错误的帮助程序函数
    function errorHandler(error) {
        // 请务必捕获 Excel.run 执行过程中出现的所有累积错误
        showNotification("错误", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // 用于显示通知的帮助程序函数
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();

function Country(name, value, mapColumn, increasement, color) {

    this.name = name;
    this.value = value;
    this.mapColumn = mapColumn;
    this.increasement = increasement;
    this.color = color;

    this.setValue = function (value) {
        this.value = value;
    }

    this.setIncreasement = function (increasement) {
        this.increasement = increasement;
    }

    this.setColor = function (color) {
        this.color = color;
    }

    this.updateIncrease = function () {
        this.value = this.value + this.increasement;
    }
}