sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageToast",
    "sap/ui/core/format/DateFormat",
    "sap/ui/core/Fragment",
    "sap/ui/model/FilterOperator",
    "sap/m/MessageBox"
], function (Controller, JSONModel, MessageToast, DateFormat, Fragment, FilterOperator,MessageBox) {
    "use strict";
    var that;
    return Controller.extend("application.controller.Calendar", {
        onInit: function () {
            that = this;
            that.allFilesData = [];
            that.unsavedChanges = false;
            var script = document.createElement('script');
            script.setAttribute('src', 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.2/xlsx.full.min.js');
            document.head.appendChild(script);
            that.getView().setModel(new JSONModel({ items: [] }), "weeklyModel");
            that.getView().setModel(new JSONModel({ items: [] }), "monthlyModel");
            that.getView().setModel(new JSONModel({ items: [] }), "quarterlyModel");
            that.getView().setModel(new JSONModel({ items: [] }), "activeModel");
        },
        handleUpload: function () {
            if (!that.upload) {
                that.upload = sap.ui.xmlfragment("application.Fragments.upload", that);
                that.getView().addDependent(that.upload);
            }
            that.upload.open();
        },
        onFileChange: function (oEvent) {
            var file = oEvent.getParameter("files")[0];
            if (!file) return;
            that.allFilesData = [];
            var reader = new FileReader();
            reader.onload = function (e) {
                var data = e.target.result;
                var workbook = XLSX.read(data, { type: "array" });
                var sheetName = workbook.SheetNames[0];
                var worksheet = workbook.Sheets[sheetName];
                var jsonData = XLSX.utils.sheet_to_json(worksheet);
                that.allFilesData = that.allFilesData.concat(jsonData);
            }.bind(that);
            reader.readAsArrayBuffer(file);
        },
        ExcelUpload: function () {
            var oData = that.allFilesData;
            var oFileUploader = sap.ui.getCore().byId("myFileUploader");
            var newRecords = [];
            var combinedData = [].concat(
                that.getView().getModel("weeklyModel").getData().items,
                that.getView().getModel("monthlyModel").getData().items,
                that.getView().getModel("quarterlyModel").getData().items
            );
            combinedData.sort(function (a, b) {
                return a.StartDate - b.StartDate;
            });
            for (var i = 0; i < oData.length; i++) {
                var entry = oData[i];
                var level = entry["LEVEL"];
                var startDate = that.parseExcelDate(entry["PERIODSTART"]);
                var endDate = that.parseExcelDate(entry["PERIODEND"]);
                var adjustedStart, adjustedEnd;
                if (level === "W") {
                    var adjusted = that.adjustToCustomFiscalWeek(startDate);
                    adjustedStart = adjusted.startOfWeek.getTime();
                    adjustedEnd = adjusted.endOfWeek.getTime();
                } else {
                    adjustedStart = that.adjustToNextMondayOrSame(startDate).getTime();
                    adjustedEnd = that.adjustToNextSundayOrSame(endDate).getTime();
                }
                var isDuplicate = combinedData.some(function (item) {
                    return item.StartDate === adjustedStart && item.EndDate === adjustedEnd;
                });
                if (!isDuplicate) {
                    newRecords.push({
                        Level: level,
                        StartDate: adjustedStart,
                        EndDate: adjustedEnd,
                        PeriodDesc: that.generatePeriodDesc(level, new Date(adjustedEnd)),
                        WeakWeight: entry["WEEKWEIGHT"],
                        MonthWeight: entry["MONTHWEIGHT"]
                    });
                }
            }
            newRecords.sort(function (a, b) {
                return a.StartDate - b.StartDate;
            });
            if (combinedData.length > 0 && newRecords.length > 0) {
                var lastEndDate = new Date(combinedData[combinedData.length - 1].EndDate);
                var expectedStartDate = that.addDaysToDate(lastEndDate, 1);
                var firstNewStartDate = new Date(newRecords[0].StartDate);
                if (!that.datesAreEqual(firstNewStartDate, expectedStartDate)) {
                    MessageToast.show("Continuity error: New data must start immediately after last existing period ends (Monday to Sunday week).");
                    oFileUploader.clear();
                    that.upload.close();
                    return;
                }
            }
            var allData = combinedData.concat(newRecords);
            var weeklyData = allData.filter(item => item.Level === "W");
            var monthlyData = allData.filter(item => item.Level === "M");
            var quarterlyData = allData.filter(item => item.Level === "Q");
            that.getView().setModel(new JSONModel({ items: weeklyData }), "weeklyModel");
            that.getView().setModel(new JSONModel({ items: monthlyData }), "monthlyModel");
            that.getView().setModel(new JSONModel({ items: quarterlyData }), "quarterlyModel");
            var selectedKey = that.getView().byId("iconTabBar").getSelectedKey();
            that.switchActiveModel(selectedKey);
            var oTable = that.getView().byId("calendarTable");
            var oBinding = oTable.getBinding("items");
            if (oBinding) {
                var oSorter = new sap.ui.model.Sorter("Level", false);
                oSorter.fnCompare = function (a, b) {
                    var order = ["W", "M", "Q"];
                    return order.indexOf(a) - order.indexOf(b);
                };
                oBinding.sort([oSorter]);
            }
        
            if (newRecords.length > 0) {
                MessageToast.show("New records uploaded successfully.");
            } else {
                MessageToast.show("Duplicates found. Excel upload skipped.");
            }
        
            that.upload.close();
            oFileUploader.clear();
        },
        adjustToNextMondayOrSame: function(date) {
            var day = date.getDay();
            if (day === 1) return new Date(date); 
            var diff = (8 - day) % 7;
            var adjusted = new Date(date);
            adjusted.setDate(adjusted.getDate() + diff);
            return adjusted;
        },
        adjustToNextSundayOrSame: function(date) {
            var day = date.getDay();
            if (day === 0) return new Date(date); 
            var diff = 7 - day;
            var adjusted = new Date(date);
            adjusted.setDate(adjusted.getDate() + diff);
            return adjusted;
        },
        adjustToCustomFiscalWeek: function (date) {
            var fiscalStart = new Date("2024-02-26");
            var msInDay = 1000 * 60 * 60 * 24;
            var daysSinceStart = Math.floor((date - fiscalStart) / msInDay);
            var weekOffset = Math.floor(daysSinceStart / 7);
            var startOfWeek = new Date(fiscalStart);
            startOfWeek.setDate(fiscalStart.getDate() + weekOffset * 7);
            var endOfWeek = new Date(startOfWeek);
            endOfWeek.setDate(startOfWeek.getDate() + 6);
            return {
                startOfWeek,
                endOfWeek
            };
        },
        adjustToWeekStartEnd: function (date) {
            var dayOfWeek = date.getDay(); 
            var diffToMonday;
            if (dayOfWeek === 0) {
                diffToMonday = -6; 
            } else {
                diffToMonday = 1 - dayOfWeek;
            }
            var startOfWeek = new Date(date);
            startOfWeek.setDate(startOfWeek.getDate() + diffToMonday);
            var endOfWeek = new Date(startOfWeek);
            endOfWeek.setDate(startOfWeek.getDate() + 6);
            return {
                 startOfWeek,
                 endOfWeek
            };
        }, 
        generatePeriodDesc: function (level, endOfWeekDate) {
            var year = endOfWeekDate.getFullYear();
            var month = endOfWeekDate.getMonth(); 
            var fiscalYear;
            if (month >= 2) { 
                fiscalYear = year + 1;
            } else {
                fiscalYear = year;
            }
            if (level === "W") {
                var oFormatter = DateFormat.getDateInstance({ pattern: "yyyy/MM/dd" });
                return "CW " + oFormatter.format(endOfWeekDate);
            }
            var yearShort = fiscalYear.toString().slice(-2);
            if (level === "M") {
                var fiscalMonth;
                if (month >= 2) {
                    fiscalMonth = month - 1;
                } else {
                    fiscalMonth = month + 11;
                }
                var monthStr = fiscalMonth.toString().padStart(2, "0");
                return "FY" + yearShort + " P" + monthStr;
            } else if (level === "Q") {
                var fiscalMonth;
                if (month >= 2) {
                    fiscalMonth = month - 1;
                } else {
                    fiscalMonth = month + 11;
                }
                var fiscalQuarter = Math.floor((fiscalMonth - 1) / 3) + 1;
                return "FY" + yearShort + " Q" + fiscalQuarter;
            }
        },
        switchActiveModel: function (key) {
            var modelName = "";
            if (key === "W") {
                modelName = "weeklyModel";  
            } else if (key === "M") {
                modelName = "monthlyModel";
            } else if (key === "Q") {
                modelName = "quarterlyModel";
            }
            var selectedModel = that.getView().getModel(modelName);
            that.getView().setModel(selectedModel, "activeModel");
        },
        adjustToWeekStartEnd: function (date) {
            var dayOfWeek = date.getDay(); 
            var diffToMonday;
            if (dayOfWeek === 0) {
                diffToMonday = -6; 
            } else {
                diffToMonday = 1 - dayOfWeek;
            }
            var startOfWeek = new Date(date);
            startOfWeek.setDate(startOfWeek.getDate() + diffToMonday);
            var endOfWeek = new Date(startOfWeek);
            endOfWeek.setDate(startOfWeek.getDate() + 6);
            return {
                 startOfWeek,
                 endOfWeek
            };
        },
        parseExcelDate: function (value) {
            return new Date((value - 25569) * 86400 * 1000);
        },
        addDaysToDate: function (date, days) {
            var result = new Date(date);
            result.setDate(result.getDate() + days);
            return result;
        },
        datesAreEqual: function (d1, d2) {
            return d1.getFullYear() === d2.getFullYear() &&
                   d1.getMonth() === d2.getMonth() &&
                   d1.getDate() === d2.getDate();
        },
        formatDate: function (timestamp) {
            var jsDate = new Date(timestamp);
            if (isNaN(jsDate.getTime())) {
                return "";
            }
            var oFormatter = DateFormat.getDateInstance({ pattern: "yyyy-MM-dd" });
            return oFormatter.format(jsDate);
        },
        onTabSelect: function (oEvent) {
            var key = oEvent.getParameter("key");
            if (that.unsavedChanges) {
                sap.m.MessageBox.warning("There are unsaved changes. Do you want continue?", {
                    actions: [sap.m.MessageBox.Action.YES, sap.m.MessageBox.Action.NO],
                    onClose: function (action) {
                        if (action === sap.m.MessageBox.Action.YES) {
                            that.unsavedChanges = false; 
                            that.switchActiveModel(key);
                        } else {
                            var iconTabBar = that.getView().byId("iconTabBar");
                            var currentKey = iconTabBar.getSelectedKey();
                            iconTabBar.setSelectedKey(currentKey); 
                        }
                    }
                });
            } else {
                that.switchActiveModel(key);
            }
        },
        onTabSelect: function (oEvent) {
            var key = oEvent.getParameter("key");
            if (that.unsavedChanges) {
                sap.m.MessageBox.warning("There are unsaved changes. Do you want to continue?", {
                    actions: [sap.m.MessageBox.Action.YES, sap.m.MessageBox.Action.NO],
                    onClose: function (action) {
                        if (action === sap.m.MessageBox.Action.YES) {
                            that.unsavedChanges = false;
                            that.switchActiveModel(key); 
                        } else {
                            var iconTabBar = that.byId("iconTabBar");
                            var currentKey = iconTabBar.getSelectedKey();
                            iconTabBar.setSelectedKey(currentKey); 
                        }
                    }
                });
            } else {
                that.switchActiveModel(key); 
            }
        },
        onInputChange: function (oEvent) {
            var newDesc = oEvent.getSource().getValue();
            var oInput = oEvent.getSource();
            var oModel = that.getView().getModel("activeModel");
            var data = oModel.getData().items;
            var context = oInput.getBindingContext("activeModel");
            var oldDesc = context.getObject().PeriodDesc;
            var validPattern = /^[a-zA-Z0-9/ ]*$/;
            var isValidFormat = validPattern.test(newDesc);
            var descriptionExists = false;
            data.forEach(function (item) {
                if (item.PeriodDesc === newDesc && item.PeriodDesc !== oldDesc) {
                    descriptionExists = true;
                }
            });
            if (!isValidFormat && descriptionExists) {
                oInput.setValueState("Error");
                oInput.setValueStateText("Special characters are not allowed. This description already exists.");
            } else if (!isValidFormat) {
                oInput.setValueState("Error");
                oInput.setValueStateText("Special characters are not allowed.");
            } else if (descriptionExists) {
                oInput.setValueState("Error");
                oInput.setValueStateText("This description already exists.");
            } else {
                oInput.setValueState("None");
                if (newDesc !== oldDesc) {
                    that.unsavedChanges = true;
                } else {
                    that.unsavedChanges = false;
                }
            }
        },
        close: function () {
            var oFileUploader = sap.ui.getCore().byId("myFileUploader");
            that.upload.close();
            oFileUploader.clear();
        }
    });
});
