sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageToast",
    "sap/ui/core/format/DateFormat",
    "sap/ui/core/Fragment",
    "sap/ui/model/FilterOperator",
    "sap/m/MessageBox"
], function (Controller, JSONModel, MessageToast, DateFormat, Fragment, FilterOperator, MessageBox) {
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
                var adjustedStart = that.adjustToNextMonday(startDate).getTime();
                var adjustedEnd = that.adjustToNextSunday(endDate).getTime();
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
            var levels = ["W", "M", "Q"];
            for (var l = 0; l < levels.length; l++) {
                var level = levels[l];
                var levelRecords = [];
                for (var i = 0; i < newRecords.length; i++) {
                    if (newRecords[i].Level === level) {
                        levelRecords.push(newRecords[i]);
                    }
                }
                levelRecords.sort(function (a, b) {
                    return a.StartDate - b.StartDate;
                });
                for (var j = 1; j < levelRecords.length; j++) {
                    var prevEnd = new Date(levelRecords[j - 1].EndDate);
                    var expectedNextStart = that.addDaysToDate(prevEnd, 1);
                    var actualStart = new Date(levelRecords[j].StartDate);
                    if (!that.datesAreEqual(expectedNextStart, actualStart)) {
                        MessageToast.show("Continuity error Please fix the upload file.");
                        that.allFilesData = [];
                        var oFileUploader = sap.ui.getCore().byId("myFileUploader");
                        oFileUploader.clear();
                        that.upload.close();
                        return;
                    }
                }
            }
            if (combinedData.length > 0 && newRecords.length > 0) {
                var lastEndDate = new Date(combinedData[combinedData.length - 1].EndDate);
                var expectedStartDate = that.addDaysToDate(lastEndDate, 1);
                var firstNewStartDate = new Date(newRecords[0].StartDate);
                if (!that.datesAreEqual(firstNewStartDate, expectedStartDate)) {
                    MessageToast.show("Continuity error: Excel upload failed, upload correct file");
                    that.allFilesData = [];
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
        adjustToNextMonday: function (date) {
            var result = new Date(date);
            var day = result.getDay();
            var daysToAdd = (day === 1) ? 0 : (8 - day) % 7;
            result.setDate(result.getDate() + daysToAdd);
            return result;
        },
        adjustToNextSunday: function (date) {
            var result = new Date(date);
            var day = result.getDay();
            var daysToAdd = (day === 0) ? 0 : (7 - day) % 7;
            result.setDate(result.getDate() + daysToAdd);
            return result;
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
            var oFormatter = DateFormat.getDateInstance({ pattern: "yyyy-MM-dd" });
            return oFormatter.format(jsDate);
        },
        generatePeriodDesc: function (level, endOfWeekDate) {
            var year = endOfWeekDate.getFullYear();
            var month = endOfWeekDate.getMonth();
            var fiscalYear = (month >= 2) ? year + 1 : year;
            if (level === "W") {
                var oFormatter = DateFormat.getDateInstance({ pattern: "yyyy/MM/dd" });
                return "CW " + oFormatter.format(endOfWeekDate);
            }
            var yearShort = fiscalYear.toString().slice(-2);
            var fiscalMonth = (month >= 2) ? month - 1 : month + 11;
            if (level === "M") {
                return "FY" + yearShort + " P" + fiscalMonth.toString().padStart(2, "0");
            } else if (level === "Q") {
                return "FY" + yearShort + " Q" + (Math.floor((fiscalMonth - 1) / 3) + 1);
            } else {
                return "";
            }
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
        },
    });
}); 
