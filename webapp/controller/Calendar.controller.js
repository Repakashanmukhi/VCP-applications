sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageToast",
    "sap/ui/core/format/DateFormat",
    "sap/ui/core/Fragment",
     "sap/ui/model/FilterOperator"
], function (Controller, JSONModel, MessageToast, DateFormat, Fragment, FilterOperator) {
    "use strict";
     var that;
     return Controller.extend("application.controller.Calendar", {
        onInit: function () {
            that = this;
            that.allFilesData = [];
            var script = document.createElement('script');
            script.setAttribute('src', 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.2/xlsx.full.min.js');
            document.head.appendChild(script);
            var oModel = new JSONModel({ 
                items: [] 
            });
            that.getView().setModel(oModel, "oNewModel");
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
            var data = that.getView().getModel("oNewModel").getData().items;
            for (var i = 0; i < oData.length; i++) {
                var entry = oData[i];
                var startRaw = entry["PERIODSTART"];
                var endRaw = entry["PERIODEND"];
                var cleanStart = Math.floor(startRaw);
                var cleanEnd = Math.floor(endRaw);
                var isDuplicate = false;
                for (var j = 0; j < data.length; j++) {
                    if (data[j].StartDate === cleanStart && data[j].EndDate === cleanEnd) {
                        isDuplicate = true;
                        break;
                    }
                }
                if (!isDuplicate) {
                    newRecords.push({
                        Level: entry["LEVEL"],
                        StartDate: cleanStart,
                        EndDate: cleanEnd,
                        PeriodDesc: entry["PERIODDESC"],
                        WeakWeight: entry["WEEKWEIGHT"],
                        MonthWeight: entry["MONTHWEIGHT"],
                    });
                }
            }
            newRecords.sort(function (a, b) {
                var dateA = that.parseExcelDate(a.StartDate);
                var dateB = that.parseExcelDate(b.StartDate);
                return dateA - dateB;
            });
            if (data.length > 0 && newRecords.length > 0) {
                var lastEndDateStr = data[data.length - 1].EndDate;
                var lastEndDate = that.parseExcelDate(lastEndDateStr);
                var expectedStartDate = that.addDaysToDate(lastEndDate, 1);          
                var firstNewStartDate = that.parseExcelDate(newRecords[0].StartDate);
                if (!that.datesAreEqual(firstNewStartDate, expectedStartDate)) {
                    MessageToast.show("Please check your Excel file. Continuity of data is missing.");
                    oFileUploader.clear();
                    that.upload.close();
                    return;
                }
            }
            var finalData = data.concat(newRecords);
            var oModel = new sap.ui.model.json.JSONModel({ items: finalData });
            that.getView().setModel(oModel, "oNewModel");
            var oTabBar = that.getView().byId("iconTabBar");
            var selectedKey = oTabBar.getSelectedKey();
            that.onTabSelect({ 
                getParameter: function () 
                { 
                    return selectedKey; 
                } 
            });
            var oTable = that.getView().byId("data");
            var oBinding = oTable.getBinding("items");
            var oSorter = new sap.ui.model.Sorter("Level", false);
            oSorter.fnCompare = function (a, b) {
                var order = ["W", "M", "Q"];
                return order.indexOf(a) - order.indexOf(b);
            };
            oBinding.sort([oSorter]);
            if (newRecords.length > 0) {
                MessageToast.show("New records uploaded successfully.");
            } else {
                MessageToast.show("Duplicates found Excel upload failed.");
            }
            that.upload.close();
        
            oFileUploader.clear();
        }, 
        parseExcelDate: function (value) {
            return new Date((value - 25569) * 86400 * 1000);
        },
        addDaysToDate: function (date, days) {
            let result = new Date(date);
            result.setDate(result.getDate() + days);
            return result;
        },
        datesAreEqual: function (d1, d2) {
            return d1.getFullYear() === d2.getFullYear() &&
                   d1.getMonth() === d2.getMonth() &&
                   d1.getDate() === d2.getDate();
        }, 
        formatDate: function (excelDate) {
            var jsDate = new Date((excelDate - 25569) * 86400 * 1000);
            var oFormatter = DateFormat.getDateInstance({ pattern: "yyyy-MM-dd" });
            return oFormatter.format(jsDate);
        },
        onTabSelect: function (oEvent) {
            var key = oEvent.getParameter("key");
            var oTable = that.getView().byId("data");
            var oBinding = oTable.getBinding("items");
            var oFilter = new sap.ui.model.Filter("Level",FilterOperator.EQ, key);
            oBinding.filter([oFilter]);  
        },
        onInputChange: function (oEvent) {
            var newDesc = oEvent.getSource().getValue();  
            var oInput = oEvent.getSource();  
            var data = that.getView().getModel("oNewModel").getData().items;  
            var context = oInput.getBindingContext("oNewModel");
            var oldDesc = context.getObject().PeriodDesc;
            var descriptionExists = false;
            data.forEach(function(item){
                if( item.PeriodDesc === newDesc && item.PeriodDesc !== oldDesc ){
                    descriptionExists = true
                }
            })
            if (descriptionExists) {
                oInput.setValueState("Error");
                oInput.setValueStateText("This description already exists.");
            } else {
                oInput.setValueState("None");
            }
        },
        close: function () {
            that.upload.close();
        }
    });
});


