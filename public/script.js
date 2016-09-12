'use strict';

var ExcelParser = angular.module('ExcelParser', ['zingchart-angularjs']);

ExcelParser.controller('MainCtrl', ['$scope', 'ExcelParserService',
    function ($scope, Service) {
        $scope.fileSelected = null;
        $scope.sheet_name_list = [];
        $scope.houseDetails = [];
        $scope.houseConsSummer = [];
        $scope.houseConsWinter = [];
        $scope.houseApplianceSummer = [];
        $scope.houseApplianceWinter = [];
        $scope.currentCycle = 'summer';
        $scope.workbook = null;
        $scope.houseDetailsFields = Service.houseDetailsFields;
        $scope.globalDetailsFields = Service.globalDetailsFields;
        $scope.sampleDetailsFields = Service.sampleDetailsFields;

        $scope.loadFile = function () {

            /* set up XMLHttpRequest */
            var url = "HHM-V6.xlsx";
            var oReq = new XMLHttpRequest();
            oReq.open("GET", url, true);
            oReq.responseType = "arraybuffer";

            oReq.onload = function (e) {
                var arraybuffer = oReq.response;

                /* convert data to binary string */
                var data = new Uint8Array(arraybuffer);
                var arr = new Array();
                for (var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
                var bstr = arr.join("");

                /* Call XLSX */
                $scope.workbook = XLSX.read(bstr, {type: "binary"});
                //console.log(workbook);

                $scope.sheet_name_list = $scope.workbook.SheetNames;
                $scope.welcome = "Hello Again";
                $scope.init();
                Service.safeApply($scope);

                //sheet_name_list.forEach(function(y) { /* iterate through sheets */
                //    var worksheet = workbook.Sheets[y];
                //    for (var z in worksheet) {
                //        /* all keys that do not begin with "!" correspond to cell addresses */
                //        if(z[0] === '!') continue;
                //        console.log(y + "!" + z + "=" + JSON.stringify(worksheet[z].v));
                //    }
                //});

                /* DO SOMETHING WITH workbook HERE */
            };
            oReq.send();
        };

        $scope.changeCycle = function (cycle) {
            switch (cycle) {
                case 'summer':
                    $scope.currentCycle = cycle;
                    $scope.currentApplianceList = $scope.houseApplianceSummer[$scope.selectedSample];
                    break;
                case 'winter':
                    $scope.currentCycle = cycle;
                    $scope.currentApplianceList = $scope.houseApplianceWinter[$scope.selectedSample];
                    break;
                default:
                    console.log("Invalid cycle.");
            }
        };

        $scope.init = function () {
            $scope.houseDetails = XLSX.utils.sheet_to_json($scope.workbook.Sheets["House Details"]);
            console.log("There are " + $scope.houseDetails.length + " samples.");
        };

        $scope.loadHouseDetails = function () {
            $scope.currentHouseDetails = $scope.houseDetails[$scope.selectedSample];
            $scope.globalStats();
            $scope.currentCycle = 'summer';
            $scope.currentApplianceList = $scope.houseApplianceSummer[$scope.selectedSample];
            $scope.populateDataForDeviceChart();
            $scope.sampleDetailsFields.totalSummer.value = $scope.houseConsSummer[$scope.selectedSample];
            $scope.sampleDetailsFields.avgSummer.value = $scope.houseConsSummer[$scope.selectedSample] / $scope.houseApplianceSummer[$scope.selectedSample].length;
            $scope.sampleDetailsFields.totalWinter.value = $scope.houseConsWinter[$scope.selectedSample];
            $scope.sampleDetailsFields.avgWinter.value = $scope.houseConsWinter[$scope.selectedSample] / $scope.houseApplianceWinter[$scope.selectedSample].length;
            $scope.sampleDetailsFields.totalCombined.value = $scope.sampleDetailsFields.totalSummer.value + $scope.sampleDetailsFields.totalWinter.value;
            $scope.sampleDetailsFields.avgCombined.value = $scope.sampleDetailsFields.totalCombined.value / ($scope.houseApplianceSummer[$scope.selectedSample].length + $scope.houseApplianceWinter[$scope.selectedSample].length);
        };

        $scope.globalStats = function () {
            $scope.calcSummer();
            $scope.calcWinter();
            $scope.globalDetailsFields.totalSummer.value = $scope.houseConsSummer[0];
            $scope.globalDetailsFields.totalWinter.value = $scope.houseConsWinter[0];

            $scope.globalDetailsFields.maxHouseId.value = 0;
            $scope.globalDetailsFields.maxHouseTotal.value = $scope.houseConsSummer[0] + $scope.houseConsWinter[0];
            $scope.globalDetailsFields.minHouseId.value = 0;
            $scope.globalDetailsFields.minHouseTotal.value = $scope.houseConsSummer[0] + $scope.houseConsWinter[0];
            for (var i = 1; i < $scope.houseDetails.length; i++) {
                var temp = $scope.houseConsSummer[i] + $scope.houseConsWinter[i];
                $scope.globalDetailsFields.totalSummer.value += $scope.houseConsSummer[i];
                $scope.globalDetailsFields.totalWinter.value += $scope.houseConsWinter[i];
                if (temp < $scope.globalDetailsFields.minHouseTotal.value) {
                    $scope.globalDetailsFields.minHouseId.value = i;
                    $scope.globalDetailsFields.minHouseTotal.value = temp;
                }
                if (temp > $scope.globalDetailsFields.maxHouseTotal.value) {
                    $scope.globalDetailsFields.maxHouseId.value = i;
                    $scope.globalDetailsFields.maxHouseTotal.value = temp;
                }
            }
        };

        $scope.calcSummer = function () {
            var consSummerTop = 3, avgPowerColPrefix = 'D', devNameColPrefix = 'A';
            $scope.summerDeviceAvg = {};

            for (var i = consSummerTop; ; i++) {
                var avgPower = $scope.workbook.Sheets["Consumption summer"][avgPowerColPrefix + i];
                var devName = $scope.workbook.Sheets["Consumption summer"][devNameColPrefix + i];
                if (devName === undefined && avgPower === undefined) break;
                devName = devName === undefined ? "NA" : devName.v;
                avgPower = avgPower === undefined ? 0 : avgPower.w;
                var tempObj = {};
                tempObj[devName] = avgPower;
                angular.extend($scope.summerDeviceAvg, tempObj);
            }

            // Appliance Ownership data not used
            var sheetName = "HHS appliances  Time of use S";
            var summerRows = XLSX.utils.sheet_to_json($scope.workbook.Sheets[sheetName]);
            for (var i = 0; i < $scope.houseDetails.length; i++) {
                //$scope.houseConsSummer[i]
                var currentRow = summerRows[i];
                var sumTotal = 0;
                $scope.houseApplianceSummer[i] = [];
                for (var appliance in currentRow) {
                    if (currentRow[appliance] > 0 && $scope.summerDeviceAvg[appliance] !== undefined) {
                        sumTotal += currentRow[appliance] * $scope.summerDeviceAvg[appliance];
                        var tempObj = {};
                        tempObj[appliance] = currentRow[appliance];
                        $scope.houseApplianceSummer[i].push(tempObj);
                    }
                }
                $scope.houseConsSummer[i] = sumTotal;
            }
            //console.log(summerRows);
        };

        $scope.calcWinter = function () {
            var timeUseSummerTop = 1, timeUseSummerLeft = 2;
            var consWinterTop = 3, avgPowerColPrefix = 'D', devNameColPrefix = 'A';
            $scope.winterDeviceAvg = {};

            for (var i = consWinterTop; ; i++) {
                var avgPower = $scope.workbook.Sheets["Consumption winter"][avgPowerColPrefix + i];
                var devName = $scope.workbook.Sheets["Consumption winter"][devNameColPrefix + i];
                if (devName === undefined && avgPower === undefined) break;
                devName = devName === undefined ? "NA" : devName.v;
                avgPower = avgPower === undefined ? 0 : avgPower.w;
                var tempObj = {};
                tempObj[devName] = avgPower;
                angular.extend($scope.winterDeviceAvg, tempObj);
            }

            var sheetName = "HHS appliances time of use W";
            var winterRows = XLSX.utils.sheet_to_json($scope.workbook.Sheets[sheetName]);
            for (var i = 0; i < $scope.houseDetails.length; i++) {
                //$scope.houseConsSummer[i]
                var currentRow = winterRows[i];
                var sumTotal = 0;
                $scope.houseApplianceWinter[i] = [];
                for (var appliance in currentRow) {
                    if (currentRow[appliance] > 0 && $scope.winterDeviceAvg[appliance] !== undefined) {
                        sumTotal += currentRow[appliance] * $scope.winterDeviceAvg[appliance];
                        var tempObj = {};
                        tempObj[appliance] = currentRow[appliance];
                        $scope.houseApplianceWinter[i].push(tempObj);
                    }
                }
                $scope.houseConsWinter[i] = sumTotal;
            }
            //console.log(winterRows);
        };

        $scope.populateDataForDeviceChart = function () {
            $scope.deviceChartData = {};
            $scope.deviceChartData.type = 'pie3d';
            $scope.deviceChartData.title = {'text': 'Appliance Usage'};
            $scope.deviceChartData.plot = {
                "offset-r": "25%" //provide percentage value
            };
            //$scope.deviceChartData.legend = {};
            $scope.deviceChartData.series = [];
            for (var i in $scope.currentApplianceList) {
                for (var key in $scope.currentApplianceList[i]) {
                    var newObj = {"values": $scope.currentApplianceList[i][key], "text": key, "legend-text": key};
                    $scope.deviceChartData.series.push(newObj);
                }
            }
            console.log($scope.deviceChartData);
        };

        $scope.populateDataForDeviceHourlyCons = function () {

        };

        $scope.populateDataForDeviceMinuteCons = function () {

        };

    }
]);

ExcelParser.service('ExcelParserService', ['$http',
    function ($http) {

        this.safeApply = function ($scope, fn) {
            var phase = $scope.$root.$$phase;
            if (phase == '$apply' || phase == '$digest') {
                if (fn && (typeof(fn) === 'function')) {
                    fn();
                }
            } else {
                $scope.$apply(fn);
            }
        };

        this.houseDetailsFields = [
            {
                'id': 'Gender',
                'display': 'Gender'
            },
            {
                'id': 'housesize',
                'display': 'House Size'
            },
            {
                'id': 'housetype',
                'display': 'House Type'
            },
            {
                'id': 'Noppl',
                'display': 'No. of People'
            },
            {
                'id': 'employment',
                'display': 'Employment Status'
            },
            {
                'id': 'age',
                'display': 'Family Age Group'
            },
            {
                'id': 'statementsapply',
                'display': 'Statement Apply to You'
            }
        ];

        this.globalDetailsFields = {};
        this.globalDetailsFields.maxHouseId = {
            'display': 'Max House Sample No.',
            'value': 0
        };
        this.globalDetailsFields.maxHouseTotal = {
            'display': 'Max Consumption(KwH)',
            'value': 0
        };
        this.globalDetailsFields.minHouseId = {
            'display': 'Min House Sample No.',
            'value': 0
        };
        this.globalDetailsFields.minHouseTotal = {
            'display': 'Min Consumption(KwH)',
            'value': 0
        };
        this.globalDetailsFields.totalSummer = {
            'display': 'Total Consumption Summer(KwH)',
            'value': 0
        };
        this.globalDetailsFields.totalWinter = {
            'display': 'Total Consumption Winter(KwH)',
            'value': 0
        };

        this.sampleDetailsFields = {};
        this.sampleDetailsFields.totalSummer = {
            'display': 'Total Electricity Consumption(S)',
            'value': 0
        };
        this.sampleDetailsFields.totalWinter = {
            'display': 'Total Electricity Consumption(W)',
            'value': 0
        };
        this.sampleDetailsFields.totalCombined = {
            'display': 'Total Electricity Consumption(Combined)',
            'value': 0
        };
        this.sampleDetailsFields.avgSummer = {
            'display': 'Average Electricity Consumption(S)',
            'value': 0
        };
        this.sampleDetailsFields.avgWinter = {
            'display': 'Average Electricity Consumption(W)',
            'value': 0
        };
        this.sampleDetailsFields.avgCombined = {
            'display': 'Average Electricity Consumption(Combined)',
            'value': 0
        };

    }
]);

ExcelParser.filter('to_trusted', ['$sce',
    function ($sce) {
        return function (text) {
            return $sce.trustAsHtml(text);
        }
    }
]);