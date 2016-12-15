'use strict';

var ExcelParser = angular.module('ExcelParser', ['nvd3']);

function capitalize(str) {
    return str.replace(/\w\S*/g, function (txt) {
        return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
    });
}

ExcelParser.controller('MainCtrl', ['$scope', 'ExcelParserService',
    function ($scope, Service) {
        $scope.selectedSample = undefined;
        $scope.currentCycle = 'summer';
        $scope.houseDetails = [];

        $scope.ownershipSummer = [];
        $scope.timeOfUseSummer = [];
        $scope.powerUsedSummer = [];
        $scope.avgDevicePowerSummer = [];

        $scope.ownershipWinter = [];
        $scope.timeOfUseWinter = [];
        $scope.powerUsedWinter = [];
        $scope.avgDevicePowerWinter = [];

        $scope.optionalDevices = [];
        $scope.optionalAvgDevicePowerSummer = [];
        $scope.optionalAvgDevicePowerWinter = [];

        $scope.currentApplianceList = [];
        $scope.deviceChartData = {};

        $scope.houseDetailsFields = Service.houseDetailsFields;
        $scope.globalDetailsFields = Service.globalDetailsFields;
        $scope.sampleDetailsFields = Service.sampleDetailsFields;

        $scope.loadFile = function () {

            /* set up XMLHttpRequest */
            var url = "HHM-revised.xlsx";
            var oReq = new XMLHttpRequest();
            oReq.open("GET", url, true);
            oReq.responseType = "arraybuffer";

            oReq.onload = function (e) {
                console.debug(e);
                var arraybuffer = oReq.response;

                /* convert data to binary string */
                var data = new Uint8Array(arraybuffer);
                var arr = [];
                for (var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
                var bstr = arr.join("");

                /* Call XLSX */
                $scope.workbook = XLSX.read(bstr, {type: "binary"});
                $scope.init();
                Service.safeApply($scope);

                //$scope.sheet_name_list = $scope.workbook.SheetNames;
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
                    $scope.currentApplianceList = [];
                    $scope.currentCycle = cycle;
                    for (var appliance of Object.keys($scope.ownershipSummer[$scope.selectedSample])) {
                        if ($scope.ownershipSummer[$scope.selectedSample][appliance] == 1) {
                            var tempObj = {};
                            tempObj.name = appliance;
                            tempObj.hoursUsed = $scope.timeOfUseSummer[$scope.selectedSample][appliance];
                            tempObj.avgPowerForDevice = $scope.avgDevicePowerSummer[appliance];
                            $scope.currentApplianceList.push(tempObj);
                        }
                    }
                    break;
                case 'winter':
                    $scope.currentApplianceList = [];
                    $scope.currentCycle = cycle;
                    for (appliance of Object.keys($scope.ownershipSummer[$scope.selectedSample])) {
                        if ($scope.ownershipSummer[$scope.selectedSample][appliance] == 1) {
                            tempObj = {};
                            tempObj.name = appliance;
                            tempObj.hoursUsed = $scope.timeOfUseSummer[$scope.selectedSample][appliance];
                            tempObj.avgPowerForDevice = $scope.avgDevicePowerSummer[appliance];
                            $scope.currentApplianceList.push(tempObj);
                        }
                    }
                    break;
                default:
                    console.log("Invalid cycle.");
            }
            $scope.populateDataForChart();
        };

        $scope.init = function () {
            $scope.houseDetails = XLSX.utils.sheet_to_json($scope.workbook.Sheets["House Details"]);
            console.debug("There are " + $scope.houseDetails.length + " samples.");
            $scope.globalStats();
        };

        $scope.populateAvgPower = function (season) {
            var sheetName = ($scope.optional ? 'Optional ' : '') + 'Consumption ' + capitalize(season);
            var objHolder = ($scope.optional ? 'optionalAvgDevicePower' : 'avgDevicePower') + capitalize(season);
            $scope[objHolder] = {};
            var consumptionData = XLSX.utils.sheet_to_json($scope.workbook.Sheets[sheetName]);
            for (var row of consumptionData) {
                var tempObj = {};
                tempObj[row['Appliances']] = row['average Power rate KWh'];
                angular.extend($scope[objHolder], tempObj);
            }
        };

        $scope.calcSummer = function () {
            $scope.ownershipSummer = XLSX.utils.sheet_to_json($scope.workbook.Sheets["Appliance Ownership Summer"]);
            $scope.timeOfUseSummer = XLSX.utils.sheet_to_json($scope.workbook.Sheets["Time of Use Summer"]);
            $scope.populateAvgPower('summer');

            for (var i = 0; i < $scope.houseDetails.length; i++) {
                $scope.powerUsedSummer[i] = 0;
                for (var j of Object.keys($scope.ownershipSummer[i])) {
                    if ($scope.ownershipSummer[i][j] == 1) {
                        $scope.powerUsedSummer[i] += $scope.timeOfUseSummer[i][j] * ($scope.avgDevicePowerSummer[j] / 60);
                    }
                }
            }
        };

        $scope.calcWinter = function () {
            $scope.ownershipWinter = XLSX.utils.sheet_to_json($scope.workbook.Sheets["Appliance Ownership Winter"]);
            $scope.timeOfUseWinter = XLSX.utils.sheet_to_json($scope.workbook.Sheets["Time of Use Winter"]);
            $scope.populateAvgPower('winter');

            for (var i = 0; i < $scope.houseDetails.length; i++) {
                $scope.powerUsedWinter[i] = 0;
                for (var j of Object.keys($scope.ownershipWinter[i])) {
                    if ($scope.ownershipWinter[i][j] == 1) {
                        $scope.powerUsedWinter[i] += $scope.timeOfUseWinter[i][j] * ($scope.avgDevicePowerWinter[j] / 60);
                    }
                }
            }
        };

        $scope.globalStats = function () {
            $scope.calcSummer();
            $scope.calcWinter();

            $scope.globalDetailsFields.totalSummer.value = $scope.powerUsedSummer[0];
            $scope.globalDetailsFields.totalWinter.value = $scope.powerUsedWinter[0];
            $scope.globalDetailsFields.maxHouseId.value = 0;
            $scope.globalDetailsFields.maxHouseTotal.value = $scope.powerUsedSummer[0] + $scope.powerUsedWinter[0];
            $scope.globalDetailsFields.minHouseId.value = 0;
            $scope.globalDetailsFields.minHouseTotal.value = $scope.powerUsedSummer[0] + $scope.powerUsedWinter[0];

            for (var i = 1; i < $scope.houseDetails.length; i++) {
                var temp = $scope.powerUsedSummer[i] + $scope.powerUsedWinter[i];
                $scope.globalDetailsFields.totalSummer.value += $scope.powerUsedSummer[i];
                $scope.globalDetailsFields.totalWinter.value += $scope.powerUsedWinter[i];
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

        $scope.calcAvgPowerForHouse = function (cycle) {
            var deviceCount = 0;
            var totalPower = 0;
            switch (cycle) {
                case 'summer':
                    for (var i of Object.keys($scope.ownershipSummer[$scope.selectedSample])) {
                        if ($scope.ownershipSummer[$scope.selectedSample][i] == 1) deviceCount++;
                    }
                    totalPower = $scope.powerUsedSummer[$scope.selectedSample];
                    $scope.sampleDetailsFields.summerDeviceCount.value = deviceCount;
                    console.debug("summer device count: " + deviceCount);
                    break;
                case 'winter':
                    for (var j of Object.keys($scope.ownershipWinter[$scope.selectedSample])) {
                        if ($scope.ownershipWinter[$scope.selectedSample][j] == 1) deviceCount++;
                    }
                    totalPower = $scope.powerUsedWinter[$scope.selectedSample];
                    $scope.sampleDetailsFields.summerDeviceCount.value = deviceCount;
                    console.debug("winter device count: " + deviceCount);
                    break;
                case 'combined':
                    for (var k of Object.keys($scope.ownershipSummer[$scope.selectedSample])) {
                        if ($scope.ownershipSummer[$scope.selectedSample][k] == 1) deviceCount++;
                    }
                    for (var l of Object.keys($scope.ownershipWinter[$scope.selectedSample])) {
                        if ($scope.ownershipWinter[$scope.selectedSample][l] == 1) deviceCount++;
                    }
                    totalPower = $scope.powerUsedSummer[$scope.selectedSample] + $scope.powerUsedWinter[$scope.selectedSample];
                    console.debug("combined device count: " + deviceCount);
                    break;
                default:
                    console.error('Invalid cycle encountered.');
            }
            return (totalPower / deviceCount);
        };

        $scope.loadHouseDetails = function () {
            $scope.currentHouseDetails = $scope.houseDetails[$scope.selectedSample];
            $scope.changeCycle('summer');
            $scope.calculateSAPCoeff();
            $scope.sampleDetailsFields.totalSummer.value = $scope.powerUsedSummer[$scope.selectedSample];
            $scope.sampleDetailsFields.avgSummer.value = $scope.calcAvgPowerForHouse('summer');
            $scope.sampleDetailsFields.totalWinter.value = $scope.powerUsedWinter[$scope.selectedSample];
            $scope.sampleDetailsFields.avgWinter.value = $scope.calcAvgPowerForHouse('winter');
            $scope.sampleDetailsFields.totalCombined.value = $scope.sampleDetailsFields.totalSummer.value + $scope.sampleDetailsFields.totalWinter.value;
            $scope.sampleDetailsFields.avgCombined.value = $scope.calcAvgPowerForHouse('combined');
            $scope.optionalChanged();
            $scope.populateDataForChart();
            $scope.showHoursGraph();
        };

        $scope.calculateSAPCoeff = function () {
            $scope.currentHouseDetails.Ea = 0;
            $scope.currentHouseDetails.Eb = 0;
            $scope.currentHouseDetails.tfa = 0;
            var TFA = Service.TFAData($scope.currentHouseDetails.housesize, $scope.currentHouseDetails.Noppl);
            var n = parseInt($scope.currentHouseDetails.housesize);
            var coeff = Math.pow((TFA * n), 0.4714);
            $scope.currentHouseDetails.tfa = TFA;
            $scope.currentHouseDetails.Ea = 207.8 * coeff;
            $scope.currentHouseDetails.Eb = 59.73 * coeff;
            $scope.currentHouseDetails.Ea = $scope.currentHouseDetails.Ea.toFixed(5);
            $scope.currentHouseDetails.Eb = $scope.currentHouseDetails.Eb.toFixed(5);
        };

        $scope.populateDataForChart = function () {
            $scope.deviceChartData.hourSeries = [];
            $scope.deviceChartData.percentageSeries = [];
            for (var appliance of $scope.currentApplianceList) {
                var newObj1 = {"key": appliance.name, "y": appliance.hoursUsed};
                var newObj2 = {"key": appliance.name, "y": (appliance.hoursUsed * (appliance.avgPowerForDevice / 60))};
                $scope.deviceChartData.hourSeries.push(newObj1);
                $scope.deviceChartData.percentageSeries.push(newObj2);
            }
            console.debug($scope.deviceChartData);
            Service.safeApply($scope);
        };

        $scope.showPercentageGraph = function () {
            $scope.applianceGraph = 'KWH';
        };

        $scope.showHoursGraph = function () {
            $scope.applianceGraph = 'HRS';
        };

        $scope.xFunction = function () {
            return function (d) {
                return d.key;
            };
        };

        $scope.yFunction = function () {
            return function (d) {
                return d.y;
            };
        };

        $scope.options = {
            chart: {
                type: 'pieChart',
                x: $scope.xFunction(),
                y: $scope.yFunction(),
                showLegend: true,
                tooltips: true,
                showLabels: false
            }
        };

        $scope.insertIntoChart = function () {
            for (var appliance of $scope.optionalDevices) {
                var newObj1 = {"key": appliance.name, "y": appliance.hoursUsed};
                var newObj2 = {"key": appliance.name, "y": (appliance.hoursUsed * (appliance.avgPowerForDevice / 60))};
                $scope.deviceChartData.hourSeries.push(newObj1);
                $scope.deviceChartData.percentageSeries.push(newObj2);
            }
            console.debug($scope.deviceChartData);
        };

        $scope.modifySampleDataPoint = function (operation) {
            switch (operation) {
                case 'add':
                    $scope.sampleDetailsFields['total' + capitalize($scope.currentCycle)].value += $scope.optionalDeviceContribution;
                    $scope.sampleDetailsFields.totalCombined.value += $scope.optionalDeviceContribution;
                    $scope.sampleDetailsFields['avg' + capitalize($scope.currentCycle)].value = ($scope.sampleDetailsFields['total' + capitalize($scope.currentCycle)].value) / ($scope.currentApplianceList.length + $scope.optionalDevices.length);
                    $scope.sampleDetailsFields.avgCombined.value = $scope.sampleDetailsFields.totalCombined.value / ($scope.sampleDetailsFields.summerDeviceCount.value + $scope.sampleDetailsFields.winterDeviceCount.value + $scope.optionalDevices.length);
                    break;
                case 'deduct':
                    $scope.sampleDetailsFields['total' + capitalize($scope.currentCycle)].value -= $scope.optionalDeviceContribution;
                    $scope.sampleDetailsFields.totalCombined.value -= $scope.optionalDeviceContribution;
                    $scope.sampleDetailsFields['avg' + capitalize($scope.currentCycle)].value = ($scope.sampleDetailsFields['total' + capitalize($scope.currentCycle)].value) / ($scope.currentApplianceList.length);
                    $scope.sampleDetailsFields.avgCombined.value = $scope.sampleDetailsFields.totalCombined.value / ($scope.sampleDetailsFields.summerDeviceCount.value + $scope.sampleDetailsFields.winterDeviceCount.value);
                    break;
                default:
                    console.log('Unexpected operation encountered..');
            }

        };

        $scope.optionalChanged = function () {
            console.debug($scope.optional);
            if ($scope.optional) {
                $scope.optionalDevices = [];
                $scope.optionalDeviceContribution = 0;
                var optionalOwnership = XLSX.utils.sheet_to_json($scope.workbook.Sheets["Optional Appliance Ownership " + capitalize($scope.currentCycle)]);
                var optionalTimeUsage = XLSX.utils.sheet_to_json($scope.workbook.Sheets["Optional Time of Use " + capitalize($scope.currentCycle)]);
                $scope.populateAvgPower($scope.currentCycle);
                for (var appliance of Object.keys(optionalOwnership[$scope.selectedSample])) {
                    if (optionalOwnership[$scope.selectedSample][appliance] == 1) {
                        var tempObj = {};
                        tempObj.name = appliance;
                        tempObj.hoursUsed = optionalTimeUsage[$scope.selectedSample][appliance];
                        tempObj.avgPowerForDevice = $scope['optionalAvgDevicePower' + capitalize($scope.currentCycle)][appliance];
                        $scope.optionalDeviceContribution += tempObj.hoursUsed * (tempObj.avgPowerForDevice / 60);
                        $scope.optionalDevices.push(tempObj);
                    }
                }
                $scope.modifySampleDataPoint('add');
                $scope.insertIntoChart();
            } else if ($scope.optional == false) {
                $scope.deviceChartData.hourSeries.splice(($scope.deviceChartData.hourSeries.length - $scope.optionalDevices.length), $scope.optionalDevices.length);
                $scope.deviceChartData.percentageSeries.splice(($scope.deviceChartData.percentageSeries.length - $scope.optionalDevices.length), $scope.optionalDevices.length);
                console.debug($scope.deviceChartData);
                $scope.modifySampleDataPoint('deduct');
                $scope.optionalDevices = [];
            }
            //console.debug($scope.optionalDevices);
            $scope.calculateSAPCoeff();
        };
    }
]);

ExcelParser.service('ExcelParserService', [
    function () {

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
            },
            {
                'id': 'tfa',
                'display': 'TFA(Total Floor Area)'
            },
            {
                'id': 'Ea',
                'display': 'Ea Coefficient'
            },
            {
                'id': 'Eb',
                'display': 'Eb Coefficient'
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
            'value': 0,
            'visible': true
        };
        this.sampleDetailsFields.summerDeviceCount = {
            'display': 'Total Device Count(S)',
            'value': 0,
            'visible': false
        };
        this.sampleDetailsFields.totalWinter = {
            'display': 'Total Electricity Consumption(W)',
            'value': 0,
            'visible': true
        };
        this.sampleDetailsFields.winterDeviceCount = {
            'display': 'Total Device Count(W)',
            'value': 0,
            'visible': false
        };
        this.sampleDetailsFields.totalCombined = {
            'display': 'Total Electricity Consumption(Combined)',
            'value': 0,
            'visible': true
        };
        this.sampleDetailsFields.avgSummer = {
            'display': 'Average Electricity Consumption(S)',
            'value': 0,
            'visible': true
        };
        this.sampleDetailsFields.avgWinter = {
            'display': 'Average Electricity Consumption(W)',
            'value': 0,
            'visible': true
        };
        this.sampleDetailsFields.avgCombined = {
            'display': 'Average Electricity Consumption(Combined)',
            'value': 0,
            'visible': true
        };

        this.TFAData = function (houseSize, personCount) {
            if (houseSize == 1 && personCount == 1) {
                return 37;
            } else if (houseSize == 1 && personCount == 2) {
                return 50;
            } else if (houseSize == 2 && personCount == 3) {
                return 61;
            } else if (houseSize == 2 && personCount == 4) {
                return 70;
            } else if (houseSize == 3 && personCount == 4) {
                return 74;
            } else if (houseSize == 3 && personCount == 5) {
                return 86;
            } else if (houseSize == 3 && personCount == 6) {
                return 95;
            } else if (houseSize == 4 && personCount == 5) {
                return 90;
            } else if (houseSize == 4 && personCount == 6) {
                return 99;
            }
        };

    }
])
;

ExcelParser.filter('to_trusted', ['$sce',
    function ($sce) {
        return function (text) {
            return $sce.trustAsHtml(text);
        }
    }
]);