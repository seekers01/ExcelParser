'use strict';

var ExcelParser = angular.module('ExcelParser', []);

ExcelParser.controller('MainCtrl', ['$scope', 'ExcelParserService',
    function($scope, Service){
        $scope.fileSelected = null;
        $scope.sheet_name_list = [];
        $scope.welcome = "Hello";

        $scope.loadFile = function(){
            console.log($scope.fileSelected);

            /* set up XMLHttpRequest */
            var url = "HHM-V6.xlsx";
            var oReq = new XMLHttpRequest();
            oReq.open("GET", url, true);
            oReq.responseType = "arraybuffer";

            oReq.onload = function(e) {
                var arraybuffer = oReq.response;

                /* convert data to binary string */
                var data = new Uint8Array(arraybuffer);
                var arr = new Array();
                for(var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
                var bstr = arr.join("");

                /* Call XLSX */
                $scope.workbook = XLSX.read(bstr, {type:"binary"});
                //console.log(workbook);

                $scope.sheet_name_list = $scope.workbook.SheetNames;
                console.log($scope.sheet_name_list);
                $scope.welcome = "Hello Again";
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
    }
]);

ExcelParser.service('ExcelParserService', ['$http',
    function($http) {

        this.safeApply = function($scope, fn) {
            var phase = $scope.$root.$$phase;
            if(phase == '$apply' || phase == '$digest') {
                if(fn && (typeof(fn) === 'function')) {
                    fn();
                }
            } else {
                $scope.$apply(fn);
            }
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