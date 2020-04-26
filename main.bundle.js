webpackJsonp(["main"],{

/***/ "../../../../../src/$$_gendir lazy recursive":
/***/ (function(module, exports) {

function webpackEmptyAsyncContext(req) {
	// Here Promise.resolve().then() is used instead of new Promise() to prevent
	// uncatched exception popping up in devtools
	return Promise.resolve().then(function() {
		throw new Error("Cannot find module '" + req + "'.");
	});
}
webpackEmptyAsyncContext.keys = function() { return []; };
webpackEmptyAsyncContext.resolve = webpackEmptyAsyncContext;
module.exports = webpackEmptyAsyncContext;
webpackEmptyAsyncContext.id = "../../../../../src/$$_gendir lazy recursive";

/***/ }),

/***/ "../../../../../src/app/app.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "@media print{\r\n    .noprint{\r\n        display: none;\r\n    }\r\n}\r\n\r\n\r\nli.item{\r\n    cursor: pointer;\r\n    white-space:nowrap;\r\n}\r\nli.item:hover{\r\n    background: #296192;\r\n    color: white;\r\n}\r\n::-webkit-scrollbar {\r\n    width: 8px;\r\n    height: 8px;\r\n}\r\n\r\n::-webkit-scrollbar-track {\r\n    -webkit-box-shadow: inset 0 0 6px rgba(0,0,0,0.3); \r\n    border-radius: 4px;\r\n}\r\n\r\n::-webkit-scrollbar-thumb {\r\n    border-radius: 4px;\r\n    -webkit-box-shadow: inset 0 0 6px rgba(0,0,0,0.5); \r\n}\r\n\r\n.testTable th,.testTable td{\r\n    width: 300px;\r\n}\r\n\r\ndiv{\r\n    scrollbar-base-color: #000;\r\n    scrollbar-face-color: #000;\r\n    scrollbar-3dlight-color: #000;\r\n    scrollbar-highlight-color: #000;\r\n    scrollbar-track-color: #c0c0c0;\r\n    scrollbar-arrow-color: black;\r\n    scrollbar-shadow-color: #000;\r\n    scrollbar-dark-shadow-color: #000;\r\n    scrollbar-width:8px;\r\n    scrollbar-height:8px;\r\n  }", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/app.component.html":
/***/ (function(module, exports) {

module.exports = "<div (window:resize)=\"onResize($event)\">\n<div [style.visibility]=\"oplandb.isLoading || oplandb.isPosting? 'visible' : 'hidden'\">\n        <div  \n            [style.top]=\"getMaskTop() + 'px'\"\n            [style.left]=\"getMaskLeft() + 'px'\"\n            style=\"opacity: 0.5;position:absolute;z-index:9999;background:white;height:100vh;width:100%;\">\n                \n        </div>\n        <div class=\"p-2 px-3\" id=\"spinMessage\" \n            style=\"z-index:10000;position:absolute;background:white;border-radius:10px;box-shadow: 5px 5px 15px #c0c0c0;\" \n            [style.top]=\"oplandb.getMessageTop + 'px'\"\n            [style.left]=\"oplandb.getMessageLeft + 'px'\"\n        >\n            <div style=\"background:white;border:1px solid transparent;border-radius:10px;\">\n                <table><tr>\n                    <td><i class=\"fa fa-spinner fa-spin fa-2x fa-fw\"></i></td>\n                    <td style=\"padding-top:8px;\">\n                        <label class=\"form-label\">\n                            {{ oplandb.isPosting ? oplandb.isPostingMessage : oplandb.isLoadingMessage}}\n                        </label></td>\n                </tr></table>\n            </div>\n        </div>\n    </div>\n\n<app-frm-chart *ngIf=\"!udf.isDisplayDataComponent() && udf.currentModule=='Calendar'\"></app-frm-chart>\n<app-frm-chart-activities *ngIf=\"!udf.isDisplayDataComponent() && udf.currentModule=='CalendarActivity'\"></app-frm-chart-activities>\n<app-frm-chart-activities-mtiap *ngIf=\"!udf.isDisplayDataComponent() && udf.currentModule=='CalendarActivityMTIAP'\"></app-frm-chart-activities-mtiap>\n \n\n\n<div id=\"mainWrapper\">\n\n\n        <div class=\"noprint\" style=\"position:absolute;width:100%;z-index:5;\"\n            [style.top]=\"(udf.bannerHeight-udf.bannerMenuHeight)+'px'\"\n            (mouseenter)=\"mouseEvent($event)\"\n        >\n                <div class=\"row\">\n                    <div class=\"col\"><!--START OF LEFT COLUMN-->\n\n                        <div  class=\"form-group noprint\"\n                            style=\"display:flex;margin-top:5px;width:750px;\" \n                            *ngIf=\"!udf.isDisplayDataComponent() && udf.currentModule=='CalendarActivityMTIAP'\"\n                        >\n\n                            <label class=\"form-control-label px-2\" style=\"font-size:0.9em;padding:0px;margin:0px;margin-top:4px;white-space: nowrap;\">Select Site:</label>\n                            <select class=\"form-control form-control-sm mx-2\" \n                                [(ngModel)]=\"udf.calActivitySiteMTIAPDisp\" \n                                name=\"calSite\" \n                                style=\"text-align:center;\"\n                                [style.width.px]=\"80\"\n                                (change)=\"onCalParamChange()\"\n                                >\n                                <option value=\"SWP\">SWP</option>\n                                <option value=\"OGP\">OGP</option>\n                            </select>     \n\n                            <label class=\"form-control-label px-2\" \n                                style=\"font-size:0.9em;padding:0px;margin:0px;margin-top:4px;white-space: nowrap;\">Select Year:</label>\n\n                                <select class=\"form-control form-control-sm mx-2\" \n                                [(ngModel)]=\"udf.calActivityMTIAPYearDisp\" \n                                name=\"calYear\" \n                                style=\"text-align:center;\"\n                                [style.width.px]=\"70\"\n                                (change)=\"onCalParamChange()\"\n                                >\n                                <option [value]=\"yr\" *ngFor=\"let yr of getYears()\" >{{yr}}</option>\n                            </select>      \n\n\n                            <label class=\"form-control-label px-2\" \n                                style=\"font-size:0.9em;padding:0px;margin:0px;margin-top:4px;white-space: nowrap;\">Pages/Year:</label>\n\n                                                            \n                            <select class=\"form-control form-control-sm mx-2\" \n                                [(ngModel)]=\"udf.calActivityMTIAPPagePerYearDisp\" \n                                name=\"calPages\" \n                                style=\"text-align:center;\"\n                                [style.width.px]=\"50\"\n                                (change)=\"onCalParamChange()\"\n                                >\n                                <option value=\"2\">2</option>\n                                <option value=\"3\">3</option>\n                                <option value=\"4\">4</option>\n                                <option value=\"6\">6</option>\n                                <!--option value=\"12\">12</option--><!--Removed temporarily because of minor issue on band positioning-->\n                            </select>        \n                            \n                            <button \n                                [title]=\"'Load summary data for MTIAP'\"\n                                class=\"btn btn-link p-0 px-2\" \n                                [style.cursor]=\"isScopeSelectorChangedMTIAP()?'pointer':''\"\n                                [disabled]=\"isScopeSelectorChangedMTIAP()?'':'disabled'\" \n                                onfocus=\"this.blur();\"\n                                (click)=\"getSummaryDataMTIAP()\">\n                                    <i class=\"fa fa-refresh\" aria-hidden=\"true\"\n                                    style=\"font-size:20px;border:none;\"\n                                ></i>\n                            </button>                               \n\n                            <label class=\"form-control-label px-2\" \n                                style=\"font-size:0.9em;padding:0px;margin:0px;margin-top:4px;white-space: nowrap;\">Page:</label>\n\n                            <select class=\"form-control form-control-sm text-center mx-2\" \n                                [(ngModel)]=\"udf.calActivityMTIAPPage\" \n                                name=\"calCurrentPage\" \n                                style=\"text-align:center;z-index:10;\"\n                                [style.width.px]=\"140\"\n                                (change)=\"switchPage(period)\"\n                                >\n                                <option class=\"p-2 text-center\" *ngFor=\"let period of udf.MTIAPPeriods(true)\" \n                                    value=\"{{period.sem}}\">{{getMonthRange(period.sem)}}</option>\n                            </select>                                                           \n                                \n                            <!--button \n                                *ngFor=\"let period of udf.MTIAPPeriods(true)\"\n                                class=\"btn btn-secondary p-0\" \n                                (click)=\"switchPage(period.sem)\" \n                                [ngClass]=\"{'disabled':udf.calActivityMTIAPPage==period.sem}\"\n                                style=\"cursor:pointer;font-size:.8em;height:28px;width:28px;\">\n                                {{period.sem}}\n                            </button-->\n\n                        </div>\n                        \n                        <div  class=\"form-group noprint\"\n                            style=\"display:flex;margin-top:5px;width:750px;\" \n                            *ngIf=\"!udf.isDisplayDataComponent() && udf.currentModule=='CalendarActivity'\"\n                        >\n                        <label class=\"form-control-label pl-2\" style=\"font-size:0.9em;padding:0px;margin:0px;margin-top:4px;white-space: nowrap;\">Select Site:</label>\n                        <select class=\"form-control form-control-sm mx-2\" \n                            [(ngModel)]=\"udf.calActivitySiteDisp\" \n                            name=\"calSite\" \n                            style=\"text-align:center;\"\n                            [style.width.px]=\"80\"\n                            (change)=\"onCalParamChange()\"\n                            >\n                            <option value=\"SWP\">SWP</option>\n                            <option value=\"OGP\">OGP</option>\n                        </select>     \n                        \n                        <label class=\"form-control-label\" style=\"font-size:0.9em;padding:0px;margin:0px;margin-top:4px;white-space: nowrap;\">Select Month:</label>\n                        <select class=\"form-control form-control-sm mx-2\" \n                            [(ngModel)]=\"udf.calActivityMonthDisp\" \n                            name=\"calMonth\" \n                            style=\"width:120px; text-align:center;\"\n                            (change)=\"onCalParamChange()\"\n                            >\n                            <option [value]=\"midx\" *ngFor=\"let mo of  getMonths();let midx=index;\">{{mo}}</option>\n                        </select>     \n\n                        <label class=\"form-control-label\" style=\"font-size:0.9em;padding:0px;margin:0px;margin-top:4px;white-space: nowrap;\">Select Year:</label>\n                        <select class=\"form-control form-control-sm px-1 mx-2\" \n                            [(ngModel)]=\"udf.calActivityYearDisp\" \n                            name=\"calYear\" \n                            style=\"text-align:center;\"\n                            [style.width.px]=\"70\"\n                            (change)=\"onCalParamChange()\"\n                            >\n                            <option [value]=\"yr\" *ngFor=\"let yr of getYears()\" >{{yr}}</option>\n                        </select>     \n\n<button \n    [title]=\"'Load summary data for ' + getMonths()[udf.calActivityMonthDisp] + ' ' + udf.calActivityYearDisp\"\n    class=\"btn btn-link p-0 px-1\" \n    [style.cursor]=\"isScopeSelectorChanged()?'pointer':''\"\n    [disabled]=\"isScopeSelectorChanged()?'':'disabled'\" \n    onfocus=\"this.blur();\"\n    (click)=\"getSummaryData()\">\n    <i class=\"fa fa-refresh\" aria-hidden=\"true\" style=\"font-size:20px;border:none;\"></i>\n</button>\n\n<label class=\"form-control-label px-0\" \n    style=\"font-size:0.9em;padding:0px;margin:0px;margin-top:4px;margin-left:5px;\">Zoom:</label>\n<select class=\"form-control form-control-sm mx-2\" \n(change)=\"scaleChanged()\"\n[(ngModel)]=\"udf.calActivityChartScale\" \nname=\"calActivityChartScale\" \nstyle=\"text-align:center;\"\n[style.width.px]=\"80\"\ntitle=\"Change display zoom factor...\"\n>\n<option value=\"1\">100%</option>\n<option value=\"1.1\">110%</option>\n<option value=\"1.25\">125%</option>\n<option value=\"1.5\">150%</option>\n<option value=\"1.75\">175%</option>\n<option value=\"2.0\">200%</option>\n</select>     \n\n\n\n<!--button style=\"position:absolute;top:0px;left:0px;z-index:50;left:450px;cursor:pointer;\" \nclass=\"btn btn-link\"\ndata-toggle=\"tooltip\" \n[style.top]=\"(udf.bannerHeight-udf.bannerMenuHeight)+'px'\"\n(click)=\"reloadCalendar()\"\ntitle=\"Reload calendar with the new specified scope dates ...\"\n[disabled]=\"disableReload()\"\n><i class=\"fa fa-refresh\" aria-hidden=\"true\"\nstyle=\"font-size:20px;\"\n></i></button-->\n\n                        </div>\n                        <app-pobprofile *ngIf=\"!udf.isDisplayDataComponent() && udf.currentModule=='POBProfile'\"></app-pobprofile>\n\n\n                    <div class=\"form-group\" style=\"display:flex;margin-top:5px;\" \n                        *ngIf=\"!udf.isDisplayDataComponent() && (udf.currentModule=='POBProfile')\">\n                        <label class=\"form-control-label px-1\" style=\"font-size:0.9em;padding:0px;margin:0px;margin-top:1px;\">Start date:</label>\n                        <div class=\"input-group input-group-sm mb-2 mr-sm-2 mb-sm-0\"\n                            style=\"width:150px;\">\n                            <input class=\"form-control form-control-sm text-center\" \n                                type=\"text\" \n                                id=\"start-date\"\n                                [(ngModel)]=\"udf.profDispStartDate\" \n                                name=\"profDispStartDate\"\n                                [dpDayPicker]=\"udf.datePickerConfig\"\n                                onfocus=\"this.blur();\" \n                            >\n                            <div class=\"input-group-addon\"\n                                (click)=\"showCalendar($event,'start-date')\"\n                            >\n                                <a href=\"#\">\n                                    <i class=\"fa fa-calendar text-primary\"></i>\n                                </a>\n                            </div>\n                        </div>                    \n                        <label class=\"form-control-label px-1\" style=\"font-size:0.9em;padding:0px;margin:0px;margin-top:1px;\">End date:</label>\n                        <div class=\"input-group input-group-sm mb-2 mr-sm-2 mb-sm-0\"\n                            style=\"width:150px;\">\n                            <input class=\"form-control form-control-sm text-center\" \n                                type=\"text\" \n                                id=\"end-date\"\n                                [(ngModel)]=\"udf.profDispEndDate\" \n                                name=\"profDispEndDate\"\n                                [dpDayPicker]=\"udf.datePickerConfig\"\n                                onfocus=\"this.blur();\" \n                            >\n                            <div class=\"input-group-addon\"\n                                (click)=\"showCalendar($event,'end-date')\"\n                            >\n                                <a href=\"#\">\n                                    <i class=\"fa fa-calendar text-primary\"></i>\n                                </a>\n                            </div>\n                        </div>                    \n                        <div  style=\"flex-grow:1;\">&nbsp;</div>\n                    </div>                        \n\n                        <div class=\"form-group\" style=\"display:flex;margin-top:5px;\" \n                            *ngIf=\"!udf.isDisplayDataComponent() && (udf.currentModule=='Calendar')\">\n                            <label class=\"form-control-label px-1\" style=\"font-size:0.9em;padding:0px;margin:0px;margin-top:1px;\">Start date:</label>\n                            <div class=\"input-group input-group-sm mb-2 mr-sm-2 mb-sm-0\"\n                                style=\"width:150px;\">\n                                <input class=\"form-control form-control-sm text-center\" \n                                    type=\"text\" \n                                    id=\"start-date\"\n                                    [(ngModel)]=\"udf.calDispStartDate\" \n                                    name=\"calDispStartDate\"\n                                    [dpDayPicker]=\"udf.datePickerConfig\"\n                                    onfocus=\"this.blur();\" \n                                >\n                                <div class=\"input-group-addon\"\n                                    (click)=\"showCalendar($event,'start-date')\"\n                                >\n                                    <a href=\"#\">\n                                        <i class=\"fa fa-calendar text-primary\"></i>\n                                    </a>\n                                </div>\n                            </div>                    \n                            <label class=\"form-control-label px-1\" style=\"font-size:0.9em;padding:0px;margin:0px;margin-top:1px;\">End date:</label>\n                            <div class=\"input-group input-group-sm mb-2 mr-sm-2 mb-sm-0\"\n                                style=\"width:150px;\">\n                                <input class=\"form-control form-control-sm text-center\" \n                                    type=\"text\" \n                                    id=\"end-date\"\n                                    [(ngModel)]=\"udf.calDispEndDate\" \n                                    name=\"calDispEndDate\"\n                                    [dpDayPicker]=\"udf.datePickerConfig\"\n                                    onfocus=\"this.blur();\" \n                                >\n                                <div class=\"input-group-addon\"\n                                    (click)=\"showCalendar($event,'end-date')\"\n                                >\n                                    <a href=\"#\">\n                                        <i class=\"fa fa-calendar text-primary\"></i>\n                                    </a>\n                                </div>\n                            </div>                    \n                            <div  style=\"flex-grow:1;\">&nbsp;</div>\n                        </div>\n                    </div>  <!-- END OF MAIN LEFT COLUMN -->\n                    <div class=\"pr-4 col\" >   <!-- START OF MAIN RIGHT COLUMN -->\n                        <div class=\"row\" style=\"position: relative;\">\n                        <div class=\"form-group col\" \n                            style=\"text-align:right;margin-top:5px;margin-right:50px;white-space: nowrap;\" >\n                            <label class=\"form-control-label\" style=\"font-size:0.9em;\" >\n                                User: {{oplandb.logInfo.fullName}} ({{oplandb.logInfo.role}})\n                            </label>\n                        </div>\n                        <div class=\"noprint\" style=\"text-align:right;margin-top:2px;\">\n\n                            <nav class=\"navbar  navbar-light\" style=\"padding:0px;margin:3px;height:35px;\">\n                                <button id=\"navToggler\"\n                                    style=\"margin:0px;padding:3px;padding-top:0px;padding-bottom:0px;cursor:pointer;\" \n                                    class=\"navbar-toggler navbar-toggler-right\" \n                                    type=\"button\" data-toggle=\"collapse\" \n                                    data-target=\"#collapsingNavbar2\">\n                                    <span class=\"navbar-toggler-icon my-toggler\" style=\"font-size:16px;\"></span>\n                                </button>\n                                <a href=\"/\" class=\"navbar-brand\" style=\"margin:0px;padding:0px;\" href=\"#\">&nbsp;</a>\n                                <div class=\"navbar-collapse collapse\" id=\"collapsingNavbar2\" \n                                    [class.show]=\"isMenuOn\"\n                                    style=\"border:1px solid #c0c0c0;background:#fefefe;z-index:250;\">\n                                    <ul class=\"navbar-nav mx-auto\">\n                                        <li class=\"nav-item pr-5 text-nowrap active\"  id=\"acn_Personnel\" (click)=\"imClicked($event)\"  style=\"text-align:left;padding-left:5px;\">\n                                            <a class=\"nav-link\" href=\"#\"><i class=\"fa fa-user-o\" aria-hidden=\"true\"></i>&nbsp;&nbsp;Personnel</a>\n                                        </li>\n                                        <li class=\"nav-item text-nowrap\"  id=\"acn_Teams\" (click)=\"imClicked($event)\" style=\"text-align:left;padding-left:5px;\">\n                                            <a class=\"nav-link\" href=\"#\"><i class=\"fa fa-users\" aria-hidden=\"true\"></i>&nbsp;&nbsp;Teams</a>\n                                        </li>\n                                        <li class=\"nav-item text-nowrap\"  id=\"acn_Activities\" (click)=\"imClicked($event)\" style=\"text-align:left;padding-left:5px;\">\n                                            <a class=\"nav-link\" href=\"#\"><i class=\"fa fa-list-alt\" aria-hidden=\"true\"></i>&nbsp;&nbsp;Activities</a>\n                                        </li>\n                                        <li class=\"divider p-0\"><hr class=\"m-0\"/>\n                                        <li class=\"nav-item text-nowrap\"  id=\"acn_CalendarActivity\" (click)=\"imClicked($event)\" style=\"text-align:left;padding-left:5px;\">\n                                            <a class=\"nav-link\" href=\"#\"><i class=\"fa fa-calendar\" aria-hidden=\"true\"></i>&nbsp;&nbsp;STIAP View</a>\n                                        </li>\n                                        <li  class=\"nav-item text-nowrap\"  id=\"acn_CalendarActivityMTIAP\" (click)=\"imClicked($event)\" style=\"text-align:left;padding-left:5px;\">\n                                            <a class=\"nav-link\" href=\"#\"><i class=\"fa fa-calendar\" aria-hidden=\"true\"></i>&nbsp;&nbsp;MTIAP View</a>\n                                        </li>\n                                        <li class=\"nav-item text-nowrap\"  id=\"acn_Calendar\" (click)=\"imClicked($event)\" style=\"text-align:left;padding-left:5px;\">\n                                            <a class=\"nav-link\" href=\"#\"><i class=\"fa fa-calendar\" aria-hidden=\"true\"></i>&nbsp;&nbsp;POB Calendar</a>\n                                        </li>\n                                        <li class=\"nav-item text-nowrap\"  id=\"acn_POBProfile\" (click)=\"imClicked($event)\" style=\"text-align:left;padding-left:5px;\">\n                                            <a class=\"nav-link\" href=\"#\"><i class=\"fa fa-calendar\" aria-hidden=\"true\"></i>&nbsp;&nbsp;POB Profile</a>\n                                        </li>\n                                        <li class=\"divider p-0\"><hr class=\"m-0\"/>\n                                        <li class=\"nav-item text-nowrap\"  id=\"acn_Settings\" (click)=\"imClicked($event)\" style=\"text-align:left;padding-left:5px;\">\n                                            <a class=\"nav-link\" href=\"#\"><i class=\"fa fa-cog\" aria-hidden=\"true\"></i>&nbsp;&nbsp;Settings</a>\n                                        </li>\n                                        <li class=\"divider p-0\"><hr class=\"m-0\"/>\n                                        </li>\n                                        <li class=\"nav-item text-nowrap\"  id=\"acn_UserGuide\" (click)=\"imClicked($event)\" style=\"text-align:left;padding-left:5px;\">\n                                            <a class=\"nav-link\" href=\"#\"><i class=\"fa fa-book\" aria-hidden=\"true\"></i>&nbsp;&nbsp;User Guide</a>\n                                        </li>                                        \n                                    </ul>\n                                </div>\n                            </nav>\n                            \n                        </div> \n                    </div>            \n                    </div>\n\n                </div>\n            </div>\n\n    <div class=\"noprint\" id=\"banner\" style=\"position:relative;;\"\n        [style.top]=\"'0px'\"\n        [style.left]=\"'0px'\"\n        [style.height]=\"udf.bannerHeight + 'px'\"\n        [style.width]=\"udf.winWd()+'px'\"\n        (window:resize)=\"udf.dummy($event)\"\n    >\n\n        <img src=\"assets/images/shell_logo_d.png\" style=\"margin:5px;\" alt=\"image\">\n        <label class=\"h2\" style=\"text-shadow: 1px 2px #c0c0c0;position:absolute;top:20px;left:75px;color:#296192;font-weight:bold;\">\n            {{title}}<span style=\"font-family:'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;font-weight:normal;margin-left:5px;font-size:11px;text-shadow:none;\">( {{\"Ver&nbsp;:&nbsp;\" + appVersion}} )</span></label>\n\n        <!-- Banner menu section -->\n        <div\n            style=\"position:absolute;overflow:hidden;\"\n            [style.background]=\"udf.menuBackground\"\n            [style.top]=\"(udf.bannerHeight-udf.bannerMenuHeight)+'px'\"\n            [style.left]=\"'0px'\"\n            [style.width]=\"udf.winWd()+'px'\"\n            [style.height]=\"(udf.bannerMenuHeight)+'px'\"  \n        >\n        <div *ngIf=\"false\" style=\"position:relative;background:#EED674;height:2px;\">\n\n        </div>\n            <div style=\"position:relative;background:#e8bf0a;height:2px;\"\n            [style.top]=\"'0px'\"\n            >\n            &nbsp;\n            </div>        \n\n\n            <div style=\"position:relative;background:red;height:4px;\"\n            [style.top]=\"(udf.bannerMenuHeight-4) + 'px'\"\n            >\n            &nbsp;\n            </div>\n        </div>\n    \n    </div>\n    <!-- width of 100% for the next div element solves the problem with ID on resizing the details table -->\n    <div style=\"position:absolute;background:none;width:100%;\"  *ngIf=\"udf.isDisplayDataComponent()\"\n      [style.height]=\"udf.htNoBan(true)\"\n      [style.width]=\"udf.winWd()+'px'\"\n      [style.top]=\"udf.bannerHeight+'px'\"\n    >\n            \n                <div class=\"row\">\n                    <div *ngIf=\"udf.currentModule!='Calendar' && udf.currentModule!='Settings'\" class=\"col-4 px-0 py-0\"\n                        [style.height]=\"udf.htNoBan(true)\"\n                    >\n                        <app-nav-list \n                            (itemClicked)=\"onListItemClicked($event)\"\n                            (sortSelectionClicked)=\"onSortSelectionClicked($event)\"\n                            (thisComponent)=\"getNavComponentObject($event)\"\n                        ></app-nav-list>\n                    </div>\n                    <div class=\"col\" [style.height]=\"udf.htNoBan(true)\"\n                    >\n                        <div class=\"card border-0\" >\n                            <div class=\"card-block px-0 py-0\" style=\"background:none;overflow:hidden;\"\n                                [style.height]=\"udf.htNoBan(true)\"\n                            >\n                                    <app-frm-personnel \n                                        *ngIf=\"udf.currentModule!='Activities' && \n                                            udf.currentModule!='Teams' && \n                                            udf.currentModule!='Settings' \" \n                                        [currPersonnel]=\"currPersonnel\"></app-frm-personnel>\n                                    <app-frm-activity  *ngIf=\"udf.currentModule=='Activities'\" ></app-frm-activity>\n                                    <app-frm-team  *ngIf=\"udf.currentModule=='Teams'\"></app-frm-team>\n                                    <app-frm-settings  *ngIf=\"udf.currentModule=='Settings'\"></app-frm-settings>\n                                    \n                                    <!--\n            \n                                        \n            \n                                    -->\n                            </div>\n                        </div>\n                    </div>\n                </div>\n            \n            </div>\n</div>\n<!--app-pnl-vert-split></app-pnl-vert-split-->\n</div>"

/***/ }),

/***/ "../../../../../src/app/app.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return AppComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__models_data_models__ = __webpack_require__("../../../../../src/app/models/data-models.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2_jquery__ = __webpack_require__("../../../../jquery/dist/jquery.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2_jquery___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_2_jquery__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3__app_globals__ = __webpack_require__("../../../../../src/app/app.globals.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_4__components_frm_chart_activities_frm_chart_activities_component__ = __webpack_require__("../../../../../src/app/components/frm-chart-activities/frm-chart-activities.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_5__app_udfs__ = __webpack_require__("../../../../../src/app/app.udfs.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_6__app_oplandata__ = __webpack_require__("../../../../../src/app/app.oplandata.ts");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};







var AppComponent = (function () {
    function AppComponent(udf, oplandb) {
        var _this = this;
        this.udf = udf;
        this.oplandb = oplandb;
        this.allowed = true;
        this.MTIAPPage = 1;
        this.ctr = 1;
        this.displayDataComponents = (this.udf.currentModule != "Calendar" && this.udf.currentModule != "POBProfile" && this.udf.currentModule != "CalendarActivity");
        this.currPersonnel = new __WEBPACK_IMPORTED_MODULE_1__models_data_models__["d" /* dataPersonnel */]();
        this.title = '@SPEX Online Planning Tool';
        this.appVersion = " Beta 1.0.5 - Build:20191029-VM "; // Reveal XBR and Brithdate fields
        //appVersion:string=" Beta 1.0.5 - Build:20191029 ";  // Reveal XBR and Brithdate fields
        // appVersion:string=" Beta 1.0.4 - Build:20191028 ";  // Fix falseWeek header in MTIAP
        // appVersion:string=" Beta 1.0.3 - Build:20190815 ";  // Temporarily hide XBR field
        //appVersion:string=" Beta 1.0.2 - Build:20190813 ";  // Fixed milestone box in MTIAP
        //appVersion:string=" Beta 1.0.1 - Build:20190730 ";  // Fixed overlapping boxes in MTIAP
        //appVersion:string=" Beta 1.0.0 - Build:20190729 ";  // added PTW Onshore/Offshore
        //appVersion:string=" Alpha 2.4.5 - Build:20190626 ";
        //appVersion:string=" Alpha 2.4.3 - Build:20190220 ";
        //appVersion:string=" Alpha 2.4.2 - Build:20190123 ";
        //appVersion:string=" Alpha 2.3.8 - Build:20181213 ";
        //appVersion:string=" Alpha 2.3.7 - Build:20181024 ";
        //appVersion:string=" Alpha 2.3.4 - Build:20180803 ";
        //appVersion:string=" Alpha 2.3.0 - Build:20180515 ";
        this.chart = [];
        this.myChartData = { "message": "", "cod": "200", "city_id": 2643743, "calctime": 0.0875, "cnt": 3, "list": [{ "main": { "temp": 279.946, "temp_min": 279.946, "temp_max": 279.946, "pressure": 1016.76, "sea_level": 1024.45, "grnd_level": 1016.76, "humidity": 100 }, "wind": { "speed": 4.59, "deg": 163.001 }, "clouds": { "all": 92 }, "weather": [{ "id": 500, "main": "Rain", "description": "light rain", "icon": "10n" }], "rain": { "3h": 2.69 }, "dt": 1485717216 }, { "main": { "temp": 282.597, "temp_min": 282.597, "temp_max": 282.597, "pressure": 1012.12, "sea_level": 1019.71, "grnd_level": 1012.12, "humidity": 98 }, "wind": { "speed": 4.04, "deg": 226 }, "clouds": { "all": 92 }, "weather": [{ "id": 500, "main": "Rain", "description": "light rain", "icon": "10n" }], "rain": { "3h": 0.405 }, "dt": 1485745061 }, { "main": { "temp": 279.38, "pressure": 1011, "humidity": 93, "temp_min": 278.15, "temp_max": 280.15 }, "wind": { "speed": 2.6, "deg": 30 }, "clouds": { "all": 90 }, "weather": [{ "id": 701, "main": "Mist", "description": "mist", "icon": "50d" }, { "id": 741, "main": "Fog", "description": "fog", "icon": "50d" }], "dt": 1485768552 }] };
        this.isMenuOn = false;
        this.devOptions = {};
        this.devColDefs = [
            { width: 100 },
            { width: 300 },
            { width: 500 },
            { width: 800 },
            { width: 300 },
        ];
        this.devDat = [
            { c1: 'R1C1', c2: 'R1C2', c3: 'R1C3', c4: 'R1C4', c5: 'R1C5' },
            { c1: 'R2C1', c2: 'R2C2', c3: 'R2C3', c4: 'R2C4', c5: 'R2C5' },
            { c1: 'R3C1', c2: 'R3C2', c3: 'R3C3', c4: 'R3C4', c5: 'R3C5' },
            { c1: 'R4C1', c2: 'R4C2', c3: 'R4C3', c4: 'R4C4', c5: 'R4C5' },
            { c1: 'R5C1', c2: 'R5C2', c3: 'R5C3', c4: 'R5C4', c5: 'R5C5' },
            { c1: 'R6C1', c2: 'R6C2', c3: 'R6C3', c4: 'R6C4', c5: 'R6C5' },
            { c1: 'R7C1', c2: 'R7C2', c3: 'R7C3', c4: 'R7C4', c5: 'R7C5' },
            { c1: 'R8C1', c2: 'R8C2', c3: 'R8C3', c4: 'R8C4', c5: 'R8C5' },
            { c1: 'R9C1', c2: 'R9C2', c3: 'R9C3', c4: 'R9C4', c5: 'R9C5' },
            { c1: 'R10C1', c2: 'R10C2', c3: 'R10C3', c4: 'R10C4', c5: 'R10C5' },
            { c1: 'R11C1', c2: 'R11C2', c3: 'R11C3', c4: 'R11C4', c5: 'R11C5' },
            { c1: 'R12C1', c2: 'R12C2', c3: 'R12C3', c4: 'R12C4', c5: 'R12C5' },
            { c1: 'R13C1', c2: 'R13C2', c3: 'R13C3', c4: 'R13C4', c5: 'R13C5' },
            { c1: 'R14C1', c2: 'R14C2', c3: 'R14C3', c4: 'R14C4', c5: 'R14C5' },
            { c1: 'R15C1', c2: 'R15C2', c3: 'R15C3', c4: 'R15C4', c5: 'R15C5' },
            { c1: 'R16C1', c2: 'R16C2', c3: 'R16C3', c4: 'R16C4', c5: 'R16C5' },
            { c1: 'R17C1', c2: 'R17C2', c3: 'R17C3', c4: 'R17C4', c5: 'R17C5' },
            { c1: 'R18C1', c2: 'R18C2', c3: 'R18C3', c4: 'R18C4', c5: 'R18C5' },
            { c1: 'R19C1', c2: 'R19C2', c3: 'R19C3', c4: 'R19C4', c5: 'R19C5' },
            { c1: 'R20C1', c2: 'R20C2', c3: 'R20C3', c4: 'R20C4', c5: 'R20C5' },
            { c1: 'R21C1', c2: 'R21C2', c3: 'R21C3', c4: 'R21C4', c5: 'R21C5' },
            { c1: 'R22C1', c2: 'R22C2', c3: 'R22C3', c4: 'R22C4', c5: 'R22C5' },
            { c1: 'R23C1', c2: 'R23C2', c3: 'R23C3', c4: 'R23C4', c5: 'R23C5' },
            { c1: 'R24C1', c2: 'R24C2', c3: 'R24C3', c4: 'R24C4', c5: 'R24C5' },
            { c1: 'R25C1', c2: 'R25C2', c3: 'R25C3', c4: 'R25C4', c5: 'R25C5' },
            { c1: 'R26C1', c2: 'R26C2', c3: 'R26C3', c4: 'R26C4', c5: 'R26C5' },
            { c1: 'R27C1', c2: 'R27C2', c3: 'R27C3', c4: 'R27C4', c5: 'R27C5' },
            { c1: 'R28C1', c2: 'R28C2', c3: 'R28C3', c4: 'R28C4', c5: 'R28C5' },
            { c1: 'R29C1', c2: 'R29C2', c3: 'R29C3', c4: 'R29C4', c5: 'R29C5' },
            { c1: 'R30C1', c2: 'R30C2', c3: 'R30C3', c4: 'R30C4', c5: 'R30C5' },
            { c1: 'R31C1', c2: 'R31C2', c3: 'R31C3', c4: 'R31C4', c5: 'R31C5' },
            { c1: 'R32C1', c2: 'R32C2', c3: 'R32C3', c4: 'R32C4', c5: 'R32C5' },
            { c1: 'R33C1', c2: 'R33C2', c3: 'R33C3', c4: 'R33C4', c5: 'R33C5' },
            { c1: 'R34C1', c2: 'R34C2', c3: 'R34C3', c4: 'R34C4', c5: 'R34C5' },
            { c1: 'R35C1', c2: 'R35C2', c3: 'R35C3', c4: 'R35C4', c5: 'R35C5' },
            { c1: 'R36C1', c2: 'R36C2', c3: 'R36C3', c4: 'R36C4', c5: 'R36C5' },
            { c1: 'R37C1', c2: 'R37C2', c3: 'R37C3', c4: 'R37C4', c5: 'R37C5' },
            { c1: 'R38C1', c2: 'R38C2', c3: 'R38C3', c4: 'R38C4', c5: 'R38C5' },
            { c1: 'R39C1', c2: 'R39C2', c3: 'R39C3', c4: 'R39C4', c5: 'R39C5' },
            { c1: 'R40C1', c2: 'R40C2', c3: 'R40C3', c4: 'R40C4', c5: 'R40C5' },
        ];
        this.tblOptions = {
            startRow: 3,
            endRow: 15,
            colDefs: [
                { width: 100, index: 0, order: 1 },
                { width: 200, index: 1, order: 0 },
                { width: 300, index: 2, order: 3 },
                { width: 400, index: 3, order: 4 },
                { width: 500, index: 4, order: 2 },
            ]
        };
        setTimeout(function () {
            //alert('hello');
            _this.oplandb.getConfig();
            _this.oplandb.getCredentials();
            //this.getSummaryData()
            _this.oplandb.setWindowDimensions();
        }, 300);
        /*
            $(document).keydown(function(event) {
              if (event.repeat != undefined) {
                this.allowed = !event.repeat;
              }
              if (!this.allowed) return;
              oplandb.isCtrlKey=false;
              oplandb.isEditing=oplandb.isCtrlKey;
              this.allowed = false;
            });
        
            $(document).keyup(function(e) {
              oplandb.isCtrlKey=false;
              this.allowed = true;
            });
            $(document).focus(function(e) {
              oplandb.isCtrlKey=false;
              oplandb.isEditing=oplandb.isCtrlKey;
              this.allowed = true;
            });
        
            */
    }
    AppComponent.prototype.onKeyPress = function (e) {
        //$(window).bind('keydown',this.onKeyDown);
        //    $(window).unbind('keydown');
        //    $(window).bind('keydown',this.onKeyDown);
        //    $(window).unbind('keyup');
    };
    AppComponent.prototype.onKeyUp = function (e) {
        //$(window).bind('keydown',this.onKeyDown);
        //    $(window).unbind('keydown');
        //    $(window).bind('keydown',this.onKeyDown);
        //    $(window).unbind('keyup');
    };
    AppComponent.prototype.onKeyDown = function (e) {
        //$(window).bind('keyup',this.onKeyUp);
        //  $(window).unbind('keyup');
        //  $(window).bind('keyup',this.onKeyUp);
        //  $(window).unbind('keydown');
    };
    AppComponent.prototype.onListItemClicked = function (item) {
        //alert("event:" + item.lastName);
        this.currPersonnel = item;
    };
    AppComponent.prototype.onSortSelectionClicked = function (item) {
        //alert(item);
        this.udf.currentModule = item;
    };
    AppComponent.prototype.getNavComponentObject = function (e) {
    };
    AppComponent.prototype.imClicked = function (e) {
        var id = __WEBPACK_IMPORTED_MODULE_2_jquery__(e.srcElement).parent().attr('id');
        if (id == 'acn_UserGuide') {
            __WEBPACK_IMPORTED_MODULE_2_jquery__('#navToggler').trigger('click');
            var openNewWin = (this.UGWin == undefined);
            if (!openNewWin)
                openNewWin = this.UGWin.closed;
            if (openNewWin)
                this.UGWin = window.open('assets/docs/user_guide.pdf', '_blank');
            try {
                this.UGWin.focus();
            }
            catch (e) {
                alert(e.message);
            }
            return;
        }
        this.udf.searchText = '';
        __WEBPACK_IMPORTED_MODULE_2_jquery__(e.srcElement.parentNode.parentNode.parentNode).removeClass('show');
        var module = id.split('_')[1];
        window.localStorage.setItem("currentModule", module);
        if (module == 'Calendar' ||
            module == 'CalendarActivity' ||
            module == 'CalendarActivityMTIAP' ||
            module == 'POBProfile') {
            this.udf.currentModule = module;
            this.displayDataComponents = false;
        }
        else {
            this.oplandb.onActivityLoadAction = true;
            if (module == 'Personnel') {
                this.udf.currentModule = 'By Name';
                this.oplandb.setPersonnelNavGroup();
            }
            else {
                this.udf.currentModule = module;
                this.oplandb.setOtherNavGroup();
            }
            this.displayDataComponents = true;
        }
    };
    AppComponent.prototype.onResize = function (e) {
        this.oplandb.setWindowDimensions();
    };
    AppComponent.prototype.getYears = function () {
        //let min:number = parseInt(this.oplandb.minDate.substring(0,4));
        //let max:number = parseInt(this.oplandb.maxDate.substring(0,4));
        /*let min:number=2010;
        let max:number=2025;
        let i:number;
        let ret:Array<number>=[];
    
        for(i=min;i<=max;i++){
          console.log(i);
          ret.push[i];
        }
        console.log(ret);
        return ret;*/
        return [2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025];
    };
    AppComponent.prototype.getMonths = function () {
        return __WEBPACK_IMPORTED_MODULE_3__app_globals__["monthNames"];
    };
    AppComponent.prototype.onCalParamChange = function () {
        //if(this.calActivities!=undefined)this.calActivities.initVars();
        this.oplandb.calActivitiesInitVars();
    };
    AppComponent.prototype.showCalendar = function (e, ctlID) {
        __WEBPACK_IMPORTED_MODULE_2_jquery__('#' + ctlID).focus();
    };
    AppComponent.prototype.isScopeSelectorChangedMTIAP = function () {
        return (this.udf.calActivityMTIAPYearDisp != this.udf.calActivityMTIAPYear ||
            this.udf.calActivitySiteMTIAPDisp != this.udf.calActivitySiteMTIAP ||
            this.udf.calActivityMTIAPPagePerYearDisp != this.udf.calActivityMTIAPPagePerYear);
    };
    AppComponent.prototype.isScopeSelectorChanged = function () {
        return (this.udf.calActivityMonthDisp != this.udf.calActivityMonth ||
            this.udf.calActivityYearDisp != this.udf.calActivityYear ||
            this.udf.calActivitySiteDisp != this.udf.calActivitySite);
    };
    AppComponent.prototype.getMaskTop = function () {
        //return this.udf.bannerHeight;
        return 0;
    };
    AppComponent.prototype.getMaskLeft = function () {
        return 0;
    };
    AppComponent.prototype.getSummaryDataMTIAP = function () {
        var _this = this;
        // commit new select filter values
        this.udf.calActivitySiteMTIAP = this.udf.calActivitySiteMTIAPDisp;
        this.udf.calActivityMTIAPYear = this.udf.calActivityMTIAPYearDisp;
        this.udf.calActivityMTIAPPagePerYear = this.udf.calActivityMTIAPPagePerYearDisp;
        this.udf.calActivityMTIAPPage = 1;
        // set local storage for bookmarking
        localStorage.setItem("mtiapSite", this.udf.calActivitySiteMTIAP);
        localStorage.setItem("mtiapYear", String(this.udf.calActivityMTIAPYear));
        localStorage.setItem("mtiapPagePerYear", String(this.udf.calActivityMTIAPPagePerYear));
        // move to dummy module
        this.udf.currentModule = "Test";
        // soad STIAP module
        setTimeout(function () { _this.udf.currentModule = "CalendarActivityMTIAP"; }, 100);
    };
    AppComponent.prototype.getSummaryData = function () {
        var _this = this;
        if (this.calActivities != undefined) {
            // commit new select filter values
            this.udf.calActivitySite = this.udf.calActivitySiteDisp;
            this.udf.calActivityMonth = this.udf.calActivityMonthDisp;
            this.udf.calActivityYear = this.udf.calActivityYearDisp;
            // set local storage for bookmarking
            localStorage.setItem("stiapSite", this.udf.calActivitySite);
            localStorage.setItem("stiapMonth", String(this.udf.calActivityMonth));
            localStorage.setItem("stiapYear", String(this.udf.calActivityYear));
            // move to dummy module
            this.udf.currentModule = "Test";
            // soad STIAP module
            setTimeout(function () { _this.udf.currentModule = "CalendarActivity"; }, 100);
        }
    };
    AppComponent.prototype.scaleChanged = function () {
        localStorage.setItem("stiapScale", String(this.udf.calActivityChartScale));
    };
    AppComponent.prototype.ngOnInit = function () {
        var _this = this;
        setTimeout(function () { _this.oplandb.getPositions(); }, 1000);
        this.devInit();
    };
    AppComponent.prototype.showProfile = function (e, dt) {
    };
    AppComponent.prototype.onItemClick = function (item) {
        alert(item.name);
    };
    AppComponent.prototype.switchPage = function (per) {
        if (!per)
            return;
        var pg = per.sem;
        this.udf.calActivityMTIAPPage = pg;
        localStorage.setItem("mtiapPage", String(this.udf.calActivityMTIAPYear));
    };
    AppComponent.prototype.getMonthRange = function (per) {
        var ret = "";
        var spy = this.udf.calActivityMTIAPPagePerYear;
        var y2 = (per > spy);
        var mos = 12 / spy;
        var yr = Number(this.udf.calActivityMTIAPYear) + (y2 ? 1 : 0);
        var mo = (per - 1 - (y2 ? spy : 0)) * mos;
        if (mos == 1) {
            ret = __WEBPACK_IMPORTED_MODULE_3__app_globals__["monthNames"][mo].substr(0, 3) + ' ' + yr;
        }
        else {
            ret = __WEBPACK_IMPORTED_MODULE_3__app_globals__["monthNames"][mo].substr(0, 3) + " - " + __WEBPACK_IMPORTED_MODULE_3__app_globals__["monthNames"][mo + mos - 1].substr(0, 3) + ' ' + yr;
        }
        return ret;
    };
    AppComponent.prototype.mouseEvent = function (e) {
        if (e.type == "mouseenter") {
            this.oplandb.showHilights = false;
        }
    };
    /** Table development objects */
    AppComponent.prototype.devInit = function () {
        //console.log(this.devDat.slice(1,4))
        //console.log(this.devDat);
    };
    AppComponent.prototype.tblScroll = function (e) {
        var cont = __WEBPACK_IMPORTED_MODULE_2_jquery__(e.target);
        var head = cont.parent().find('thead');
        //$(head.find('th')[0]).css('margin-left',(-cont.scrollLeft())+'px');
        //$(head.find('th')[0]).cont.scrollLeft(cont.scrollLeft());
        head.scrollLeft(cont.scrollLeft());
    };
    AppComponent.prototype.testClick = function (e) {
        console.log('Test click from consumer page!');
        console.log(e);
    };
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["ViewChild"])(__WEBPACK_IMPORTED_MODULE_4__components_frm_chart_activities_frm_chart_activities_component__["a" /* FrmChartActivitiesComponent */]),
        __metadata("design:type", typeof (_a = typeof __WEBPACK_IMPORTED_MODULE_4__components_frm_chart_activities_frm_chart_activities_component__["a" /* FrmChartActivitiesComponent */] !== "undefined" && __WEBPACK_IMPORTED_MODULE_4__components_frm_chart_activities_frm_chart_activities_component__["a" /* FrmChartActivitiesComponent */]) === "function" && _a || Object)
    ], AppComponent.prototype, "calActivities", void 0);
    AppComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-root',
            template: __webpack_require__("../../../../../src/app/app.component.html"),
            //templateUrl: './app.component3.html',
            //templateUrl: './app.component6.html',
            //template:`<app-test-window-freeze></app-test-window-freeze>`,
            /*template:`
            <app-alv-table #myTable [headerTemplate]="sampleTable"></app-alv-table>
          
            <ng-template #sampleTable>
              <table class="testTable">
                <thead>
                  <tr>
                    <th>ID</th>
                    <th>Name</th>
                    <th>Name 2</th>
                  </tr>
                </thead>
                <tbody>
                  <tr *ngFor="let row of myTable.data">
                    <td>{{row.id}}</td>
                    <td>{{row.name}}</td>
                    <td>{{row.name}}</td>
                  </tr>
                </tbody>
              </table>
            </ng-template>
          
            `,*/
            styles: [__webpack_require__("../../../../../src/app/app.component.css")]
        }),
        __metadata("design:paramtypes", [typeof (_b = typeof __WEBPACK_IMPORTED_MODULE_5__app_udfs__["AppUDFService"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_5__app_udfs__["AppUDFService"]) === "function" && _b || Object, typeof (_c = typeof __WEBPACK_IMPORTED_MODULE_6__app_oplandata__["a" /* AppOplpanDataService */] !== "undefined" && __WEBPACK_IMPORTED_MODULE_6__app_oplandata__["a" /* AppOplpanDataService */]) === "function" && _c || Object])
    ], AppComponent);
    return AppComponent;
    var _a, _b, _c;
}());

//# sourceMappingURL=app.component.js.map

/***/ }),

/***/ "../../../../../src/app/app.globals.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
Object.defineProperty(__webpack_exports__, "__esModule", { value: true });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "datSourceUrl", function() { return datSourceUrl; });
/* harmony export (immutable) */ __webpack_exports__["getUrlRequest"] = getUrlRequest;
/* harmony export (immutable) */ __webpack_exports__["pad0"] = pad0;
/* harmony export (immutable) */ __webpack_exports__["padRS"] = padRS;
/* harmony export (immutable) */ __webpack_exports__["epochToDate"] = epochToDate;
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "monthNames", function() { return monthNames; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "monthIndex", function() { return monthIndex; });
/* harmony export (immutable) */ __webpack_exports__["getMonthName"] = getMonthName;
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "daysOfWeek", function() { return daysOfWeek; });
/* harmony export (immutable) */ __webpack_exports__["getDayOfWeek"] = getDayOfWeek;
/* harmony export (immutable) */ __webpack_exports__["stringToDate"] = stringToDate;
/* harmony export (immutable) */ __webpack_exports__["setFocus"] = setFocus;
/* harmony export (immutable) */ __webpack_exports__["dateToString"] = dateToString;
/* harmony export (immutable) */ __webpack_exports__["monthDiff"] = monthDiff;
/* harmony export (immutable) */ __webpack_exports__["daysDiff"] = daysDiff;
/* harmony export (immutable) */ __webpack_exports__["addDays"] = addDays;
/* harmony export (immutable) */ __webpack_exports__["addDaysVal"] = addDaysVal;
/* harmony export (immutable) */ __webpack_exports__["shortToLongDate"] = shortToLongDate;
/* harmony export (immutable) */ __webpack_exports__["longToShortDate"] = longToShortDate;
/* harmony export (immutable) */ __webpack_exports__["longToDate"] = longToDate;
/* harmony export (immutable) */ __webpack_exports__["groupBy"] = groupBy;
/* harmony export (immutable) */ __webpack_exports__["sortBy"] = sortBy;
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "AppDataService", function() { return AppDataService; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "NavService", function() { return NavService; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__angular_http__ = __webpack_require__("../../../http/@angular/http.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2_rxjs_add_operator_map__ = __webpack_require__("../../../../rxjs/add/operator/map.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2_rxjs_add_operator_map___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_2_rxjs_add_operator_map__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3_jquery__ = __webpack_require__("../../../../jquery/dist/jquery.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3_jquery___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_3_jquery__);
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};




// main application data source url
var datSourceUrl = 'http://soga-alv/oplan/';
//export const datSourceUrl:string='http://localhost:52737/';
//export const datSourceUrl:string='http://localhost:50117/';
//http://localhost:50117/
//export const datSourceUrl:string='';
// returns the complete url part of the request to be made by the data service of a component
function getUrlRequest(service, method, params) {
    var loc = __WEBPACK_IMPORTED_MODULE_3_jquery__(location).attr('href').substr(0, 16).toLowerCase();
    //loc = loc.replace("http://","http://alv:@dmin02@");
    var url = (loc == 'http://localhost' ?
        datSourceUrl : '') + service + '.asmx/' + method + (params == undefined ? '' : '?' + params);
    //  let url:string="http://alv:@dmin02@192.168.32.149/oplandep/" + service + '.asmx/'+ method + (params==undefined?'':'?'+params);
    //console.log("@LOCATION!:",loc,"URL:",url);
    return url;
}
function pad0(num, chrs) {
    // up to 50 characters
    var retVal = ("0".repeat(50) + String(num));
    retVal = retVal.substr(retVal.length - chrs);
    return retVal;
}
function padRS(str, chrs) {
    // up to 255 characters
    if (chrs == undefined)
        chrs = 255;
    var retVal = (str + " ".repeat(255));
    retVal = retVal.substr(0, chrs);
    return retVal;
}
function epochToDate(epochDate, retLong) {
    if (retLong == undefined)
        retLong = false;
    var dt = new Date(parseFloat(epochDate.substr(6)));
    if (retLong) {
        return this.dateToString(dt, true);
    }
    else {
        return dt;
    }
}
var monthNames = [
    "January", "February", "March",
    "April", "May", "June",
    "July", "August", "September",
    "October", "November", "December"
];
var monthIndex = {
    "jan": 0, "feb": 1, "mar": 2,
    "apr": 3, "may": 4, "jun": 5,
    "jul": 6, "aug": 7, "sep": 8,
    "oct": 9, "nov": 10, "dec": 11
};
function getMonthName(monthIndex, short) {
    if (short == undefined)
        short = false;
    return short ? monthNames[monthIndex].substr(0, 3) : monthNames[monthIndex];
}
var daysOfWeek = [
    "Sunday", "Monday", "Tuesday", "Wednesday",
    "Thursday", "Friday", "Saturday"
];
function getDayOfWeek(dayIndex, short) {
    if (short == undefined)
        short = false;
    return short ? daysOfWeek[dayIndex].substr(0, 3) : daysOfWeek[dayIndex];
}
function stringToDate(strrDate) {
    if (strrDate == undefined) {
        return new Date();
    }
    else {
        return new Date(parseInt(strrDate.substr(0, 4)), parseInt(strrDate.substr(4, 2)) - 1, parseInt(strrDate.substr(6, 2)));
    }
}
function setFocus(ctlID) {
    setTimeout(function () { __WEBPACK_IMPORTED_MODULE_3_jquery__('#' + ctlID).focus(); }, 10);
}
function dateToString(dateObj, retLong) {
    if (retLong == undefined)
        retLong = false;
    var mo = String(dateObj.getMonth() + 1);
    var day = String(dateObj.getDate());
    if (retLong) {
        return (day.length == 1 ? "0" : "") + day + ' ' +
            String(monthNames[dateObj.getMonth()]).substr(0, 3) + ' ' +
            String(dateObj.getFullYear());
    }
    else {
        return String(dateObj.getFullYear()) +
            (mo.length == 1 ? "0" : "") + mo +
            (day.length == 1 ? "0" : "") + day;
    }
}
function monthDiff(d1, d2) {
    var months;
    months = (d2.getFullYear() - d1.getFullYear()) * 12;
    months -= d1.getMonth() + 1;
    months += d2.getMonth();
    return months <= 0 ? 0 : months;
}
function daysDiff(d1, d2) {
    var diff = Math.abs(d1.getTime() - d2.getTime());
    return Math.ceil(diff / (1000 * 3600 * 24));
}
function addDays(date, days) {
    var newDate = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    newDate.setDate(newDate.getDate() + parseInt(String(days)));
    return newDate;
}
function addDaysVal(date, days) {
    var dt = new Date(date.getFullYear(), date.getMonth(), date.getDay());
    dt.setDate(date.getDate() + parseInt(String(days)));
    return dt;
}
function shortToLongDate(yyyymmdd) {
    if (yyyymmdd.length == 0)
        return "";
    var dt = stringToDate(yyyymmdd);
    return dateToString(dt, true);
}
function longToShortDate(ddmmmyyyy) {
    if (ddmmmyyyy.length == 0)
        return "";
    var dArr = ddmmmyyyy.split(" ");
    if (dArr.length != 3)
        dArr = ddmmmyyyy.split("-"); // if date is hyphen(-) delimited
    var mo = "0" + (monthIndex[dArr[1].toLowerCase()] + 1);
    if (dArr[2].length == 2)
        dArr[2] = "20" + dArr[2]; // if year specitied is not 4 digits
    if (dArr[0].length == 1)
        dArr[0] = "0" + dArr[0]; // if year specitied is not 4 digits
    return dArr[2] + mo.substr(mo.length - 2, 2) + dArr[0];
}
function longToDate(ddmmmyyyy) {
    try {
        var strDate = this.longToShortDate(ddmmmyyyy) + "";
        if (strDate.length == 0) {
            return null;
        }
        else {
            return this.stringToDate(strDate);
        }
    }
    catch (e) {
        return null;
    }
}
//G.stringToDate(G.longToShortDate(String(this.udf.calDispStartDate)));
function groupBy(arr, key, objLookup, lookupKey, lookupFields, sortField) {
    var newArr = [], types = {}, newItem, i, j, cur;
    for (i = 0, j = arr.length; i < j; i++) {
        cur = arr[i];
        if (!(cur[key] in types)) {
            types[cur[key]] = { type: cur[key], data: [] };
            newArr.push(types[cur[key]]);
        }
        types[cur[key]].data.push(cur);
    }
    if (objLookup != undefined) {
        if (newArr) {
            newArr.forEach(function (item) {
                var obj = objLookup.find(function (lkpItem) { return item['type'] == lkpItem[lookupKey ? lookupKey : 'id']; });
                if (obj) {
                    item['name'] = obj['name'] != undefined ? obj['name'] : "unknown";
                }
                else {
                    item['name'] = "unknown";
                }
                if (lookupFields != undefined) {
                    // attach all fields from the lookup item to the current item of the input array
                    lookupFields.forEach(function (field) {
                        if (obj) {
                            item[field] = obj[field];
                        }
                        else {
                            item[field] = null;
                        }
                    });
                }
            });
        }
        if (sortField != undefined && sortBy != null) {
            this.sortBy(newArr, sortField);
        }
        else {
            this.sortBy(newArr, 'name');
        }
    }
    else {
        newArr.forEach(function (item) { item['name'] = item['type']; });
    }
    return newArr;
}
function sortBy(arr, key, desc) {
    if (arr == undefined)
        return;
    if (desc == undefined)
        desc = false;
    if (desc) {
        arr.sort(function (a, b) {
            return (a[key] < b[key]) ? 1 : ((b[key] < a[key]) ? -1 : 0);
        });
    }
    else {
        arr.sort(function (a, b) {
            return (a[key] > b[key]) ? 1 : ((b[key] > a[key]) ? -1 : 0);
        });
    }
}
var AppDataService = (function () {
    function AppDataService(http) {
        this.http = http;
    }
    AppDataService.prototype.getAppData = function (requestUrl, body, options) {
        var varBody = "";
        var varOptions = null;
        if (body != undefined)
            varBody = body;
        if (options != undefined)
            varOptions = options;
        return this.http
            .post(requestUrl, varBody, varOptions)
            .map(function (response) { return response.json(); });
    };
    AppDataService.prototype.getAppDataWS = function (service, method, params) {
        var url;
        if (params != undefined) {
            url = getUrlRequest(service, method) + '?' + params;
        }
        else {
            url = getUrlRequest(service, method);
        }
        //      alert(url);
        return this.getAppData(url);
    };
    AppDataService.prototype.saveAppDataWS = function (service, method, body, options) {
        var url;
        var varBody = "";
        var varOptions = undefined;
        if (body != undefined)
            varBody = body;
        if (options != undefined)
            varOptions = options;
        url = getUrlRequest(service, method);
        return this.getAppData(url, varBody, varOptions);
    };
    AppDataService = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Injectable"])(),
        __metadata("design:paramtypes", [typeof (_a = typeof __WEBPACK_IMPORTED_MODULE_1__angular_http__["a" /* Http */] !== "undefined" && __WEBPACK_IMPORTED_MODULE_1__angular_http__["a" /* Http */]) === "function" && _a || Object])
    ], AppDataService);
    return AppDataService;
    var _a;
}());

var NavService = (function () {
    function NavService() {
        this.navchange = new __WEBPACK_IMPORTED_MODULE_0__angular_core__["EventEmitter"]();
    }
    NavService.prototype.emitNavChangeEvent = function (e) {
        this.navchange.emit(e);
    };
    NavService.prototype.getNavChangeEmitter = function () {
        return this.navchange;
    };
    return NavService;
}());

//# sourceMappingURL=app.globals.js.map

/***/ }),

/***/ "../../../../../src/app/app.module.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return AppModule; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_platform_browser__ = __webpack_require__("../../../platform-browser/@angular/platform-browser.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__angular_http__ = __webpack_require__("../../../http/@angular/http.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3__angular_common_http__ = __webpack_require__("../../../common/@angular/common/http.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_4__angular_forms__ = __webpack_require__("../../../forms/@angular/forms.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_5_ng2_date_picker__ = __webpack_require__("../../../../ng2-date-picker/index.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_5_ng2_date_picker___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_5_ng2_date_picker__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_6__app_component__ = __webpack_require__("../../../../../src/app/app.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_7__components_nav_list_nav_list_component__ = __webpack_require__("../../../../../src/app/components/nav-list/nav-list.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_8__components_pnl_vert_split_pnl_vert_split_component__ = __webpack_require__("../../../../../src/app/components/pnl-vert-split/pnl-vert-split.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_9__components_frm_personnel_frm_personnel_component__ = __webpack_require__("../../../../../src/app/components/frm-personnel/frm-personnel.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_10__components_frm_team_frm_team_component__ = __webpack_require__("../../../../../src/app/components/frm-team/frm-team.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_11__components_frm_activity_frm_activity_component__ = __webpack_require__("../../../../../src/app/components/frm-activity/frm-activity.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_12__models_data_models__ = __webpack_require__("../../../../../src/app/models/data-models.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_13__app_udfs__ = __webpack_require__("../../../../../src/app/app.udfs.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_14__app_globals__ = __webpack_require__("../../../../../src/app/app.globals.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_15__app_oplandata__ = __webpack_require__("../../../../../src/app/app.oplandata.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_16__components_frm_chart_frm_chart_component__ = __webpack_require__("../../../../../src/app/components/frm-chart/frm-chart.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_17__components_test_panels_test_panels_component__ = __webpack_require__("../../../../../src/app/components/test-panels/test-panels.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_18__components_frm_table_panels_frm_table_panels_component__ = __webpack_require__("../../../../../src/app/components/frm-table-panels/frm-table-panels.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_19__components_ctl_user_select_ctl_user_select_component__ = __webpack_require__("../../../../../src/app/components/ctl-user-select/ctl-user-select.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_20__components_frm_settings_frm_settings_component__ = __webpack_require__("../../../../../src/app/components/frm-settings/frm-settings.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_21__components_frm_chart_activities_frm_chart_activities_component__ = __webpack_require__("../../../../../src/app/components/frm-chart-activities/frm-chart-activities.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_22__components_pobprofile_pobprofile_component__ = __webpack_require__("../../../../../src/app/components/pobprofile/pobprofile.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_23__components_temp_component_temp_component_component__ = __webpack_require__("../../../../../src/app/components/temp-component/temp-component.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_24__components_ctl_drop_down_ctl_drop_down_component__ = __webpack_require__("../../../../../src/app/components/ctl-drop-down/ctl-drop-down.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_25__components_test_counter_test_counter_component__ = __webpack_require__("../../../../../src/app/components/test-counter/test-counter.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_26__test_window_freeze_test_window_freeze_component__ = __webpack_require__("../../../../../src/app/test-window-freeze/test-window-freeze.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_27__components_report_header_report_header_component__ = __webpack_require__("../../../../../src/app/components/report-header/report-header.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_28__components_alv_table_alv_table_component__ = __webpack_require__("../../../../../src/app/components/alv-table/alv-table.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_29__directives_test_dir_directive__ = __webpack_require__("../../../../../src/app/directives/test-dir.directive.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_30__components_alv_table_panel_alv_table_panel_component__ = __webpack_require__("../../../../../src/app/components/alv-table-panel/alv-table-panel.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_31__components_test_master_test_master_component__ = __webpack_require__("../../../../../src/app/components/test-master/test-master.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_32__components_test_child_test_child_component__ = __webpack_require__("../../../../../src/app/components/test-child/test-child.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_33__components_our_table_our_table_component__ = __webpack_require__("../../../../../src/app/components/our-table/our-table.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_34__components_frm_chart_activities_mtiap_frm_chart_activities_mtiap_component__ = __webpack_require__("../../../../../src/app/components/frm-chart-activities-mtiap/frm-chart-activities-mtiap.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_35__components_alv_table_template_alv_table_template_component__ = __webpack_require__("../../../../../src/app/components/alv-table-template/alv-table-template.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_36__components_alv_table_template_header_alv_table_template_header_component__ = __webpack_require__("../../../../../src/app/components/alv-table-template-header/alv-table-template-header.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_37__components_alv_table_template_filter_alv_table_template_filter_component__ = __webpack_require__("../../../../../src/app/components/alv-table-template-filter/alv-table-template-filter.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_38__components_alv_table_template_detail_alv_table_template_detail_component__ = __webpack_require__("../../../../../src/app/components/alv-table-template-detail/alv-table-template-detail.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_39__components_alv_table_template_body_alv_table_template_body_component__ = __webpack_require__("../../../../../src/app/components/alv-table-template-body/alv-table-template-body.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_40__components_alv_table_template_detail_row_alv_table_template_detail_row_component__ = __webpack_require__("../../../../../src/app/components/alv-table-template-detail-row/alv-table-template-detail-row.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_41__components_alv_table_template_header_row_alv_table_template_header_row_component__ = __webpack_require__("../../../../../src/app/components/alv-table-template-header-row/alv-table-template-header-row.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_42__components_alv_test_component_alv_test_component_component__ = __webpack_require__("../../../../../src/app/components/alv-test-component/alv-test-component.component.ts");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};












































var AppModule = (function () {
    function AppModule() {
    }
    AppModule = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_1__angular_core__["NgModule"])({
            declarations: [
                __WEBPACK_IMPORTED_MODULE_6__app_component__["a" /* AppComponent */],
                __WEBPACK_IMPORTED_MODULE_7__components_nav_list_nav_list_component__["a" /* NavListComponent */],
                __WEBPACK_IMPORTED_MODULE_8__components_pnl_vert_split_pnl_vert_split_component__["a" /* PnlVertSplitComponent */],
                __WEBPACK_IMPORTED_MODULE_9__components_frm_personnel_frm_personnel_component__["a" /* FrmPersonnelComponent */],
                __WEBPACK_IMPORTED_MODULE_10__components_frm_team_frm_team_component__["a" /* FrmTeamComponent */],
                __WEBPACK_IMPORTED_MODULE_11__components_frm_activity_frm_activity_component__["a" /* FrmActivityComponent */],
                __WEBPACK_IMPORTED_MODULE_16__components_frm_chart_frm_chart_component__["a" /* FrmChartComponent */],
                __WEBPACK_IMPORTED_MODULE_17__components_test_panels_test_panels_component__["a" /* TestPanelsComponent */],
                __WEBPACK_IMPORTED_MODULE_18__components_frm_table_panels_frm_table_panels_component__["a" /* FrmTablePanelsComponent */],
                __WEBPACK_IMPORTED_MODULE_19__components_ctl_user_select_ctl_user_select_component__["a" /* CtlUserSelectComponent */],
                __WEBPACK_IMPORTED_MODULE_20__components_frm_settings_frm_settings_component__["a" /* FrmSettingsComponent */],
                __WEBPACK_IMPORTED_MODULE_21__components_frm_chart_activities_frm_chart_activities_component__["a" /* FrmChartActivitiesComponent */],
                __WEBPACK_IMPORTED_MODULE_22__components_pobprofile_pobprofile_component__["a" /* PobprofileComponent */],
                __WEBPACK_IMPORTED_MODULE_23__components_temp_component_temp_component_component__["a" /* TempComponentComponent */],
                __WEBPACK_IMPORTED_MODULE_24__components_ctl_drop_down_ctl_drop_down_component__["a" /* CtlDropDownComponent */],
                __WEBPACK_IMPORTED_MODULE_25__components_test_counter_test_counter_component__["a" /* TestCounterComponent */],
                __WEBPACK_IMPORTED_MODULE_26__test_window_freeze_test_window_freeze_component__["a" /* TestWindowFreezeComponent */],
                __WEBPACK_IMPORTED_MODULE_27__components_report_header_report_header_component__["a" /* ReportHeaderComponent */],
                __WEBPACK_IMPORTED_MODULE_28__components_alv_table_alv_table_component__["a" /* AlvTableComponent */],
                __WEBPACK_IMPORTED_MODULE_29__directives_test_dir_directive__["a" /* TestDirDirective */],
                __WEBPACK_IMPORTED_MODULE_30__components_alv_table_panel_alv_table_panel_component__["a" /* AlvTablePanelComponent */],
                __WEBPACK_IMPORTED_MODULE_31__components_test_master_test_master_component__["a" /* TestMasterComponent */],
                __WEBPACK_IMPORTED_MODULE_32__components_test_child_test_child_component__["a" /* TestChildComponent */],
                __WEBPACK_IMPORTED_MODULE_33__components_our_table_our_table_component__["c" /* OurTableComponent */],
                __WEBPACK_IMPORTED_MODULE_33__components_our_table_our_table_component__["d" /* OurTableHeaderComponent */],
                __WEBPACK_IMPORTED_MODULE_33__components_our_table_our_table_component__["a" /* OurTableBodyComponent */],
                __WEBPACK_IMPORTED_MODULE_34__components_frm_chart_activities_mtiap_frm_chart_activities_mtiap_component__["a" /* FrmChartActivitiesMtiapComponent */],
                __WEBPACK_IMPORTED_MODULE_35__components_alv_table_template_alv_table_template_component__["a" /* AlvTableTemplateComponent */],
                __WEBPACK_IMPORTED_MODULE_36__components_alv_table_template_header_alv_table_template_header_component__["a" /* AlvTableTemplateHeaderComponent */],
                __WEBPACK_IMPORTED_MODULE_37__components_alv_table_template_filter_alv_table_template_filter_component__["a" /* AlvTableTemplateFilterComponent */],
                __WEBPACK_IMPORTED_MODULE_38__components_alv_table_template_detail_alv_table_template_detail_component__["a" /* AlvTableTemplateDetailComponent */],
                __WEBPACK_IMPORTED_MODULE_39__components_alv_table_template_body_alv_table_template_body_component__["a" /* AlvTableTemplateBodyComponent */],
                __WEBPACK_IMPORTED_MODULE_40__components_alv_table_template_detail_row_alv_table_template_detail_row_component__["a" /* AlvTableTemplateDetailRowComponent */],
                __WEBPACK_IMPORTED_MODULE_41__components_alv_table_template_header_row_alv_table_template_header_row_component__["a" /* AlvTableTemplateHeaderRowComponent */],
                __WEBPACK_IMPORTED_MODULE_42__components_alv_test_component_alv_test_component_component__["a" /* AlvTestComponentComponent */]
            ],
            imports: [
                __WEBPACK_IMPORTED_MODULE_0__angular_platform_browser__["a" /* BrowserModule */],
                __WEBPACK_IMPORTED_MODULE_5_ng2_date_picker__["DpDatePickerModule"],
                __WEBPACK_IMPORTED_MODULE_2__angular_http__["b" /* HttpModule */],
                __WEBPACK_IMPORTED_MODULE_3__angular_common_http__["b" /* HttpClientModule */],
                __WEBPACK_IMPORTED_MODULE_4__angular_forms__["FormsModule"]
            ],
            providers: [
                __WEBPACK_IMPORTED_MODULE_12__models_data_models__["d" /* dataPersonnel */],
                __WEBPACK_IMPORTED_MODULE_12__models_data_models__["c" /* dataActivity */],
                __WEBPACK_IMPORTED_MODULE_12__models_data_models__["e" /* dataTeam */],
                __WEBPACK_IMPORTED_MODULE_33__components_our_table_our_table_component__["b" /* OurTableColumnDef */],
                __WEBPACK_IMPORTED_MODULE_12__models_data_models__["g" /* lookupTemplate */],
                __WEBPACK_IMPORTED_MODULE_13__app_udfs__["AppUDFService"],
                __WEBPACK_IMPORTED_MODULE_15__app_oplandata__["a" /* AppOplpanDataService */],
                __WEBPACK_IMPORTED_MODULE_14__app_globals__["AppDataService"]
            ],
            bootstrap: [__WEBPACK_IMPORTED_MODULE_6__app_component__["a" /* AppComponent */]]
        })
    ], AppModule);
    return AppModule;
}());

//# sourceMappingURL=app.module.js.map

/***/ }),

/***/ "../../../../../src/app/app.oplandata.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return AppOplpanDataService; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__angular_http__ = __webpack_require__("../../../http/@angular/http.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__angular_common_http__ = __webpack_require__("../../../common/@angular/common/http.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3__models_data_models__ = __webpack_require__("../../../../../src/app/models/data-models.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_4__app_globals__ = __webpack_require__("../../../../../src/app/app.globals.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_5__app_udfs__ = __webpack_require__("../../../../../src/app/app.udfs.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_6_jquery__ = __webpack_require__("../../../../jquery/dist/jquery.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_6_jquery___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_6_jquery__);
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};







var AppOplpanDataService = (function () {
    function AppOplpanDataService(http, http2, udf) {
        this.http = http;
        this.http2 = http2;
        this.udf = udf;
        this.showHilights = false;
        this.isCtrlKey = false;
        this.isEditing = false;
        this.getActiveActivities = true;
        this.hideIds = true;
        this.isLoading = false;
        this.isLoadingMessage = 'Loading. Please wait...';
        this.isPosting = false;
        this.isPostingMessage = 'Saving data. Please wait...';
        this.winWidth = 0;
        this.winHeight = 0;
        this.getMessageTop = 0;
        this.getMessageLeft = -500;
        this.onActivityLoadAction = true;
        this.msgTempPrompt = '';
        this.msgTempPromptModal = '';
        this.msgTempInfoPrompt = '';
        this.msgRequestPrompt = '';
        this.currConfig = null;
        this.minDate = "20170101";
        this.maxDate = "20221231";
        this.activityModifiedInCalendar = false;
        this.config = {
            stdPOBLimit: {
                title: "Standard POB Limit",
                limits: [],
                newLimit: { setDate: "", value: "" }
            },
            upmPOBLimit: {
                title: "Upmanning POB Limit",
                limits: [],
                newLimit: { setDate: "", value: "" }
            },
            upmanPeriods: {
                title: "Upmanning Period",
                periods: [],
                newPeriod: { startDate: "", endDate: "" }
            },
            mobLimit: {
                title: "Mobilization Limit",
                limits: [],
                newLimit: { setDate: "", value: "" }
            },
            demobLimit: {
                title: "Demobilization Limit",
                limits: [],
                newLimit: { setDate: "", value: "" }
            },
            calStart: this.udf.calStartDate,
            calEnd: this.udf.calEndDate,
            stdPOBValues: [],
            upmPOBValues: [],
            actualPOBValues: [],
            // and takes value from upmValue if within any of the upmPeriods
            mobValues: [],
            demobValues: [],
            teamGroupsArr: [],
            idxTeamGroups: {},
            activityCounts: {},
            noFlights: [],
            arranged: false
        };
        this.msgPersonUpdateSuccess = '';
        this.msgPersonUpdateInstruction = '';
        this.msgPersonUpdateError = '';
        this.msgPersonUpdateConfirm = '';
        this.msgActivityUpdateSuccess = '';
        this.msgActivityUpdateInstruction = '';
        this.msgActivityUpdateError = '';
        this.msgActivityUpdateConfirm = '';
        this.logInfo = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["f" /* loginDetails */]();
        this.alertYesNoMessage = {
            label: 'WARNING',
            message: 'You are about to delete the current record and related items. Do you want to proceed?',
            pos: 'Yes',
            neg: 'No'
        };
        this.userSelection = [
            { name: "A", data: [
                    { id: 1, fullName: "Personnel 1" },
                    { id: 2, fullName: "Personnel 2" },
                    { id: 3, fullName: "Personnel 3" }
                ] }
        ];
        this.currPersonnelObj = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["d" /* dataPersonnel */]();
        this.currTeamObj = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["e" /* dataTeam */]();
        this.currActivityObj = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["c" /* dataActivity */]();
        this.currPersonnel = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["d" /* dataPersonnel */]();
        this.currActivity = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["c" /* dataActivity */]();
        this.currTeam = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["e" /* dataTeam */]();
        this.newActivityMember = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["b" /* activityMember */]();
        this.currDelObj = null;
        this.datCompany = [];
        this.idxCompany = {};
        this.datPosition = [];
        this.idxPosition = {};
        this.idxPersonnel = {};
        this.datPersonnel = [];
        this.idxActivity = {};
        this.datActivities = [];
        this.datCertificates = [
            { id: 0, name: "", required: false, issued: null, expiry: null, current: false, mandatory: false, default: false, frequency: 0 },
            { id: 1, name: "FTW Clearance-Offshore", required: true, issued: "", expiry: "", current: false, mandatory: true, default: false, frequency: 1 },
            { id: 9, name: "FTW Clearance-Onshore", required: false, issued: "", expiry: "", current: false, mandatory: false, default: false, frequency: 4 },
            { id: 2, name: "BOSIET/FOET", required: true, issued: "", expiry: "", current: false, mandatory: false, default: false, frequency: 4 },
            { id: 3, name: "HUET", required: false, issued: null, expiry: null, current: false, mandatory: false, default: false, frequency: 0 },
            { id: 4, name: "HUET Dispensation", required: false, issued: null, expiry: null, current: false, mandatory: false, default: false, frequency: 0 },
            { id: 5, name: "PTW Users Training", required: true, issued: null, expiry: null, current: false, mandatory: true, default: false, frequency: 0 },
            { id: 6, name: "PICWS Training", required: false, issued: null, expiry: null, current: false, mandatory: false, default: false, frequency: 1 },
            { id: 7, name: "Malaria Briefing", required: false, issued: null, expiry: null, current: false, mandatory: false, default: false, frequency: 1 },
            { id: 8, name: "Donut Training", required: false, issued: null, expiry: null, current: false, mandatory: false, default: false, frequency: 0 },
        ];
        this.idxTeam = {};
        this.datTeams = [];
        this.datPriorities = [];
        this.lkpSubTeams = {};
        this.lkpActivityName = {};
        this.tmpArray = [];
        this.expTeams = [];
        this.currActivityMember = null;
        this.currActivityComp = [];
    }
    AppOplpanDataService.prototype.resetCerts = function () {
        this.datCertificates.forEach(function (c) {
            c.required = false;
            c.current = false;
            c.issued = null;
            c.expiry = null;
        });
    };
    AppOplpanDataService.prototype.getConfig = function (onProcessed) {
        var _this = this;
        this.setLoading("Loading configurations. Please wait...");
        this.clearSubscription(this.subsConfig);
        this.subsConfig = this.http
            .get(__WEBPACK_IMPORTED_MODULE_4__app_globals__["getUrlRequest"]("oplanData", "getConfig"))
            .subscribe(function (data) {
            _this.processConfig(data);
            _this.noLoading();
            // if after processing event is supplied, execute the event
            if (onProcessed != undefined) {
                onProcessed();
            }
        }, function (err) {
            _this.withError(err);
        });
        return this.subsConfig;
    };
    AppOplpanDataService.prototype.processConfig = function (data) {
        //this.config['stdPOBLimit']['limits'].sort()
        var _this = this;
        this.config['stdPOBLimit']['limits'] = [];
        this.config['upmPOBLimit']['limits'] = [];
        this.config['mobLimit']['limits'] = [];
        this.config['demobLimit']['limits'] = [];
        this.config['upmanPeriods']['periods'] = [];
        this.config['idxTeamGroups'] = {};
        this.config['teamGroupsArr'] = [];
        this.config['noFlights'] = [];
        console.log(data);
        //
        data['noflight'].forEach(function (item) {
            _this.config['noFlights'].push({ id: item[0], date: item[1], notes: item[2] });
        });
        data['teamGroups'].forEach(function (item) {
            var cfg = {
                //id:Number(item[0]),
                id: item[0],
                name: item[1]
            };
            _this.config['teamGroupsArr'].push(cfg);
            _this.config['idxTeamGroups']['r' + item[0]] = cfg;
        });
        data['stdPOBLimit'].forEach(function (item) {
            var cfg = {
                id: Number(item[0]),
                setDate: item[1].length == 0 ? new Date(1980, 0, 1) : __WEBPACK_IMPORTED_MODULE_4__app_globals__["stringToDate"](item[1]),
                value: Number(item[2])
            };
            _this.config['stdPOBLimit']['limits'].push(cfg);
        });
        data['upmPOBLimit'].forEach(function (item) {
            var cfg = {
                id: Number(item[0]),
                setDate: item[1].length == 0 ? new Date(1980, 0, 1) : __WEBPACK_IMPORTED_MODULE_4__app_globals__["stringToDate"](item[1]),
                value: Number(item[2])
            };
            _this.config['upmPOBLimit']['limits'].push(cfg);
        });
        data['mobLimit'].forEach(function (item) {
            var cfg = {
                id: Number(item[0]),
                setDate: item[1].length == 0 ? new Date(1980, 0, 1) : __WEBPACK_IMPORTED_MODULE_4__app_globals__["stringToDate"](item[1]),
                value: Number(item[2])
            };
            _this.config['mobLimit']['limits'].push(cfg);
        });
        data['demobLimit'].forEach(function (item) {
            var cfg = {
                id: Number(item[0]),
                setDate: item[1].length == 0 ? new Date(1980, 0, 1) : __WEBPACK_IMPORTED_MODULE_4__app_globals__["stringToDate"](item[1]),
                value: Number(item[2])
            };
            _this.config['demobLimit']['limits'].push(cfg);
        });
        // {id:1, startDate:new Date(2017,3,15),endDate:new Date(2017,3,25)}
        data['upmanPeriods'].forEach(function (item) {
            var cfg = {
                id: Number(item[0]),
                startDate: item[1].length == 0 ? new Date(1980, 0, 1) : __WEBPACK_IMPORTED_MODULE_4__app_globals__["stringToDate"](item[1]),
                endDate: item[2].length == 0 ? new Date(1980, 0, 1) : __WEBPACK_IMPORTED_MODULE_4__app_globals__["stringToDate"](item[2])
            };
            _this.config['upmanPeriods']['periods'].push(cfg);
        });
        /**
         *   datCertificates:Array<any>=[
        {id:0, name:"" , required:false,issued:null,expiry:null,current:false, mandatory:false, frequency:0},
        {id:1, name:"FTW Clearance" , required:true,issued:"",expiry:"",current:false, mandatory:true, frequency:1},
        {id:2, name:"BOSIET/FOET" , required:true,issued:"",expiry:"",current:false, mandatory:true, frequency:4},
        {id:3, name:"HUET" , required:false,issued:null,expiry:null,current:false, mandatory:false,frequency:0},
        {id:4, name:"HUET Dispensation" , required:false,issued:null,expiry:null,current:false, mandatory:false, frequency:0},
        {id:5, name:"PTW Users Training" , required:true,issued:null,expiry:null,current:false, mandatory:true, frequency:0},
        {id:6, name:"PICWS Training" , required:false,issued:null,expiry:null,current:false, mandatory:false, frequency:1},
        {id:7, name:"Malaria Briefing" , required:false,issued:null,expiry:null,current:false, mandatory:false, frequency:1},
        {id:8, name:"Donut Training" , required:false,issued:null,expiry:null,current:false, mandatory:false, frequency:0},
      ];
    
         *
         */
        this.datCertificates = [{ id: 0, name: "", required: false, issued: null, expiry: null, current: false, mandatory: false, frequency: 0 }];
        data['premobCerts'].forEach(function (item) {
            _this.datCertificates.push({
                id: item[0],
                name: item[1],
                required: item[4],
                issued: null,
                expiry: null,
                current: false,
                mandatory: item[3],
                default: item[4],
                frequency: item[2]
            });
        });
        console.log('this.datCertificates:');
        console.log(this.datCertificates);
        //premobCerts
        this.arrangeConfig();
    };
    AppOplpanDataService.prototype.arrangeConfig = function () {
        // to prevent processing when data are not yet available...
        if (this.config['stdPOBLimit']['limits'].length == 0)
            return;
        // sort all configuration items to make the first element as the latest entry
        __WEBPACK_IMPORTED_MODULE_4__app_globals__["sortBy"](this.config['stdPOBLimit']['limits'], 'setDate', true);
        __WEBPACK_IMPORTED_MODULE_4__app_globals__["sortBy"](this.config['upmPOBLimit']['limits'], 'setDate', true);
        __WEBPACK_IMPORTED_MODULE_4__app_globals__["sortBy"](this.config['mobLimit']['limits'], 'setDate', true);
        __WEBPACK_IMPORTED_MODULE_4__app_globals__["sortBy"](this.config['demobLimit']['limits'], 'setDate', true);
        __WEBPACK_IMPORTED_MODULE_4__app_globals__["sortBy"](this.config['upmanPeriods']['periods'], 'startDate', true);
        __WEBPACK_IMPORTED_MODULE_4__app_globals__["sortBy"](this.config['noFlights'], 'date', true);
        //let days:number=G.daysDiff(this.config.calStart,this.config.calEnd);
        // calculate the number of days between the startDate and endDate selected
        // in the POB calendar
        var days = __WEBPACK_IMPORTED_MODULE_4__app_globals__["daysDiff"](this.udf.calStartDate, this.udf.calEndDate);
        var sd = __WEBPACK_IMPORTED_MODULE_4__app_globals__["stringToDate"](__WEBPACK_IMPORTED_MODULE_4__app_globals__["dateToString"](this.udf.calStartDate));
        var ctr;
        var cdt;
        var curLimitArr = null;
        // remove all elements from configuration types
        this.config.stdPOBValues = [];
        this.config.upmPOBValues = [];
        this.config.actualPOBValues = [];
        this.config.mobValues = [];
        this.config.demobValues = [];
        for (ctr = 0; ctr <= days; ctr++) {
            cdt = __WEBPACK_IMPORTED_MODULE_4__app_globals__["addDays"](sd, ctr);
            curLimitArr = this.config.stdPOBLimit.limits.filter(function (item) {
                return (item.setDate <= cdt);
            });
            this.config.stdPOBValues.push(curLimitArr[0]['value']);
            curLimitArr = this.config.upmPOBLimit.limits.filter(function (item) {
                return (item.setDate <= cdt);
            });
            this.config.upmPOBValues.push(curLimitArr[0]['value']);
            curLimitArr = this.config.mobLimit.limits.filter(function (item) {
                return (item.setDate <= cdt);
            });
            this.config.mobValues.push(curLimitArr[0]['value']);
            curLimitArr = this.config.demobLimit.limits.filter(function (item) {
                return (item.setDate <= cdt);
            });
            this.config.demobValues.push(curLimitArr[0]['value']);
            // upmanning on...
            curLimitArr = this.config.upmanPeriods.periods.filter(function (item) {
                return (cdt >= item.startDate && cdt <= item.endDate);
            });
            if (curLimitArr.length == 0) {
                // standard POB only
                this.config.actualPOBValues.push(this.config.stdPOBValues[this.config.stdPOBValues.length - 1]);
            }
            else {
                // upmanning POB active
                this.config.actualPOBValues.push(this.config.upmPOBValues[this.config.upmPOBValues.length - 1]);
            }
        } // for (...)
        this.config.arranged = true;
    };
    AppOplpanDataService.prototype.isNoFlight = function (shortDate) {
        if (shortDate.indexOf(" ") != -1) {
            shortDate = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](shortDate);
        }
        else if (shortDate.indexOf("-") != -1) {
            //item.date
            var dArr = shortDate.split('-');
            if (dArr[2].length == 2)
                dArr[2] = "20" + dArr[2]; // from 2 char year to 4 chars
            dArr[0] = "00" + dArr[0];
            dArr[0] = dArr[0].substr(dArr[0].length - 2);
            shortDate = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](dArr[0] + " " + dArr[1] + " " + dArr[2]);
        }
        var obj = this.config['noFlights'].find(function (item) { return item.date == shortDate; });
        if (!obj) {
            return null;
        }
        else {
            return obj;
        }
    };
    AppOplpanDataService.prototype.isActMemNoFlight = function (mem, testDemob) {
        var ret = null;
        if (testDemob == undefined)
            testDemob = true; // by default testDemob is true
        if (testDemob && mem.demob) {
            var nextDay = __WEBPACK_IMPORTED_MODULE_4__app_globals__["addDays"](__WEBPACK_IMPORTED_MODULE_4__app_globals__["stringToDate"](__WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](mem.endDate)), 1);
            if (mem.startDate == mem.endDate) {
                // same start date and end date
                if (mem.isDay) {
                    // test the endDate as demob date
                    ret = this.isNoFlight(mem.endDate);
                }
                else {
                    // test the endDate + 1 as demob date
                    //addDays(date: Date, days: number): Date
                    ret = this.isNoFlight(__WEBPACK_IMPORTED_MODULE_4__app_globals__["dateToString"](nextDay));
                }
            }
            else {
                // test the endDate + 1 as demob date
                ret = this.isNoFlight(__WEBPACK_IMPORTED_MODULE_4__app_globals__["dateToString"](nextDay));
            }
        }
        else if (!testDemob && mem.mob) {
            ret = this.isNoFlight(mem.startDate);
        }
        var nof = false;
        var nofDate = "";
        if (ret == null) {
        }
        else {
            nof = true;
            nofDate = __WEBPACK_IMPORTED_MODULE_4__app_globals__["shortToLongDate"](ret['date']);
        }
        return { noFlight: nof, date: nofDate };
    };
    AppOplpanDataService.prototype.getConfigObject = function (type, idx) {
        //let ret:any = this.oplandb.
        if (idx == undefined)
            idx = -1;
        var cfgName = "";
        switch (type) {
            case 'SPOB':
                cfgName = "stdPOBLimit";
                break;
            case 'UPOB':
                cfgName = "upmPOBLimit";
                break;
            case 'UPER':
                cfgName = "upmanPeriods";
                break;
            case 'MLIM':
                cfgName = "mobLimit";
                break;
            case 'DLIM':
                cfgName = "demobLimit";
                break;
        }
        if (idx == -1) {
            return this.config[cfgName];
        }
        else {
            return this.config[cfgName][type == 'UPER' ? 'periods' : 'limits'][idx];
        }
    };
    AppOplpanDataService.prototype.addConfigValue = function (type) {
        var _this = this;
        this.currConfig = this.getConfigObject(type);
        var cfg = null;
        if (this.currConfig != null)
            cfg = this.currConfig[type == 'UPER' ? 'newPeriod' : 'newLimit'];
        if (cfg != null) {
            // proceed with data validation and posting of new value
            if (type == 'UPER') {
                if (cfg['startDate'] == "" || cfg['endDate'] == "") {
                    this.setTempMessage(undefined, "Please enter valid Upmanning period Start and End dates ...", 4000);
                    return undefined;
                }
                else {
                    if (__WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](cfg['startDate']) > __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](cfg['endDate'])) {
                        this.setTempMessage(undefined, "Start date cannot be later than the End date ...", 4000);
                        return undefined;
                    }
                }
            }
            else {
                if (cfg['setDate'] == "") {
                    this.setTempMessage(undefined, "Please enter a valid " + this.currConfig['title'] + " set date ...", 4000);
                    return undefined;
                }
                if (cfg['value'] == "" || cfg['value'] <= 0 || isNaN(cfg['value'])) {
                    this.setTempMessage(undefined, "Please enter a valid " + this.currConfig['title'] + " value (i.e. greater than 0) ...", 4000);
                    return undefined;
                }
            }
            var data = {
                "type": type,
                "setDate": type == "UPER" ? "" : __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](cfg['setDate']),
                "value": type == "UPER" ? 0 : cfg['value'],
                "startDate": type != "UPER" ? "" : __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](cfg['startDate']),
                "endDate": type != "UPER" ? "" : __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](cfg['endDate']),
            };
            this.isPosting = true;
            this.msgRequestPrompt = "Saving new entry for " + this.currConfig['title'] + ". Please wait...";
            this.subsProcessConfig = this.processData({
                method: "saveNewConfig",
                data: {
                    "type": type,
                    "setDate": type == "UPER" ? "" : __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](cfg['setDate']),
                    "value": type == "UPER" ? 0 : cfg['value'],
                    "startDate": type != "UPER" ? "" : __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](cfg['startDate']),
                    "endDate": type != "UPER" ? "" : __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](cfg['endDate']),
                },
                isPosting: true,
                processMessage: this.msgRequestPrompt,
                subscription: this.subsProcessConfig,
                onSuccess: function (data) {
                    _this.msgRequestPrompt = "";
                    //this.noPosting();
                    //this.msgRequestPrompt="";
                    /*this.setTempMessage("msgTempInfoPrompt",
                      "New configuration setting was succesfully created under " + this.currConfig['title']);
          
                    let cfgArr:Array<object> = this.currConfig[type=="UPER" ? 'periods' : 'limits'];
                    let newCfg:object={id:data['resInt32']};
          
                    if(type=="UPER"){
                      newCfg['startDate']=G.longToDate(cfg['startDate']);
                      newCfg['endDate']=G.longToDate(cfg['endDate']);
                      cfg['startDate']="";
                      cfg['endDate']="";
                    }else{
                      newCfg['setDate']=G.longToDate(cfg['setDate']);
                      newCfg['value']=G.longToDate(cfg['value']);
                      cfg['setDate']="";
                      cfg['value']="";
                    }
          
          
                    cfgArr.push(newCfg);
          
                    this.arrangeConfig();
          
                    this.isPosting=false;*/
                    var cfgArr = _this.currConfig[type == "UPER" ? 'periods' : 'limits'];
                    var newCfg = { id: data['resInt32'] };
                    if (type == "UPER") {
                        newCfg['startDate'] = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](cfg['startDate']);
                        newCfg['endDate'] = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](cfg['endDate']);
                        cfg['startDate'] = "";
                        cfg['endDate'] = "";
                    }
                    else {
                        newCfg['setDate'] = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](cfg['setDate']);
                        newCfg['value'] = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](cfg['value']);
                        cfg['setDate'] = "";
                        cfg['value'] = "";
                    }
                    _this.getConfig();
                }
            });
        }
        return this.subsProcessConfig;
    };
    AppOplpanDataService.prototype.recordExpandedTeams = function () {
        var _this = this;
        this.expTeams = [];
        this.datTeams.forEach(function (team) {
            if (team.expView)
                _this.expTeams.push(team.id);
        });
    };
    AppOplpanDataService.prototype.setExpandedTeams = function () {
        var _this = this;
        this.expTeams.forEach(function (teamID) {
            var team = _this.idxTeam['r' + teamID];
            if (team != undefined)
                team.expView = true;
        });
    };
    AppOplpanDataService.prototype.scatterData = function (item) {
        if (this.udf.currentModule == 'Teams') {
            this.scatterTeam(item);
        }
        else if (this.udf.currentModule == 'Activities') {
            this.scatterActivity(item);
        }
        else {
            this.scatterPersonnel(item);
        }
    };
    AppOplpanDataService.prototype.scatterTeam = function (dat) {
        if (this.udf.modeTeam != "normal")
            return;
        if (dat != undefined) {
            this.currTeamObj = dat;
        }
        else {
            dat = this.currTeamObj;
        }
        this.currTeam.id = dat.id;
        this.currTeam.name = dat.name;
        this.currTeam.order = dat.order;
        this.currTeam.group = dat.group;
        this.currTeam.description = dat.description;
        this.currTeam.upman = dat.upman;
        this.currTeam.core = dat.core;
        this.currTeam.members = this.cloneObject(dat.members);
        this.setTeamUserSelection();
        this.isLoading = false;
    };
    AppOplpanDataService.prototype.scatterActivity = function (dat) {
        var _this = this;
        if (!this.udf.isActivityNormal())
            return; // do not invoke scatter while in edit or add mode.
        if (dat != undefined) {
            this.currActivityObj = dat;
        }
        else {
            dat = this.currActivityObj;
        }
        this.currActivity.id = dat.id;
        this.currActivity.name = dat.name;
        this.currActivity.description = dat.description;
        this.currActivity.startDate = dat.startDate;
        this.currActivity.endDate = dat.endDate;
        this.currActivity.type = dat.type;
        this.currActivity.showInChart = dat.showInChart;
        this.currActivity.noPOBCount = dat.noPOBCount;
        this.currActivity.ready = dat.ready;
        this.currActivity.parent = dat.parent;
        this.currActivity.members = [];
        this.currActivity.site = dat.site;
        this.currActivity.vessel = dat.vessel;
        this.currActivity.upmanning = dat.upmanning;
        this.currActivity.extracted = dat.extracted;
        this.currActivity.obsolete = dat.obsolete;
        this.currActivity.removed = dat.removed;
        // save activity id to local storage....
        localStorage.setItem("defActID", String(dat.id));
        dat.members.forEach(function (mem) {
            var newMem = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["b" /* activityMember */]();
            newMem.id = mem.id;
            newMem.idAct = mem.idAct;
            newMem.teamID = mem.teamID;
            newMem.perID = mem.perID;
            newMem.startDate = mem.startDate;
            newMem.endDate = mem.endDate;
            newMem.mob = mem.mob;
            newMem.demob = mem.demob;
            newMem.night = mem.night;
            newMem.isDay = mem.isDay;
            newMem.noPOB = mem.noPOB;
            newMem.coyID = mem.coyID;
            newMem.posID = mem.posID;
            newMem.grpID = mem.grpID;
            _this.currActivity.members.push(newMem);
        });
        this.getActivityConflicts();
    };
    AppOplpanDataService.prototype.showUpdatedActivityMember = function (data) {
    };
    AppOplpanDataService.prototype.showUpdatedActivity = function (data) {
        // this line must be placed at the beginning for scatter data to work
        // scatter data only works when record status is normal!!!
        this.udf.isActivityNormal(true);
        // if new record was created, append the new record
        var isNew = (data['action'] == 'AddActivity');
        if (isNew) {
            this.currActivityObj = this.currActivity;
            var newId = Number(data['resInt32']);
            this.currActivityObj.id = newId; // new id is stored in resInt32 field
            this.datActivities.push(this.currActivityObj);
            this.currActivity = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["c" /* dataActivity */]();
            this.scatterData(this.currActivityObj);
        }
        else {
            this.currActivityObj.name = this.currActivity.name;
            this.currActivityObj.description = this.currActivity.description;
            this.currActivityObj.startDate = this.currActivity.startDate;
            this.currActivityObj.endDate = this.currActivity.endDate;
        }
        this.setOtherNavGroup();
    };
    AppOplpanDataService.prototype.showUpdatedPersonnel = function (data) {
        var _this = this;
        var isNew = (data['action'] == 'AddPersonnel');
        if (isNew) {
            var newId = Number(data['resInt32']);
            this.currPersonnelObj = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["d" /* dataPersonnel */]();
            this.currPersonnelObj.id = newId; // new id is stored in resInt32 field
            // if ne personnel, create default certificate list
            this.currPersonnelObj.certificates = [];
            this.datCertificates.forEach(function (c) {
                if (c.default)
                    _this.currPersonnelObj.certificates.push({ id: c.id, req: true, issue: "", expiry: "" });
            });
        }
        setTimeout(function () {
            _this.currPersonnelObj.fullName = _this.currPersonnel.fullName;
            _this.currPersonnelObj.nameGroup = _this.currPersonnel.fullName.substr(0, 1).toUpperCase();
            _this.currPersonnelObj.email = _this.currPersonnel.email;
            _this.currPersonnelObj.companyID = _this.currPersonnel.companyID;
            _this.currPersonnelObj.positionID = _this.currPersonnel.positionID;
            _this.currPersonnelObj.gender = _this.currPersonnel.gender;
            _this.currPersonnelObj.birthDate = _this.currPersonnel.birthDate;
            _this.currPersonnelObj.xl = _this.currPersonnel.xl;
            if (isNew) {
                _this.datPersonnel.push(_this.currPersonnelObj);
                _this.scatterPersonnel();
            }
            // re-sort personnel name according to fullName then regroup
            __WEBPACK_IMPORTED_MODULE_4__app_globals__["sortBy"](_this.datPersonnel, "fullName");
            _this.setPersonnelNavGroup();
        }, 500); // set timeout to allow data refresh to take place before rendering
        this.setTempMessage("msgPersonUpdateSuccess", "Personnel data was successfully saved.");
        this.udf.modePersonnel = "normal";
    };
    AppOplpanDataService.prototype.scatterPersonnel = function (dat) {
        // remember selected personnel that will be edited
        if (this.udf.modePersonnel != "normal")
            return;
        if (dat != undefined) {
            this.currPersonnelObj = dat;
        }
        else {
            dat = this.currPersonnelObj;
        }
        var per = this.currPersonnel;
        per.id = dat.id;
        per.fullName = dat.fullName;
        per.email = dat.email;
        per.companyID = dat.companyID;
        per.positionID = dat.positionID;
        per.activities = dat.activities;
        per.maxDays = dat.maxDays;
        per.gender = dat.gender;
        per.birthDate = dat.birthDate;
        per.certificates = dat.certificates;
        per.xl = dat.xl;
        if (per.certificates) {
            this.scatterCerts(per.certificates);
        }
        else {
            // retrieve certs
            this.getPersonnelCerts(dat);
        }
        /*
            let certs:Array<any> = per.certificates;
            this.datCertificates.forEach((c:any)=>{
        
              c.current = false;
        
              let pc:any = certs.find((cert:any)=>{
                return cert.id == c.id;
              })
              if(pc){
                c.required = pc.req;
                c.issued = pc.issue ? pc.issue : null;
                c.expiry = pc.expiry ? pc.expiry : null;
                c.frequency = pc.frequency;
              }else{
                c.required = c.mandatory ? true : false;
                c.issued = null;
                c.expiry = null;
                c.frequency = 0;
              }
            });
        
            */
    };
    AppOplpanDataService.prototype.getPersonnelCerts = function (person) {
        var _this = this;
        // get personnel certificates on demand
        this.processData({
            method: "getPersonCertificates",
            data: {
                id: person.id,
            },
            isPosting: true,
            processMessage: "Retrieving certificates. Please wait...",
            onSuccess: function (certs) {
                if (certs.result == "success") {
                    person.certificates = [];
                    if (_this.currPersonnel)
                        _this.currPersonnel.certificates = [];
                    console.log(certs.certs);
                    certs.certs.forEach(function (cert) {
                        // set actual personnel object in the collection
                        person.certificates[person.certificates.length] = {
                            id: cert[0],
                            req: (cert[1] == "1"),
                            issue: __WEBPACK_IMPORTED_MODULE_4__app_globals__["shortToLongDate"](cert[2]),
                            expiry: __WEBPACK_IMPORTED_MODULE_4__app_globals__["shortToLongDate"](cert[3]),
                        };
                        // set scattered personnel object
                        if (_this.currPersonnel) {
                            _this.currPersonnel.certificates[_this.currPersonnel.certificates.length] = {
                                id: cert[0],
                                req: (cert[1] == "1"),
                                issue: __WEBPACK_IMPORTED_MODULE_4__app_globals__["shortToLongDate"](cert[2]),
                                expiry: __WEBPACK_IMPORTED_MODULE_4__app_globals__["shortToLongDate"](cert[3]),
                            };
                        }
                    });
                    _this.scatterCerts(person.certificates);
                }
            },
            onError: function (errObj) {
                console.log(errObj);
                alert(errObj.message);
            }
        });
    };
    AppOplpanDataService.prototype.scatterCerts = function (certs) {
        this.datCertificates.forEach(function (c) {
            c.current = false;
            var pc = certs.find(function (cert) {
                return cert.id == c.id;
            });
            if (pc) {
                //c.required = c.mandatory ? true :  pc.req;
                //c.required = c.default ? true :  pc.req;
                c.required = c.mandatory ? true : pc.req;
                c.issued = pc.issue ? pc.issue : null;
                c.expiry = pc.expiry ? pc.expiry : null;
                //c.frequency = pc.frequency;
            }
            else {
                //c.required = c.mandatory ? true : false;
                c.required = c.mandatory ? true : false;
                //c.required = c.default ? true : false;
                c.issued = null;
                c.expiry = null;
                //c.frequency = 0;
            }
        });
    };
    AppOplpanDataService.prototype.deletePersonnel = function (onDelete, beforeDelete, onError) {
        var _this = this;
        this.processData({
            method: "deletePersonnel",
            data: {
                id: this.currPersonnelObj.id,
                name: this.currPersonnelObj.fullName,
            },
            isPosting: true,
            processMessage: "Deleting personnel. Please wait...",
            onSuccess: function (data) {
                _this.udf.isPersonnelNormal(true);
                _this.setTempMessage("msgTempInfoPrompt", "Personnel '" + data['resString'] + "' has been deleted.");
                var idx = _this.datPersonnel.findIndex(function (item) {
                    return item.id == data['resInt32'];
                });
                _this.datPersonnel.splice(idx, 1);
                _this.setPersonnelNavGroup();
                _this.currPersonnel = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["d" /* dataPersonnel */];
                _this.currPersonnelObj = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["d" /* dataPersonnel */];
            }
        });
    };
    AppOplpanDataService.prototype.newTeam = function () {
        this.udf.isTeamNew(true);
        this.currTeam = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["e" /* dataTeam */]();
        if (this.currTeamObj) {
            this.currTeam.core = this.currTeamObj.core;
        }
    };
    AppOplpanDataService.prototype.saveTeamMember = function () {
        alert('saveTeamMember');
    };
    AppOplpanDataService.prototype.saveTeam = function () {
        var _this = this;
        var newTeam = this.currTeam.id == -1;
        // validate data
        var tm = this.currTeam;
        var err = '';
        if (tm.name == '') {
            this.setTempMessage(undefined, "Please enter a team name");
            //setTimeout(()=>$('#teamName').focus(),10);
            __WEBPACK_IMPORTED_MODULE_4__app_globals__["setFocus"]('teamName');
            return undefined;
        }
        tm.group = tm.core ? "Core" : "Non-Core";
        this.subsSaveTeam = this.processData({
            method: "saveTeam",
            data: {
                id: tm.id,
                name: tm.name,
                description: tm.description,
                order: tm.order,
                core: tm.core,
                upman: tm.upman
            },
            subscription: this.subsSaveTeam,
            isPosting: true,
            processMessage: "Saving team details. Please wait...",
            onSuccess: function (data) {
                //set
                //this.currTeam.order=data['resInt32']
                _this.udf.isTeamNormal(true);
                _this.setTempMessage("msgTempInfoPrompt", (newTeam ? "New team" : "Team " + tm.name) + " was successfully " + (newTeam ? " created." : " saved."));
                if (newTeam) {
                    tm.id = data['resInt32']; // save new id
                    tm.order = data['resInt32B']; // save new order
                    _this.currTeamObj = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["e" /* dataTeam */]();
                    _this.currTeamObj = _this.cloneObject(tm);
                }
                if (newTeam) {
                    // append new team to datTeams array
                    _this.datTeams.push(_this.currTeamObj);
                    _this.setOtherItemGroupings();
                }
                else {
                    var cTm = _this.currTeamObj;
                    cTm.name = tm.name;
                    cTm.description = tm.description;
                    cTm.core = tm.core;
                    cTm.upman = tm.upman;
                    cTm.group = tm.group;
                }
            }
        });
        return this.subsSaveTeam;
    };
    AppOplpanDataService.prototype.deleteTeam = function (data, go) {
        var _this = this;
        var postData = {};
        if (this.udf.isTeamDeleteAlert()) {
            var tm_1 = this.currTeamObj;
            postData.id = tm_1.id;
            postData.name = tm_1.name;
            this.subsDeleteTeam = this.processData({
                method: "deleteTeam",
                data: postData,
                isPosting: true,
                processMessage: "Deleting team '" + this.currTeamObj.name + "'. Please wait...",
                subscription: this.subsDeleteTeam,
                onSuccess: function (data) {
                    // set alerts
                    _this.udf.isTeamNormal(true);
                    _this.setTempMessage(undefined, "Team '" + _this.currTeamObj.name + "' has been deleted.");
                    // remove client-side data
                    var idx = _this.datTeams.findIndex(function (item) {
                        return item.id == tm_1.id;
                    });
                    _this.datTeams.splice(idx, 1);
                    _this.setOtherItemGroupings();
                    _this.currTeam = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["e" /* dataTeam */]();
                    _this.currTeamObj = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["e" /* dataTeam */]();
                    ;
                }
            });
            // delete entire team
        }
        else if (this.udf.isTeamMemberDeleteAlert()) {
            // delete team member
            var tmm = this.currDelObj;
            postData.id = tmm.id;
            postData.name = this.idxPersonnel['r' + tmm.perID]['fullName'];
            this.subsDeleteTeam = this.processData({
                method: "deleteTeamMember",
                data: postData,
                isPosting: true,
                subscription: this.subsDeleteTeam,
                processMessage: "Deleting team member '" + postData.name + "'. Please wait...",
                onSuccess: function (data) {
                    _this.udf.isTeamNormal(true);
                    _this.setTempMessage(undefined, "Team member '" + postData.name + "' was removed.");
                    // remove item from the cached member then clone to the current team object
                    var idx = _this.currTeam.members.findIndex(function (item) {
                        return item.perID == postData.id;
                    });
                    _this.currTeam.members.splice(idx, 1);
                    _this.currTeamObj.members = _this.cloneObject(_this.currTeam.members);
                }
            });
        }
        return this.subsDeleteTeam;
    };
    AppOplpanDataService.prototype.deleteTeamMember = function (item) {
        //this.oplandb.deleteTeamMember(item);
        this.udf.isTeamDeleteAlert(true);
    };
    AppOplpanDataService.prototype.cancelTeam = function () {
        this.currTeam = this.cloneObject(this.currTeamObj);
        this.udf.isTeamNormal(true);
    };
    AppOplpanDataService.prototype.newActivity = function () {
        this.currActivity = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["c" /* dataActivity */]();
        this.udf.isActivityNew(true);
    };
    AppOplpanDataService.prototype.saveActivityMember = function (onDataPost, beforePost, onError, fromCalendar, dataArray, saveModifiedAssignment) {
        var _this = this;
        if (saveModifiedAssignment == undefined)
            saveModifiedAssignment = false;
        if (beforePost != undefined)
            beforePost();
        if (fromCalendar == undefined)
            fromCalendar = false;
        var postActivity = [];
        var act;
        var mem = null;
        if (saveModifiedAssignment) {
        }
        else {
            //
            mem = this.newActivityMember;
        }
        if (dataArray) {
            postActivity = dataArray;
        }
        else {
            // validate data
            var err = '';
            var valDt = this.valStartEnd(mem.startDate, mem.endDate, mem);
            if (mem.perID == 0)
                err = "Personnel";
            err += (!valDt.withDates ? (err.length == 0 ? '' : ', ') + valDt.msg : '');
            if (err.length)
                err += (err.indexOf(", ") != -1) ? " are mandatory fields." : " is a mandatory field.";
            if (valDt.invalid)
                err += (err.length ? ', ' : '') + valDt.msg;
            if (valDt.outBound)
                err += (err.length ? ', ' : '') + ' Start and/or End dates selected, do not fall within the activity period.';
            if (err.length) {
                this.setTempMessage('msgTempPromptModal', err, 3500);
                return undefined;
            }
            var currActID_1 = fromCalendar ? mem.idAct : this.currActivity.id;
            //let act:dataActivity=this.datActivities.find((e)=>{return e.id==this.currActivity.id});
            act = this.datActivities.find(function (e) { return e.id == currActID_1; });
            postActivity = [{
                    id: mem.id,
                    actID: act.id,
                    perID: mem.perID,
                    teamID: mem.teamID,
                    startDate: __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](mem.startDate),
                    endDate: __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](mem.endDate),
                    mob: mem.mob,
                    demob: mem.demob,
                    night: mem.night,
                    isDay: mem.isDay,
                    coyID: mem.coyID,
                    posID: mem.posID,
                    groupID: mem.grpID,
                    actExtracted: act.extracted,
                    actStart: __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](act.startDate),
                    actEnd: __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](act.endDate),
                }];
            this.msgRequestPrompt = "Saving current activity member. Please wait...";
        }
        console.log('postActivity...start.');
        console.log(postActivity);
        console.log('postActivity...end.');
        this.subsSaveActivityMember = this.processData({
            method: "saveActivityMember",
            data: postActivity,
            isPosting: true,
            processMessage: "Saving activity member. Please wait...",
            subscription: this.subsSaveActivityMember,
            onSuccess: function (respArr) {
                respArr.forEach(function (data) {
                    if (data['obsolete']) {
                        _this.setTempMessage('msgTempPromptModal', "Obsolete activity record. Activity was modified by another user prior to updating changes or creation of activity member .", 8000);
                        return;
                    }
                    var extactDate = __WEBPACK_IMPORTED_MODULE_4__app_globals__["epochToDate"](data['lastUpdated']);
                    act = _this.datActivities.find(function (e) { return e.id == data['activityId']; });
                    if (act) {
                        act.extracted = extactDate;
                        if (_this.currActivity) {
                            if (_this.currActivity.id == act.id) {
                                _this.currActivity.extracted = extactDate;
                            }
                        }
                    }
                    if (!fromCalendar) {
                        // if action is not triggered from not from the calendar
                        // this line must be placed at the beginning for scatter data to work
                        // scatter data only works when record status is normal!!!
                        _this.udf.isActivityNormal(true);
                        // if new record was created, append the new record
                        var isNew = (data['action'] == 'AddActivityMember');
                        if (isNew) {
                            var act_1 = _this.currActivityObj;
                            var mem_1 = _this.newActivityMember;
                            mem_1.idAct = act_1.id;
                            mem_1.noPOB = act_1.noPOBCount;
                            mem_1.id = Number(data['resInt32']);
                            // mem.sort=this.activityMemberSortString(mem)
                            act_1.members.push(mem_1);
                            //G.sortBy(act.members,"sort");
                            //this.currActivity=this.cloneObject(act)
                            _this.copyActivity(act_1, _this.currActivity);
                            var team = _this.idxTeam['r' + mem_1.teamID];
                            if (team != undefined)
                                team.expView = true;
                            _this.sortActivityMembers(act_1.members);
                        }
                        else {
                            var act_2 = _this.currActivity;
                            var mem_2 = act_2.members;
                            var upd = _this.newActivityMember;
                            var obj = _this.currActivity.members.find(function (item) { return item['id'] == _this.currActivityMember.id; });
                            obj.perID = upd.perID;
                            obj.teamID = upd.teamID;
                            obj.startDate = upd.startDate;
                            obj.endDate = upd.endDate;
                            obj.mob = upd.mob;
                            obj.demob = upd.demob;
                            obj.night = upd.night;
                            obj.isDay = upd.isDay;
                            obj.coyID = upd.coyID;
                            obj.posID = upd.posID;
                            obj.grpID = upd.grpID;
                            obj = _this.currActivityObj.members.find(function (item) { return item['id'] == _this.currActivityMember.id; });
                            obj.perID = upd.perID;
                            obj.teamID = upd.teamID;
                            obj.startDate = upd.startDate;
                            obj.endDate = upd.endDate;
                            obj.mob = upd.mob;
                            obj.demob = upd.demob;
                            obj.night = upd.night;
                            obj.isDay = upd.isDay;
                            obj.coyID = upd.coyID;
                            obj.posID = upd.posID;
                            obj.grpID = upd.grpID;
                            _this.sortActivityMembers(act_2.members);
                        }
                        _this.newActivityMember = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["b" /* activityMember */]();
                        _this.getActivityConflicts(); // refresh conflict flags
                        _this.cancelActivityMember();
                        // if action is not triggered from not from the calendar(end)
                    }
                    else {
                        // if action is triggered from not from the calendar
                    }
                    if (onDataPost != undefined)
                        onDataPost(data);
                });
                // close popup window
                var btn = __WEBPACK_IMPORTED_MODULE_6_jquery__('#btnClose');
                if (btn.length != 0) {
                    btn.click();
                }
            } // end of onSuccess ...
        });
        return this.subsSaveActivityMember;
    };
    AppOplpanDataService.prototype.copyActivity = function (actFrom, actTo) {
        var _this = this;
        /*
    
            id:number;
        name:string;
        group:string;
        order:string;
        description:string;
        startDate:string;
        endDate:string;
        type:number;
        showInChart:boolean;
        noPOBCount:boolean;
        personnels:Array<number>;       // obsolete... not used
        members:Array<object>;
        conflicts?:object;
        parent?:any;
        */
        actTo.id = actFrom.id;
        actTo.name = actFrom.name;
        actTo.group = actFrom.group;
        actTo.order = actFrom.order;
        actTo.description = actFrom.description;
        actTo.startDate = actFrom.startDate;
        actTo.endDate = actFrom.endDate;
        actTo.type = actFrom.type;
        actTo.showInChart = actFrom.showInChart;
        actTo.noPOBCount = actFrom.noPOBCount;
        actTo.ready = actFrom.ready;
        actTo.members = [];
        actFrom.members.forEach(function (elem) { actTo.members.push(_this.cloneObject(elem)); });
        actTo.conflicts = this.cloneObject(actFrom.conflicts);
        actTo.parent = null;
    };
    AppOplpanDataService.prototype.deleteActivity = function (go) {
        var _this = this;
        if (this.currActivity.id == -1) {
            if (this.datActivities.length == 0) {
                this.setTempMessage("msgTempInfoPrompt", "No activty left to delete.");
            }
            else {
                this.setTempMessage("msgTempInfoPrompt", "Please select an activity to delete.");
            }
            return undefined;
        }
        if (go == undefined)
            go = false;
        if (!go) {
            this.udf.isActivityDeleteAlert(true);
        }
        else {
            this.udf.isActivityNormal(true);
            var actID_1 = this.currActivity.id;
            this.msgRequestPrompt = "Deleting current activity. Please wait...";
            return this.processData({
                method: "deleteActivity",
                data: {
                    id: actID_1,
                    name: this.currActivity.name,
                    actExtracted: this.currActivity.extracted,
                    actStart: __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](this.currActivity.startDate),
                    actEnd: __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](this.currActivity.endDate),
                },
                isPosting: true,
                processMessage: this.msgRequestPrompt,
                onSuccess: function (data) {
                    if (data['obsolete']) {
                        _this.setTempMessage(undefined, "You cannot delete an activity that was just updated by another user.", 3500);
                        _this.udf.isActivityNormal(true);
                        _this.udf.isActivityMemberNormal(true);
                        return;
                    }
                    // after successful deletion
                    _this.setTempMessage(undefined, "Seleced activity has been deleted.");
                    _this.udf.isActivityNormal(true);
                    _this.udf.isActivityMemberNormal(true);
                    var idx = _this.datActivities.findIndex(function (item) {
                        return item['id'] == actID_1;
                    });
                    _this.datActivities.splice(idx, 1);
                    _this.currActivity = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["c" /* dataActivity */]();
                    _this.currActivityObj = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["c" /* dataActivity */]();
                    _this.setOtherNavGroup();
                    // purge activity member client-side data
                }
            });
        }
    };
    AppOplpanDataService.prototype.deleteActivityMember = function (item, go) {
        var _this = this;
        if (item != undefined)
            this.currActivityMember = item;
        if (go == undefined)
            go = false;
        if (!go) {
            this.udf.isActivityMemberDeleteAlert(true);
        }
        else {
            var per_1 = this.idxPersonnel['r' + this.currActivityMember.perID];
            var memID_1 = this.currActivityMember.id;
            this.msgRequestPrompt = "Deleting activity member " + per_1['fullName'] + ". Please wait...";
            var act = this.datActivities.find(function (e) {
                return e.id == _this.currActivityMember.idAct;
            });
            this.subsSaveActivityMember = this.processData({
                method: "deleteActivityMember",
                data: {
                    id: memID_1,
                    teamID: this.currActivityMember.teamID,
                    perID: this.currActivityMember.perID,
                    actID: this.currActivityMember.idAct,
                    actExtracted: act.extracted,
                    actStart: __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](act.startDate),
                    actEnd: __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](act.endDate),
                },
                isPosting: true,
                processMessage: this.msgRequestPrompt,
                subscription: this.subsSaveActivityMember,
                onSuccess: function (data) {
                    if (data['obsolete']) {
                        _this.setTempMessage(undefined, "You cannot delete a member of this activity because another user has just updated this activity.", 3500);
                        _this.udf.isActivityNormal(true);
                        _this.udf.isActivityMemberNormal(true);
                        return;
                    }
                    // after successful deletion
                    _this.setTempMessage(undefined, "Activity member '" + per_1['fullName'] + "' has been removed.");
                    _this.udf.isActivityNormal(true);
                    _this.udf.isActivityMemberNormal(true);
                    var idx = _this.currActivity.members.findIndex(function (item) {
                        return item['id'] == memID_1;
                    });
                    _this.currActivity.members.splice(idx, 1);
                    _this.currActivityObj.members = _this.cloneObject(_this.currActivity.members);
                    // purge activity member client-side data
                }
            });
        }
        return this.subsSaveActivityMember;
    };
    AppOplpanDataService.prototype.editActivityMember = function (item) {
        var mem = this.newActivityMember;
        // scatter activity member
        this.currActivityMember = item; // record selected member for later update
        mem.id = item.id;
        mem.perID = item.perID;
        mem.startDate = item.startDate;
        mem.endDate = item.endDate;
        mem.mob = item.mob;
        mem.demob = item.demob;
        mem.night = item.night;
        mem.isDay = item.isDay;
        mem.teamID = item.teamID;
        mem.coyID = item.coyID;
        mem.posID = item.posID;
        mem.grpID = item.grpID;
        this.udf.isActivityMemberEdit(true);
    };
    AppOplpanDataService.prototype.cancelActivityMember = function () {
        var mem = this.newActivityMember;
        mem.id = -1;
        mem.perID = 0;
        mem.startDate = "";
        mem.endDate = "";
        mem.mob = true;
        mem.demob = true;
        mem.night = false;
        mem.isDay = false;
        mem.teamID = 0;
        mem.coyID = 0;
        mem.posID = 0;
        mem.grpID = 0;
        this.udf.isActivityMemberNormal(true);
    };
    AppOplpanDataService.prototype.cancelActivity = function () {
        this.udf.isActivityNormal(true);
        this.msgTempPrompt = "";
        this.scatterActivity();
    };
    AppOplpanDataService.prototype.saveActivity = function (onDataPost, beforePost, onError) {
        var _this = this;
        var postActivity = {};
        var act = this.currActivity;
        // validate data
        var err = '';
        var valDt = this.valStartEnd(act.startDate, act.endDate);
        if (act.name.length == 0)
            err += 'Activity name';
        err += (!valDt.withDates ? (err.length == 0 ? '' : ', ') + valDt.msg : '');
        if (err.length)
            err += (err.indexOf(", ") != -1) ? " are mandatory fields." : " is a mandatory field.";
        if (valDt.invalid)
            err += (err.length ? ', ' : '') + valDt.msg;
        if (err.length) {
            this.setTempMessage(undefined, err, 5000);
            return false;
        }
        var spVal = this.validateDateSpan();
        if (!spVal['isValid']) {
            this.setTempMessage(undefined, "Date specified cannot accommodate duration of personnel assignments (i.e. earliest start date to latest end date).", 6000);
            return false;
        }
        act.group = act.startDate.substr(act.startDate.length - 4, 4);
        postActivity = {
            id: act.id,
            name: act.name,
            description: act.description,
            startDate: __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](act.startDate),
            endDate: __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](act.endDate),
            type: act.type,
            showInChart: act.showInChart,
            noPOBCount: act.noPOBCount,
            ready: act.ready,
            site: act.site,
            vessel: act.vessel,
            upmanning: act.upmanning,
            extracted: act.extracted,
            shiftDates: spVal['willShift'],
            shiftBy: spVal['shiftBy'],
            subscription: this.subsSaveActivity
        };
        this.msgRequestPrompt = "Saving current activity. Please wait.";
        this.subsSaveActivity = this.processData({
            method: "saveActivity",
            data: postActivity,
            isPosting: true,
            processMessage: this.msgRequestPrompt,
            onSuccess: function (data) {
                // this line must be placed at the beginning for scatter data to work
                // scatter data only works when record status is normal!!!
                var shftDays = data["shiftDays"];
                if (data['obsolete']) {
                    _this.setTempMessage(undefined, "Obsolete activity record. Activity was modified by another user prior to saving changes. Please reload [F5] to get the latest version of the current activity.", 8000);
                    return;
                }
                if (data['invalidScope']) {
                    _this.setTempMessage(undefined, "Invalid activity date scope. Changing start date will shift member assignment dates by " + shftDays + " " + (shftDays == 1 ? "day" : "days") + ", resulting to assingment dates outside of the new activity period.", 8000);
                    return;
                }
                else if (data['blockMembers']) {
                    // make member start and end date the same as activity start and end date
                    _this.currActivityObj.members.forEach(function (mem) {
                        mem.startDate = act.startDate;
                        mem.endDate = act.endDate;
                    });
                }
                else if (shftDays != 0) {
                    /*
                    let actSource = this.datActivities.find((act)=>{return act.id == this.currActivity.id});
                    if(actSource){
                      console.log(actSource);
                      actSource.members.forEach((mem:activityMember)=>{
                        let sd:Date = G.longToDate(mem.startDate)
                        let ed:Date = G.longToDate(mem.endDate)
                        sd=G.addDays(sd,shftDays);
                        ed=G.addDays(ed,shftDays);
        
                        mem.startDate= G.dateToString(sd,true);
                        mem.endDate= G.dateToString(ed,true);
                      });
                    }
        
        
                    this.currActivity.members.forEach((mem:activityMember)=>{
                      let sd:Date = G.longToDate(mem.startDate)
                      let ed:Date = G.longToDate(mem.endDate)
                      sd=G.addDays(sd,shftDays);
                      ed=G.addDays(ed,shftDays);
        
                      mem.startDate= G.dateToString(sd,true);
                      mem.endDate= G.dateToString(ed,true);
                    });*/
                    _this.currActivityObj.members.forEach(function (mem) {
                        var sd = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](mem.startDate);
                        var ed = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](mem.endDate);
                        sd = __WEBPACK_IMPORTED_MODULE_4__app_globals__["addDays"](sd, shftDays);
                        ed = __WEBPACK_IMPORTED_MODULE_4__app_globals__["addDays"](ed, shftDays);
                        mem.startDate = __WEBPACK_IMPORTED_MODULE_4__app_globals__["dateToString"](sd, true);
                        mem.endDate = __WEBPACK_IMPORTED_MODULE_4__app_globals__["dateToString"](ed, true);
                    });
                    //console.log(this.currActivity.members);
                    //console.log(this.currActivityObj.members);
                }
                _this.udf.isActivityNormal(true);
                // if new record was created, append the new record
                var isNew = (data['action'] == 'AddActivity');
                _this.setTempMessage("msgTempInfoPrompt", (isNew ? "New activity" : "Activity " + act.name) + " was successfully " + (isNew ? " created." : " saved."));
                _this.currActivityObj.extracted = __WEBPACK_IMPORTED_MODULE_4__app_globals__["epochToDate"](data.lastUpdated);
                if (isNew) {
                    _this.currActivityObj = _this.currActivity;
                    var newId = Number(data['resInt32']);
                    _this.currActivityObj.id = newId; // new id is stored in resInt32 field
                    _this.datActivities.push(_this.currActivityObj);
                    _this.currActivity = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["c" /* dataActivity */]();
                    _this.scatterData(_this.currActivityObj);
                }
                else {
                    var noUpdate = data['obsolete'] || data['removed'];
                    /*          if(noUpdate){
        
                                let origData:dataActivity=this.datActivities.find((act)=>{return act.id == this.currActivity.id});
        
                                origData.obsolete = data['obsolete'];
                                origData.removed = data['removed'];
        
                                // restore previous state and lock the record to pervent further user action...
                                this.scatterActivity(origData);
        
        
                              }else{*/
                    _this.currActivityObj.name = _this.currActivity.name;
                    _this.currActivityObj.description = _this.currActivity.description;
                    _this.currActivityObj.startDate = _this.currActivity.startDate;
                    _this.currActivityObj.endDate = _this.currActivity.endDate;
                    _this.currActivityObj.noPOBCount = _this.currActivity.noPOBCount;
                    _this.currActivityObj.ready = _this.currActivity.ready;
                    _this.currActivity.obsolete = true;
                    _this.currActivity.removed = true;
                    _this.currActivityObj.obsolete = true;
                    _this.currActivityObj.removed = true;
                    //}
                    var days = spVal['shiftBy'];
                }
                //this.getActivityConflicts(); // refresh conflict flags
                _this.setOtherNavGroup();
                // refresh display
                _this.scatterData(_this.currActivityObj);
            }
        });
        return true;
    };
    AppOplpanDataService.prototype.validateDateSpan = function () {
        var ret = { isValid: true, willShift: false, shiftBy: 0 };
        var act = this.currActivity;
        var minStart = null;
        var maxEnd = null;
        var actStart = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](this.currActivity.startDate);
        var actEnd = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](this.currActivity.endDate);
        var actStartOrig = this.currActivityObj.startDate;
        var actEndOrig = this.currActivityObj.endDate;
        var isBlockMembers = false;
        if (act.members.length > 0) {
            minStart = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](act.members[0]['startDate']);
            maxEnd = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](act.members[0]['endDate']);
            if (act.members.length > 1) {
                var ctr_1 = 0;
                act.members.forEach(function (mem) {
                    var memStart = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](mem.startDate);
                    var memEnd = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](mem.endDate);
                    if (ctr_1 == 0) {
                        isBlockMembers = (mem.startDate == actStartOrig && mem.endDate == actEndOrig);
                    }
                    else {
                        if (isBlockMembers) {
                            isBlockMembers = (mem.startDate == actStartOrig && mem.endDate == actEndOrig);
                        }
                    }
                    if (memStart < minStart)
                        minStart = memStart;
                    if (memEnd > maxEnd)
                        maxEnd = memEnd;
                    ctr_1++;
                });
            }
        }
        //alert(actStartOrig + ' , ' + actEndOrig +  ', ' + isBlockMembers);
        // bypass minStart and minEnd test to allow blockMembers 2018/07/27
        if (minStart == null || maxEnd == null || isBlockMembers) {
            ret['isValid'] = true;
        }
        else {
            var memDuration = __WEBPACK_IMPORTED_MODULE_4__app_globals__["daysDiff"](minStart, maxEnd);
            var offStart = __WEBPACK_IMPORTED_MODULE_4__app_globals__["daysDiff"](actStart, minStart); // to be used in later code revisions
            var actDuration = __WEBPACK_IMPORTED_MODULE_4__app_globals__["daysDiff"](actStart, actEnd);
            ret['isValid'] = (actDuration >= memDuration);
            ret['willShift'] = ret['isValid'] &&
                (actStart < minStart || maxEnd < minStart || actStart > maxEnd || actEnd > maxEnd);
            ret['shiftBy'] = (ret['willShift'] ? __WEBPACK_IMPORTED_MODULE_4__app_globals__["daysDiff"](minStart, actStart) : 0) * (minStart > actStart ? -1 : 1);
        }
        // must return validity, if will shift
        return ret;
    };
    AppOplpanDataService.prototype.validateDateSpanX20180730 = function () {
        var ret = { isValid: true, willShift: false, shiftBy: 0 };
        var act = this.currActivity;
        var minStart = null;
        var maxEnd = null;
        var actStart = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](this.currActivity.startDate);
        var actEnd = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](this.currActivity.endDate);
        if (act.members.length > 0) {
            minStart = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](act.members[0]['startDate']);
            maxEnd = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](act.members[0]['endDate']);
            if (act.members.length > 1) {
                act.members.forEach(function (mem) {
                    var memStart = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](mem.startDate);
                    var memEnd = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](mem.endDate);
                    if (memStart < minStart)
                        minStart = memStart;
                    if (memEnd > maxEnd)
                        maxEnd = memEnd;
                });
            }
        }
        if (minStart == null || maxEnd == null) {
            ret['isValid'] = true;
        }
        else {
            var memDuration = __WEBPACK_IMPORTED_MODULE_4__app_globals__["daysDiff"](minStart, maxEnd);
            var offStart = __WEBPACK_IMPORTED_MODULE_4__app_globals__["daysDiff"](actStart, minStart); // to be used in later code revisions
            var actDuration = __WEBPACK_IMPORTED_MODULE_4__app_globals__["daysDiff"](actStart, actEnd);
            ret['isValid'] = (actDuration >= memDuration);
            ret['willShift'] = ret['isValid'] &&
                (actStart < minStart || maxEnd < minStart || actStart > maxEnd || actEnd > maxEnd);
            ret['shiftBy'] = (ret['willShift'] ? __WEBPACK_IMPORTED_MODULE_4__app_globals__["daysDiff"](minStart, actStart) : 0) * (minStart > actStart ? -1 : 1);
        }
        // must return validity, if will shift
        return ret;
    };
    AppOplpanDataService.prototype.savePersonnel = function (onDataPost, beforePost, onError) {
        var _this = this;
        if (beforePost != undefined)
            beforePost();
        var result = null;
        var fullName = this.currPersonnel.fullName;
        if (fullName) {
            // force first character of the fullname to be in capital letter
            this.currPersonnel.fullName = fullName[0].toUpperCase() + fullName.substring(1);
        }
        this.setPosting("Saving personnel. Please wait...");
        // record edited certs into this.currPersonnel
        if (this.currPersonnel.certificates) {
            this.datCertificates.forEach(function (c) {
                var certs = _this.currPersonnel.certificates;
                var pc = certs.find(function (cert) {
                    return cert.id == c.id;
                });
                if (pc) {
                    pc.req = c.required;
                    pc.issue = c.issued ? __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](c.issued) : "";
                    pc.expiry = c.expiry ? __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](c.expiry) : "";
                }
                else if (c.id != 0 && (c.issued != null || c.expiry != null || c.required)) {
                    _this.currPersonnel.certificates.push({ id: c.id, req: c.required,
                        issue: c.issued ? __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](c.issued) : "",
                        expiry: c.expiry ? __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](c.expiry) : "" });
                }
            });
        }
        // convert long date to short date because this is the acceptable format in the
        // server-side method.
        this.currPersonnel.birthDateShort = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](this.currPersonnel.birthDate);
        this.subsSavePersonnel = this.http.post(__WEBPACK_IMPORTED_MODULE_4__app_globals__["getUrlRequest"]("oplanData", "savePersonnel"), JSON.stringify(this.currPersonnel))
            .subscribe(function (data) {
            if (onDataPost != undefined) {
                onDataPost(data);
            }
            _this.msgPersonUpdateError = '';
            if (data == null) {
                _this.msgPersonUpdateError = "Unknown error occured!!!";
            }
            else {
                if (data['stat'] == 'fail') {
                    _this.msgPersonUpdateError = "An error has occured when saving data ...\n" + data['message'];
                }
                else {
                    // this will update the current display with the new values after savig data
                    _this.showUpdatedPersonnel(data);
                }
            }
            _this.noPosting();
        }, function (err) {
            if (onError != undefined && onError != null) {
                onError(err);
            }
            _this.withError(err);
            return undefined;
        });
        return this.subsSavePersonnel;
    };
    AppOplpanDataService.prototype.resetMessages = function () {
        this.msgPersonUpdateError = '';
    };
    AppOplpanDataService.prototype.setPersonnelNavGroup = function (sort) {
        var _this = this;
        //if(sort!= undefined) this.navSortBy=sort;
        if (sort != undefined)
            this.udf.currentModule = sort;
        //let inputArray:Array<dataPersonnel>=this.items;
        var inputArray = this.datPersonnel;
        if (this.udf.searchText.length) {
            inputArray = inputArray.filter(function (item) {
                return item['fullName'].toLowerCase() != "-" && item['fullName'].toLowerCase().
                    indexOf(_this.udf.searchText.toLocaleLowerCase()) != -1;
            });
        }
        else {
            inputArray = inputArray.filter(function (item) { return item['fullName'].toLowerCase() != ""; });
        }
        switch (this.udf.currentModule) {
            case 'By Name':
                this.itemsGroupings = __WEBPACK_IMPORTED_MODULE_4__app_globals__["groupBy"](inputArray, 'nameGroup');
                break;
            case 'By Company':
                this.itemsGroupings = __WEBPACK_IMPORTED_MODULE_4__app_globals__["groupBy"](inputArray, 'companyID', this.datCompany);
                break;
            case 'By Position':
                this.itemsGroupings = __WEBPACK_IMPORTED_MODULE_4__app_globals__["groupBy"](inputArray, 'positionID', this.datPosition);
                break;
            /*
          case 'Activities':
          case 'Teams':
            alert(this.udf.currentModule)
            break;
  */
            default:
                this.itemsGroupings = [];
        }
        if (this.itemsGroupings.length) {
            if (this.itemsGroupings[0]['name'] == '' && this.itemsGroupings.length > 1) {
                this.udf.currGroup = this.itemsGroupings[1]['name'];
            }
            else {
                this.udf.currGroup = this.itemsGroupings[0]['name'];
            }
        }
    };
    AppOplpanDataService.prototype.refreshNavGroup = function () {
        if (this.udf.currentModule == 'Teams' || this.udf.currentModule == 'Activities') {
            this.setOtherNavGroup();
        }
        else {
            this.setPersonnelNavGroup();
        }
    };
    AppOplpanDataService.prototype.setOtherNavGroup = function () {
        var _this = this;
        setTimeout(function () {
            if (_this.udf.currentModule == 'Teams') {
                _this.arrangeTeamData();
            }
            else if (_this.udf.currentModule == 'Activities') {
                _this.arrangeActivityData();
            }
        }, 300);
    };
    AppOplpanDataService.prototype.filterFieldName = function () {
        return this.udf.currentModule == "Teams" ||
            this.udf.currentModule == "Activities" ?
            "name" : "fullName";
    };
    AppOplpanDataService.prototype.getPersonnelActiveAcivities = function () {
        var _this = this;
        if (!this.getActiveActivities)
            return undefined;
        return this.processData({
            method: "getPersonActivityActiveList",
            data: null,
            isPosting: false,
            processMessage: "Loading personnel active activities. Please wait...",
            onSuccess: function (data) {
                // reset all activity count badges in personnels badges
                _this.datPersonnel.forEach(function (per) {
                    if (per.activityCount) {
                        per.activityCount = 0;
                        per.activities = [];
                        per.maxDays = 0;
                    }
                });
                // process activity name lookup
                _this.lkpActivityName = {}; // reset activity name lookup
                data['activities'].forEach(function (act) { return _this.lkpActivityName["a_" + act[0]] = act[1]; });
                data["persons"].forEach(function (per) {
                    var perObj = _this.idxPersonnel["r" + per[0]];
                    if (perObj != undefined) {
                        perObj.activityCount = per[1];
                        perObj.maxDays = per[2];
                        per[3].forEach(function (act) {
                            perObj.activities.push({ "aid": act[0], "start": act[1], "end": act[2], "mob": act[3], "demob": act[4], "night": act[5], "isDay": act[6] });
                        });
                    }
                });
                _this.getActiveActivities = false;
            }
        });
    };
    AppOplpanDataService.prototype.getPersonnel = function (force, onDataLoad, beforeLoad, onError) {
        //this.clearSubscription(this.subsPersonnel);
        var _this = this;
        var person;
        if (force == undefined)
            force = false;
        if (force || this.datPersonnel.length == 0) {
            if (beforeLoad != undefined && beforeLoad != null)
                beforeLoad();
            this.setLoading("Loading personnels. Please wait...");
            return this.http
                .get(__WEBPACK_IMPORTED_MODULE_4__app_globals__["getUrlRequest"]("oplanData", "getPersonnel"))
                .subscribe(function (data) {
                // process personnel data
                _this.datPersonnel = [];
                data['personnel'].forEach(function (per) {
                    var acts = new Array();
                    //let certs:Array<object>=new Array<object>();
                    person = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["d" /* dataPersonnel */]();
                    person.id = per[0];
                    person.fullName = per[1];
                    person.nameGroup = person.fullName.substr(0, 1).toUpperCase();
                    person.email = per[2];
                    person.companyID = per[3];
                    person.positionID = per[4];
                    person.activityCount = per[5];
                    person.activities = acts;
                    person.gender = per[7];
                    person.certificates = null;
                    //person.certificates=certs;
                    person.birthDate = __WEBPACK_IMPORTED_MODULE_4__app_globals__["shortToLongDate"](per[9]);
                    person.xl = per[10];
                    console.log('per[10]', per[10]);
                    // process active and comming activities the person is member of
                    per[6].forEach(function (act) {
                        acts[acts.length] = {
                            aid: act[0],
                            start: act[1],
                            end: act[2],
                            mob: act[3],
                            demob: act[4],
                            night: act[5],
                            isDay: act[6],
                        };
                    });
                    // process active and comming activities the person is member of
                    /*per[8].forEach(cert=>{
                      certs[certs.length]={
                        id:cert[0],
                        req:(cert[1]=="1"),
                        issue:G.shortToLongDate(cert[2]),
                        expiry:G.shortToLongDate(cert[3]),
                      }
                    })*/
                    _this.datPersonnel.push(person);
                    _this.idxPersonnel['r' + person.id] = person;
                });
                // process active and future activity name lookup
                //data['actLookup'].forEach(act=>this.lkpActivityName["a_"+act[0]]=act[1]);
                //this.getPersonnelActiveAcivities();
                // sort personnel array by fullname
                __WEBPACK_IMPORTED_MODULE_4__app_globals__["sortBy"](_this.datPersonnel, "fullName");
                if (onDataLoad != undefined && onDataLoad != null) {
                    onDataLoad(_this.datPersonnel);
                }
                _this.noLoading();
            }, function (err) {
                if (onError != undefined && onError != null) {
                    onError(err);
                }
                _this.withError(err);
            });
        }
        else {
            if (onDataLoad != undefined)
                onDataLoad(this.datPersonnel);
        }
    };
    AppOplpanDataService.prototype.getCompanies = function (force, onDataLoad, beforeLoad, onError) {
        var _this = this;
        if (force == undefined)
            force = false;
        if (force || this.datCompany.length == 0) {
            if (beforeLoad != undefined && beforeLoad != null)
                beforeLoad();
            this.setLoading("Loading companies. Please wait...");
            return this.http
                .get(__WEBPACK_IMPORTED_MODULE_4__app_globals__["getUrlRequest"]("oplanData", "getCompanies"))
                .subscribe(function (data) {
                data['rows'].forEach(function (element) {
                    var coy = {
                        id: element[0],
                        name: element[1]
                    };
                    _this.datCompany.push(coy);
                    _this.idxCompany['r' + coy['id']] = coy;
                });
                if (onDataLoad != undefined && onDataLoad != null) {
                    onDataLoad(_this.datCompany);
                }
                _this.noLoading();
            }, function (err) {
                if (onError != undefined && onError != null) {
                    onError(err);
                }
                _this.withError(err);
            });
        }
    };
    AppOplpanDataService.prototype.getPositions = function (force, onDataLoad, beforeLoad, onError) {
        var _this = this;
        if (force == undefined)
            force = false;
        if (force || this.datPosition.length == 0) {
            if (beforeLoad != undefined && beforeLoad != null)
                beforeLoad();
            this.setLoading("Loading positions. Please wait...");
            return this.http
                .get(__WEBPACK_IMPORTED_MODULE_4__app_globals__["getUrlRequest"]("oplanData", "getPositions"))
                .subscribe(function (data) {
                data['rows'].forEach(function (element) {
                    var pos = {
                        id: element[0],
                        name: element[1]
                    };
                    _this.datPosition.push(pos);
                    _this.idxPosition['r' + pos['id']] = pos;
                });
                if (onDataLoad != undefined && onDataLoad != null) {
                    onDataLoad(_this.datPosition);
                }
                _this.noLoading();
            }, function (err) {
                if (onError != undefined && onError != null) {
                    onError(err);
                }
                _this.withError(err);
            });
        }
    };
    AppOplpanDataService.prototype.getTeams = function (force, onDataLoad, beforeLoad, onError) {
        //this.clearSubscription(this.subsTeams);
        var _this = this;
        var team;
        if (force == undefined)
            force = false;
        if (force || this.datActivities.length == 0) {
            if (beforeLoad != undefined && beforeLoad != null)
                beforeLoad();
            this.recordExpandedTeams();
            this.setLoading("Loading teams. Please wait...");
            return this.http
                .get(__WEBPACK_IMPORTED_MODULE_4__app_globals__["getUrlRequest"]("oplanData", "getTeams"))
                .subscribe(function (data) {
                _this.datTeams = [];
                data['teams'].forEach(function (element) {
                    /*
                    export class dataTeam{
    
                          constructor(){
                              this.id=-1;
                          }
    
                          id:number;
                          name:string;
                          description:string;
                          core:boolean;
                          order:number;
                          upman:boolean;
                          members:Array<teamMember>;
                      }
    
                      export class teamMember{
                          perID:number;
                          shift:string;
                          subTeamID:number;
                      }
    
                  }*/
                    // [1,"Operations Core Crew","Team A Description",1,1,0,[[33,"",0],[33,"",1]]
                    team = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["e" /* dataTeam */]();
                    team.id = Number(element[0]);
                    team.name = element[1];
                    team.description = element[2];
                    team.core = Boolean(element[3] == 1);
                    team.group = team.core ? "Core" : "Non-Core";
                    team.order = Number(element[4]);
                    team.upman = (element[5] == 1);
                    team.members = [];
                    //recid, tid, pid, sd, ed, mob, demob, night
                    //"2","1","179","20170328","20170411","1","1","0"
                    element[6].forEach(function (mem) {
                        var tmem = {
                            id: mem[0],
                            teamID: team.id,
                            perID: mem[1],
                            shift: mem[2],
                            subTeamID: mem[3],
                            sort: ""
                        };
                        team.members.push(tmem);
                    });
                    _this.datTeams.push(team);
                    _this.idxTeam['r' + team.id] = team;
                });
                //data[',[]]],"subTeams":[[0,""],[3,"Derusters/Painters/Blasters"],']
                _this.lkpSubTeams = {};
                data['subTeams'].forEach(function (sbt) {
                    _this.lkpSubTeams['r' + sbt[0]] = sbt[1];
                });
                if (onDataLoad != undefined && onDataLoad != null) {
                    onDataLoad(_this.datPosition);
                }
                _this.setExpandedTeams();
                _this.noLoading();
            }, function (err) {
                if (onError != undefined && onError != null) {
                    onError(err);
                }
                _this.withError(err);
            });
        }
        else {
            if (onDataLoad != undefined)
                onDataLoad(this.datTeams);
        }
        return undefined;
    };
    AppOplpanDataService.prototype.getActivities = function (force, onDataLoad, beforeLoad, onError) {
        //this.clearSubscription(this.subsActivities);
        var _this = this;
        var activity;
        if (force == undefined)
            force = false;
        if (force || this.datActivities.length == 0) {
            if (beforeLoad != undefined && beforeLoad != null)
                beforeLoad();
            this.setLoading("Loading activities. Please wait...");
            return this.http
                .get(__WEBPACK_IMPORTED_MODULE_4__app_globals__["getUrlRequest"]("oplanData", "getActivities"))
                .subscribe(function (data) {
                var extactDate = __WEBPACK_IMPORTED_MODULE_4__app_globals__["epochToDate"](data['extracted']);
                _this.datActivities = [];
                _this.idxActivity = {};
                data['rows'].forEach(function (element) {
                    /*
                    export class dataActivity{
                      id:number;
                      name:string;
                      description:string;
                      startDate:string;
                      endDate:string;
                      personnels:Array<number>;
                      members:Array<object>;
      
                      id:number;
                      teamID:number;
                      perID:number;
                      startDate:string;
                      endDate:string;
                      mob:boolean;
                      demob:boolean;
                      night:boolean;
      
                  }*/
                    activity = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["c" /* dataActivity */]();
                    activity.id = Number(element[0]);
                    activity.name = element[1];
                    activity.description = element[2];
                    activity.group = (element[3].length ? element[3].substr(0, 4) : "");
                    activity.order = element[3];
                    activity.startDate = __WEBPACK_IMPORTED_MODULE_4__app_globals__["shortToLongDate"](element[3]);
                    activity.endDate = __WEBPACK_IMPORTED_MODULE_4__app_globals__["shortToLongDate"](element[4]);
                    activity.type = element[5];
                    activity.showInChart = (element[6] == 1);
                    activity.noPOBCount = (element[7] == 1);
                    activity.ready = element[8];
                    activity.site = element[10];
                    activity.vessel = (element[11] == 1);
                    activity.upmanning = (element[12] == 1);
                    activity.extracted = extactDate;
                    activity.members = [];
                    activity.personnels = [];
                    // if processed activity is the same as the current activity, set extract date of the
                    // current activity the same as the extract date of the activity being processed
                    if (_this.currActivity) {
                        if (_this.currActivity.id == activity.id) {
                            // probably a good idea to also update details of the currentActivity
                            // with the details of the processed activity ????
                            _this.currActivity.extracted = extactDate;
                            if (_this.currActivityObj.id == activity.id)
                                _this.currActivityObj = activity;
                        }
                    }
                    //recid, tid, pid, sd, ed, mob, demob, night
                    //"2","1","179","20170328","20170411","1","1","0"
                    element[9].forEach(function (mem) {
                        var amem = {
                            sort: "",
                            id: Number(mem[0]),
                            idAct: Number(element[0]),
                            teamID: Number(mem[1]),
                            perID: Number(mem[2]),
                            startDate: __WEBPACK_IMPORTED_MODULE_4__app_globals__["shortToLongDate"](mem[3]),
                            endDate: __WEBPACK_IMPORTED_MODULE_4__app_globals__["shortToLongDate"](mem[4]),
                            mob: mem[5] == 1,
                            demob: mem[6] == 1,
                            night: mem[7] == 1,
                            isDay: mem[8] == 1,
                            noPOB: (element[7] == 1),
                            coyID: mem[9],
                            posID: mem[10],
                            grpID: mem[11]
                        };
                        amem.sort = _this.activityMemberSortString(amem);
                        activity.members.push(amem);
                    });
                    _this.datActivities.push(activity);
                    _this.sortActivityMembers(activity.members);
                    _this.idxActivity['r' + activity.id] = activity;
                    // process activity types
                    if (data['types'] != undefined) {
                        _this.datPriorities = [];
                        data['types'].forEach(function (type) {
                            var tp = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["g" /* lookupTemplate */]();
                            tp.id = Number(type[0]);
                            tp.name = type[1];
                            tp.order = Number(type[2]);
                            //
                            _this.datPriorities.push(tp);
                        });
                    }
                });
                if (onDataLoad != undefined && onDataLoad != null) {
                    onDataLoad(_this.datPosition);
                }
                _this.noLoading();
            }, function (err) {
                if (onError != undefined && onError != null) {
                    onError(err);
                }
                _this.withError(err);
            });
        }
        else {
            if (onDataLoad != undefined)
                onDataLoad(this.datActivities);
        }
        return undefined;
    };
    AppOplpanDataService.prototype.setLoading = function (msg) {
        this.setWindowDimensions();
        this.isLoading = true;
        if (msg != undefined)
            this.isLoadingMessage = msg;
    };
    AppOplpanDataService.prototype.noLoading = function () {
        this.isLoading = false;
        this.isLoadingMessage = "Loading. Please wait...";
    };
    AppOplpanDataService.prototype.setPosting = function (msg) {
        this.isPosting = true;
        if (msg != undefined)
            this.isPostingMessage = msg;
    };
    AppOplpanDataService.prototype.noPosting = function () {
        this.isPosting = false;
        this.isPostingMessage = "Loading. Please wait...";
    };
    AppOplpanDataService.prototype.withError = function (err) {
        this.noLoading();
        this.noPosting();
    };
    AppOplpanDataService.prototype.setWindowDimensions = function () {
        var win = __WEBPACK_IMPORTED_MODULE_6_jquery__(window);
        this.winHeight = win.height();
        this.winWidth = win.width();
        var div = __WEBPACK_IMPORTED_MODULE_6_jquery__('#spinMessage');
        this.getMessageLeft = (this.winWidth - div.width()) / 2;
        this.getMessageTop = (this.winHeight - div.height()) / 2;
    };
    AppOplpanDataService.prototype.dataActionClicked = function (e) {
        switch (__WEBPACK_IMPORTED_MODULE_6_jquery__(e.srcElement).attr('id')) {
            case 'per_add':
                this.currPersonnel = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["d" /* dataPersonnel */]();
                this.udf.modePersonnel = 'add';
                this.resetCerts();
                // clear certificates
                break;
            case 'per_edit':
                if (this.currPersonnel.id == -1) {
                    this.setTempMessage("msgPersonUpdateInstruction", "Information : Please select a personnel to edit.");
                    return;
                }
                this.udf.modePersonnel = 'edit';
                break;
            case 'per_delete':
                if (this.currPersonnel.id == -1) {
                    alert("Please select a personnel.");
                    return;
                }
                if (!confirm("You are about to delete the currently selected personnel.\n" +
                    "This action will also remove all the attached information.\n\nDo you want to proceed?\n\n" +
                    "Click [Ok] to proceed or [Cancel] to abort."))
                    return;
                this.deletePersonnel();
                break;
        }
    };
    AppOplpanDataService.prototype.arrangeActivityData = function () {
        var _this = this;
        var reGroup = true;
        this.clearSubscription(this.subsPersonnel);
        this.subsPersonnel = this.getPersonnel(true, function (data) {
            _this.clearSubscription(_this.subsTeams);
            _this.subsTeams = _this.getTeams(true, function (data) {
                _this.clearSubscription(_this.subsActivities);
                _this.subsActivities = _this.getActivities(true, function (data) {
                    reGroup = false;
                    _this.setOtherItemGroupings();
                });
            });
        });
        if (reGroup)
            this.setOtherItemGroupings();
    };
    AppOplpanDataService.prototype.arrangeTeamData = function () {
        var _this = this;
        var reGroup = true;
        this.clearSubscription(this.subsPersonnel);
        this.subsPersonnel = this.getPersonnel(true, function (data) {
            _this.clearSubscription(_this.subsTeams);
            _this.subsTeams = _this.getTeams(true, function (data) {
                reGroup = false;
                _this.setOtherItemGroupings();
            }); // teams
        }); // personnel
        if (reGroup)
            this.setOtherItemGroupings();
    };
    AppOplpanDataService.prototype.isExisting = function (objArr, fieldName, fieldValue) {
        var ret = false;
        var idx = objArr.findIndex(function (x) { return x[fieldName] == fieldValue; });
        return ret;
    };
    AppOplpanDataService.prototype.getItemIndex = function (objArr, fieldName, fieldValue) {
        return objArr.findIndex(function (x) { return x[fieldName] == fieldValue; });
    };
    AppOplpanDataService.prototype.getActivityGroupings = function (findAct) {
        var _this = this;
        var ret = [];
        var inputArray = this.datActivities;
        if (this.udf.searchText.length) {
            inputArray = inputArray.filter(function (item) {
                return item['name'].toLowerCase() != "-" && item['name'].toLowerCase().
                    indexOf(_this.udf.searchText.toLocaleLowerCase()) != -1;
            });
        }
        //this.datActivities.forEach((act:dataActivity)=>act.parent=null);
        inputArray.forEach(function (act) {
            act.parent = null;
            var sd = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](act.startDate);
            var ed = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](act.endDate);
            var sy = sd.getFullYear();
            var sm = sd.getMonth();
            var ey = ed.getFullYear();
            var em = ed.getMonth();
            var cy = sy;
            var cm = sm;
            var moSpan = 0;
            var tmpDt;
            while (true) {
                moSpan++;
                var idxYear = _this.getItemIndex(ret, 'year', cy);
                var objYear = void 0;
                if (idxYear == -1) {
                    // year object is not yet existing
                    objYear = { "year": cy, exp: false, "months": [] };
                    ret.push(objYear);
                }
                else {
                    objYear = ret[idxYear];
                }
                var idxMonth = _this.getItemIndex(objYear['months'], 'month', cm);
                var objMonth = void 0;
                if (idxMonth == -1) {
                    // month is not yet existing
                    objMonth = { "month": cm, exp: false, "priorities": [], parent: objYear };
                    objYear['months'].push(objMonth);
                }
                else {
                    objMonth = objYear['months'][idxMonth];
                }
                var idxPriority = _this.getItemIndex(objMonth['priorities'], 'priority', act.type);
                var objPriority = void 0;
                if (idxPriority == -1) {
                    // month is not yet existing
                    var typeObj = undefined;
                    var order = 999;
                    var priorityName = "<UnAssigned>";
                    if (act.type != 0)
                        typeObj = _this.datPriorities[_this.getItemIndex(_this.datPriorities, "id", act.type)];
                    if (typeObj != undefined) {
                        order = typeObj['order'];
                        priorityName = typeObj['name'];
                    }
                    objPriority = { "priority": act.type,
                        "order": order,
                        "name": priorityName,
                        "activities": [],
                        exp: false,
                        parent: objMonth };
                    objMonth['priorities'].push(objPriority);
                }
                else {
                    objPriority = objMonth['priorities'][idxPriority];
                }
                //if(act.parent==null)
                act.parent = objPriority;
                objPriority['activities'].push(act);
                if (cy == ey && cm == em)
                    break;
                cm++;
                tmpDt = new Date(cy, cm, 1);
                cy = tmpDt.getFullYear();
                cm = tmpDt.getMonth();
            }
        });
        if (ret.length) {
            // sort data...
            __WEBPACK_IMPORTED_MODULE_4__app_globals__["sortBy"](ret, "year", true);
            ret.forEach(function (yr) {
                __WEBPACK_IMPORTED_MODULE_4__app_globals__["sortBy"](yr['months'], "month", true);
                yr['months'].forEach(function (mo) {
                    __WEBPACK_IMPORTED_MODULE_4__app_globals__["sortBy"](mo['priorities'], "order");
                    mo['priorities'].forEach(function (pri) {
                        __WEBPACK_IMPORTED_MODULE_4__app_globals__["sortBy"](mo['activities'], "name");
                    });
                });
            });
        }
        /*
            if(ret.length){
              let yr=ret[1];
              yr['exp']=true;
              if(yr['months'].length){
                let mo=yr['months'][4];
                mo['exp']=true;
                if(mo['priorities'].length){
                  mo['priorities'][1]['exp']=true;
                }
              }
            }
        
        */
        if (this.udf.searchText.length || this.onActivityLoadAction) {
            //let ret:Array<object>=ret;
            if (this.udf.searchText.length) {
                // search text is not empty, expand all parent nodes
                for (var yIdx = 0; yIdx < ret.length; yIdx++) {
                    ret[yIdx]['exp'] = true;
                    var mos = ret[yIdx]['months'];
                    for (var mIdx = 0; mIdx < mos.length; mIdx++) {
                        mos[mIdx]['exp'] = true;
                        var pris = mos[mIdx]['priorities'];
                        for (var pIdx = 0; pIdx < pris.length; pIdx++) {
                            pris[pIdx]['exp'] = true;
                        }
                    }
                }
            }
            if (ret.length) {
                var yr = ret[0];
                yr['exp'] = true;
                if (yr['months'].length) {
                    var mo = yr['months'][0];
                    mo['exp'] = true;
                    if (mo['priorities'].length) {
                        mo['priorities'][0]['exp'] = true;
                        // set default activity (initial...)
                        // initial activity under priority 0 ...
                        var defAct = mo['priorities'][0]['activities'][0];
                        // find default activity by id
                        var defActIDSaved = localStorage.getItem("defActID");
                        if (defActIDSaved) {
                            var defActID_1 = Number(defActIDSaved);
                            var defActTemp = this.datActivities.find(function (act) { return act.id == defActID_1; });
                            if (defActTemp)
                                defAct = defActTemp; // if saved activiity id search is successful
                        }
                        this.scatterActivity(defAct);
                    }
                }
                this.onActivityLoadAction = false;
            }
        }
        this.activityGroupings = ret;
        this.searchActivity(findAct);
        return this.activityGroupings;
    };
    AppOplpanDataService.prototype.searchActivity = function (act) {
        var _this = this;
        setTimeout(function () {
            if (act == undefined)
                act = _this.currActivity;
            if (act.id == -1)
                return;
            var grp = _this.activityGroupings;
            // get year
            var yr = grp.find(function (e) {
                return (e.year == __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](act.endDate).getFullYear());
            });
            if (yr != undefined) {
                yr.exp = true;
                var mo = yr['months'].find(function (e) {
                    return (e.month == __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](act.endDate).getMonth());
                });
                if (mo != undefined) {
                    mo.exp = true;
                    var pri = mo['priorities'].find(function (e) {
                        return (e.priority == act.type);
                    });
                    if (pri != undefined)
                        pri.exp = true;
                }
            }
        }, 100);
    };
    AppOplpanDataService.prototype.setOtherItemGroupings = function () {
        var _this = this;
        var module = this.udf.currentModule;
        var inputArray = (module == 'Teams' ? this.datTeams : this.datActivities);
        this.itemsGroupings = __WEBPACK_IMPORTED_MODULE_4__app_globals__["groupBy"](inputArray.filter(function (item) {
            return item['name'].toLowerCase()
                .indexOf(_this.udf.searchText.toLocaleLowerCase()) != -1;
        }), 'group', undefined, undefined, undefined, 'order');
        if (module == 'Activities') {
            this.getActivityGroupings();
        }
        this.setItemGroupInitialExpand();
        if (this.udf.currentModule == 'Teams') {
            // scatter first team of the first group on first time loading
            if (this.itemsGroupings.length > 0) {
                if (this.itemsGroupings[0]['data'].length > 0 && this.currTeamObj.id == -1) {
                    this.scatterTeam(this.itemsGroupings[0]['data'][0]);
                }
            }
            // set user lookup group
            this.setTeamUserSelection();
        }
        else if (this.udf.currentModule == 'Activities') {
            // scatter first activity of the first group on first time loading
            if (this.itemsGroupings.length > 0) {
                if (this.itemsGroupings[0]['data'].length > 0 && this.currActivityObj.id == -1) {
                    // set default activity ...
                    this.scatterActivity(this.itemsGroupings[0]['data'][0]);
                }
            }
            this.userSelection = __WEBPACK_IMPORTED_MODULE_4__app_globals__["groupBy"](this.datPersonnel.filter(function (item) {
                return item['id'] != 0;
            }), "nameGroup");
        }
        else {
            if (this.itemsGroupings.length > 0) {
                if (this.itemsGroupings[0]['data'].length > 0 && this.currPersonnelObj.id == -1) {
                    this.scatterPersonnel(this.itemsGroupings[0]['data'][0]);
                }
            }
        }
    };
    AppOplpanDataService.prototype.setItemGroupInitialExpand = function () {
        if (this.itemsGroupings.length) {
            this.udf.currGroup = this.itemsGroupings[0]['name'];
        }
    };
    AppOplpanDataService.prototype.setTempMessage = function (msgVarName, message, timeOut) {
        var _this = this;
        if (msgVarName == undefined)
            msgVarName = "msgTempPrompt";
        this[msgVarName] = message;
        if (timeOut == undefined)
            timeOut = 2000;
        setTimeout(function () { _this[msgVarName] = ''; }, timeOut);
    };
    AppOplpanDataService.prototype.getSelectedBackground = function (item) {
        var retVal = item['id'] == this.currPersonnelObj.id ?
            this.udf.selectedTextHilight : 'white';
        return retVal;
    };
    AppOplpanDataService.prototype.getSelectedColor = function (item) {
        var retVal = item['id'] == this.currPersonnelObj.id ? 'white' : 'black';
        return retVal;
    };
    AppOplpanDataService.prototype.getSelectedStyle = function (item) {
        var curVarName = this.udf.currentModule == 'Teams' ? 'currTeamObj' :
            (this.udf.currentModule == 'Activities' ? 'currActivityObj' : 'currPersonnelObj');
        var isCurrent = item['id'] == this[curVarName]['id'];
        var retVal = {
            'background': isCurrent ? this.udf.selectedTextHilight : 'white',
            'color': isCurrent ? 'white' : 'black',
        };
        return retVal;
    };
    AppOplpanDataService.prototype.isRequesting = function () {
        return (this.msgRequestPrompt.length != 0);
    };
    AppOplpanDataService.prototype.clearData = function (module) {
        /*
          datCompany:Array<object>=[];
          idxCompany:Object={};
        
          datPosition:Array<object>=[];
          idxPosition:Object={};
        
          idxPersonnel:Object={};
          datPersonnel:Array<dataPersonnel>=[];
        
          idxActivity:Object={};
          datActivities:Array<dataActivity>=[];
        
          idxTeam:Object={};
          datTeams:Array<dataTeam>=[];
        
        
        */
        if (module == 'Teams') {
        }
    };
    AppOplpanDataService.prototype.processData = function (args) {
        var _this = this;
        this.clearSubscription(args.subscription);
        //subscription
        if (args.beforePost != undefined)
            args.beforePost();
        if (args.isPosting == undefined)
            args.isPosting = true;
        if (args.processMessage == undefined) {
            args.processMessage = args.isPosting ? 'Posting data. Please wait...' : 'Processing data. Please wait...';
        }
        if (args.isPosting) {
            // writing data...
            this.setPosting(args.processMessage);
        }
        else {
            // getting data ...
            this.setLoading(args.processMessage);
        }
        return this.http.post(__WEBPACK_IMPORTED_MODULE_4__app_globals__["getUrlRequest"]("oplanData", args.method), JSON.stringify(args.data))
            .subscribe(function (data) {
            _this.msgRequestPrompt = '';
            _this.msgActivityUpdateError = '';
            if (args.onDataPost != undefined)
                args.onDataPost(data);
            if (data == null) {
                _this.msgActivityUpdateError = "Unknown error occured!!!";
            }
            else {
                if (data['stat'] == 'fail') {
                    _this.msgActivityUpdateError = "An error has occured when saving data ...<br/>" + data['message'];
                }
                else {
                    // this will update the current display with the new values after savig data
                    if (args.onSuccess != undefined)
                        args.onSuccess(data);
                }
            }
            if (args.isPosting) {
                _this.noPosting();
            }
            else {
                _this.noLoading();
            }
        }, function (err) {
            console.log("ERROR PROCESSING: ", err);
            if (args.onError != undefined && args.onError != null) {
                args.onError(err);
            }
            _this.withError(err);
            return null;
        });
    };
    AppOplpanDataService.prototype.valStartEnd = function (sd, ed, mem) {
        var retVal = {};
        retVal.noStart = (sd == undefined || sd == null);
        if (!retVal.noStart)
            retVal.noStart = (sd.length == 0);
        retVal.noEnd = (ed == undefined || ed == null);
        if (!retVal.noEnd)
            retVal.noEnd = (ed.length == 0);
        retVal.withDates = !retVal.noStart && !retVal.noEnd;
        if (retVal.withDates) {
            var sdFmt = __WEBPACK_IMPORTED_MODULE_4__app_globals__["stringToDate"](__WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](sd));
            var edFmt = __WEBPACK_IMPORTED_MODULE_4__app_globals__["stringToDate"](__WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](ed));
            retVal.invalid = (edFmt < sdFmt);
            if (retVal.invalid)
                retVal.msg = 'End date cannot be earlier than the start date.';
        }
        else {
            if (retVal.noStart)
                retVal.msg = 'start date';
            if (retVal.noEnd)
                retVal.msg += (retVal.msg.length ? ', ' : '') + 'end date';
        }
        if (mem != undefined) {
            // called from activity member validation, check if dates are within the range of activity
            var act = this.currActivity;
            var actStart = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](act.startDate);
            var actEnd = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](act.endDate);
            var memStart = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](mem.startDate);
            var memEnd = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](mem.endDate);
            retVal.outBound = (actStart > memStart || actEnd < memEnd);
        }
        else {
            retVal.outBound = false;
        }
        return retVal;
    };
    AppOplpanDataService.prototype.sortActivityMembers = function (inputArray) {
        var _this = this;
        inputArray.forEach(function (elem) {
            elem['sort'] = _this.activityMemberSortString(elem);
        });
        __WEBPACK_IMPORTED_MODULE_4__app_globals__["sortBy"](inputArray, "sort");
    };
    AppOplpanDataService.prototype.activityMemberSortString = function (mem) {
        var team = this.idxTeam['r' + mem.teamID];
        var per = this.idxPersonnel['r' + mem.perID];
        var grpName = "";
        var grp = this.config.teamGroupsArr.find(function (e) { return e.id == mem.grpID; });
        if (grp)
            grpName = grp.name;
        /*
        let act:dataActivity = this.idxActivity['r'+mem.idAct];
    
        let sortStr:string=""
    
        if(team.core){
          sortStr="0" + G.pad0(team.order,2);
        }else{
          sortStr="1" + G.pad0(team.order,2);
        }*/
        return (team.core ? "0" : "1") + __WEBPACK_IMPORTED_MODULE_4__app_globals__["pad0"](team.order, 2) + " " +
            (grpName + " ".repeat(50)).substr(0, 50) +
            (per.fullName + " ".repeat(50)).substr(0, 50) +
            __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToShortDate"](mem.startDate);
    };
    AppOplpanDataService.prototype.cloneObject = function (obj) {
        return (JSON.parse(JSON.stringify(obj)));
    };
    AppOplpanDataService.prototype.getCompanyName = function (coyID) {
        if (this.idxCompany) {
            var coy = this.idxCompany['r' + coyID];
            if (coy)
                return coy.name;
            else
                return coyID;
        }
        else {
            return coyID;
        }
    };
    AppOplpanDataService.prototype.getPositionName = function (posID) {
        if (this.idxPosition) {
            var pos = this.idxPosition['r' + posID];
            if (pos)
                return pos.name;
            else
                return posID;
        }
        else {
            return posID;
        }
    };
    AppOplpanDataService.prototype.setTeamUserSelection = function () {
        var _this = this;
        if (this.currTeam) {
            this.userSelection = __WEBPACK_IMPORTED_MODULE_4__app_globals__["groupBy"](this.datPersonnel.filter(function (item) {
                var retVal;
                var obj = _this.currTeam.members.find(function (mem) { return mem.perID == item.id; });
                return (!obj) && item['id'] != 0;
            }), "nameGroup");
        }
        else {
            this.userSelection = __WEBPACK_IMPORTED_MODULE_4__app_globals__["groupBy"](this.datPersonnel.filter(function (item) {
                return item['id'] != 0;
            }), "nameGroup");
        }
    };
    AppOplpanDataService.prototype.getCredentialsX = function (force) {
        this.logInfo.activityAdd = true;
        this.logInfo.activityEdit = true;
        this.logInfo.activityDelete = true;
        this.logInfo.activityMemAdmin = true;
        this.logInfo.teamAdd = true;
        this.logInfo.teamEdit = true;
        this.logInfo.teamDelete = true;
        this.logInfo.teamMemAdmin = true;
        this.logInfo.personnelAdd = true;
        this.logInfo.personnelEdit = true;
        this.logInfo.personnelDelete = true;
        this.logInfo.fullName = "Administrator";
        this.logInfo.uid = 1;
        this.logInfo.login = "admin";
        this.logInfo.role = "Admin";
    };
    AppOplpanDataService.prototype.getCredentials = function (force) {
        var _this = this;
        if (force == undefined)
            force = false;
        if (this.logInfo.taken && !force)
            return undefined;
        this.subsCredentials = this.processData({
            //method:"getCredentials",
            method: "getUserInfo",
            data: { "action": "getdata" },
            isPosting: false,
            processMessage: "Authenticating user...",
            subscription: this.subsCredentials,
            onSuccess: function (data) {
                if (data['user'] != undefined) {
                    alert("Test user: " + data['user']);
                    return undefined;
                }
                if (data['error'] != undefined) {
                    alert(data['error']);
                }
                else {
                    _this.logInfo = _this.cloneObject(data);
                    _this.logInfo.taken = true;
                    console.log("LogInfo:", _this.logInfo);
                }
            }
        });
        return this.subsCredentials;
    };
    AppOplpanDataService.prototype.getActivityConflicts = function (resetAll) {
        /*
        if(resetAll==undefined)resetAll=false;
        if(resetAll){
          this.datActivities.forEach((actObj:dataActivity)=>{
              actObj.conflicts=null;
          })
          this.currActivity.conflicts=null;
        }
    
    
        if(act.conflicts!=null)return;*/
        var _this = this;
        this.clearSubscription(this.subsActivityConflicts, "subsActivityConflicts");
        var act = this.currActivity;
        this.setLoading("Loading activity conflicts. Please wait...");
        this.subsActivityConflicts = this.http
            .get(__WEBPACK_IMPORTED_MODULE_4__app_globals__["getUrlRequest"]("oplanData", "getConflict", "activityID=" + act.id))
            .subscribe(function (data) {
            var arrConf = [];
            data['rows'].forEach(function (element) {
                var conf = new __WEBPACK_IMPORTED_MODULE_3__models_data_models__["a" /* activityConflict */]();
                conf.memberId = Number(element[0]);
                conf.personnelId = Number(element[1]);
                conf.activityId1 = Number(element[2]);
                conf.teamId1 = Number(element[3]);
                conf.startDate1 = __WEBPACK_IMPORTED_MODULE_4__app_globals__["stringToDate"](element[4]);
                conf.endDate1 = __WEBPACK_IMPORTED_MODULE_4__app_globals__["stringToDate"](element[5]);
                conf.activityId2 = Number(element[6]);
                conf.teamId2 = Number(element[7]);
                conf.startDate2 = __WEBPACK_IMPORTED_MODULE_4__app_globals__["stringToDate"](element[8]);
                conf.endDate2 = __WEBPACK_IMPORTED_MODULE_4__app_globals__["stringToDate"](element[9]);
                conf.isContinuation = (element[10] != 0);
                conf.isContinuationMob = (element[11] != 0);
                conf.isContinuationDemob = (element[12] != 0);
                if (element[9] == '20170411') {
                    var dt = __WEBPACK_IMPORTED_MODULE_4__app_globals__["stringToDate"](element[9]);
                }
                arrConf.push(conf);
            });
            _this.currActivityComp = [];
            var tmpComp = [];
            // collect competency list of all members
            data['comp'].forEach(function (comp) {
                tmpComp.push({
                    memid: comp[0],
                    pid: comp[1],
                    cid: comp[2],
                    issue: __WEBPACK_IMPORTED_MODULE_4__app_globals__["shortToLongDate"](comp[3]),
                    expiry: __WEBPACK_IMPORTED_MODULE_4__app_globals__["shortToLongDate"](comp[4]),
                    notTaken: (comp[5] == 1),
                    noExpiry: (comp[6] == 1),
                    expired: (comp[7] == 1),
                    lastDayToday: (comp[8] == 1),
                    exp7D: (comp[9] == 1),
                    exp30D: (comp[10] == 1),
                });
            });
            // group competencies by personnel id
            //act.competencies = G.groupBy(tmpComp,"pid");
            act.competencies = __WEBPACK_IMPORTED_MODULE_4__app_globals__["groupBy"](tmpComp, "memid");
            // sort compentencies by competency id
            act.competencies.forEach(function (c) {
                __WEBPACK_IMPORTED_MODULE_4__app_globals__["sortBy"](c.data, "cid");
            });
            // process conflicts
            var confGroup = __WEBPACK_IMPORTED_MODULE_4__app_globals__["groupBy"](arrConf, 'memberId');
            act.conflicts = {};
            confGroup.forEach(function (elem) {
                act.conflicts['r' + elem['name']] = elem;
            });
            _this.assignConflictMessages(act);
            _this.noLoading();
        }, function (err) {
            _this.withError(err);
            return undefined;
        });
        return this.subsActivityConflicts;
    };
    AppOplpanDataService.prototype.getCertTitle = function (certId) {
        return this.datCertificates.find(function (e) {
            return e.id == certId;
        }).name;
    };
    AppOplpanDataService.prototype.preMobTitle = function (title, cert, header, withDates) {
        if (withDates == undefined)
            withDates = true;
        return title + (title.length == 0 ? header : "") + "\n" +
            " - " + this.getCertTitle(cert.cid) +
            (withDates ?
                cert.issue.length == 0 && cert.expiry.length == 0 ? '' :
                    ' (' + (cert.issue ? 'issued: ' + cert.issue : '') +
                        (cert.expiry ? (cert.issue ? ', ' : '') + 'expiry: ' + cert.expiry : '') + ')' :
                "");
        //' (issued: ' + c.issue + ', expiry: '+c.expiry +')'
    };
    AppOplpanDataService.prototype.assignConflictMessages = function (act) {
        var _this = this;
        console.log('act.competencies...');
        console.log(act.competencies);
        // assign conflict messages to members
        act.members.forEach(function (mem) {
            var conf = act.conflicts['r' + mem['id']];
            mem.conflicts = null;
            mem.confMessage = "";
            mem.infoMessage = "";
            mem.warnMessage = "";
            mem.compWarnCompliant = 0;
            mem.compWarnExp30Days = 0;
            mem.compWarnExp7Days = 0;
            mem.compWarnExpired = 0;
            mem.compWarnLastDay = 0;
            mem.compWarnNotTaken = 0;
            mem.compWarnNotTakenTitle = "";
            mem.compWarnExpiredTitle = "";
            mem.compWarnExp7DaysTitle = "";
            mem.compWarnExp30DaysTitle = "";
            mem.compWarnLastDayTitle = "";
            mem.compWarnCompliantTitle = "";
            mem.compTitle = "";
            if (act.competencies) {
                var certs = act.competencies.find(function (e) {
                    //return (e.type==String(mem['perID']));
                    return (e.type == String(mem['id']));
                });
                if (certs) {
                    certs.data.forEach(function (c) {
                        // perform summation of status
                        if (c.notTaken) {
                            mem.compWarnNotTaken++;
                            mem.compWarnNotTakenTitle = _this.preMobTitle(mem.compWarnNotTakenTitle, c, "Required course(s) not yet taken:");
                        }
                        if (c.exp7D) {
                            mem.compWarnExp7Days++;
                            mem.compWarnExp7DaysTitle = _this.preMobTitle(mem.compWarnExp7DaysTitle, c, "Expired course(s) within 7 days from last day of assignment:");
                        }
                        if (c.exp30D) {
                            mem.compWarnExp30Days++;
                            mem.compWarnExp30DaysTitle = _this.preMobTitle(mem.compWarnExp30DaysTitle, c, "Expired course(s) within 30 days from last day of assignment:");
                        }
                        if (c.expired) {
                            mem.compWarnExpired++;
                            mem.compWarnExpiredTitle = _this.preMobTitle(mem.compWarnExpiredTitle, c, "Expired course(s):");
                        }
                        if (c.lastDayToday) {
                            mem.compWarnLastDay++;
                            mem.compWarnLastDayTitle = _this.preMobTitle(mem.compWarnLastDayTitle, c, "Course(s) expiring on last day of assignment:");
                        }
                        if (!c.notTaken && !c.expired && !c.exp7D && !c.exp30D) {
                            mem.compWarnCompliant++;
                            mem.compWarnCompliantTitle = _this.preMobTitle(mem.compWarnCompliantTitle, c, "Compliant:");
                        }
                    });
                }
                else {
                    mem.compTitle = "No pre-mobilization requirement certification found.";
                }
            }
            mem.warnCount = 0;
            mem.infoCount = 0;
            mem.confCount = 0;
            //this.currActivityComp
            if (conf != undefined) {
                mem.conflicts = conf;
                // iterate to each conflict data
                conf.data.forEach(function (con) {
                    var act2 = _this.idxActivity['r' + con.activityId2];
                    var isContFrom = (__WEBPACK_IMPORTED_MODULE_4__app_globals__["dateToString"](con.startDate1) == __WEBPACK_IMPORTED_MODULE_4__app_globals__["dateToString"](__WEBPACK_IMPORTED_MODULE_4__app_globals__["addDays"](con.endDate2, 1)));
                    if (con.isContinuation && !(con.isContinuationDemob || con.isContinuationMob)) {
                        // warning only that the person was also assigned to another activity but only after or before the current one
                        if (isContFrom) {
                            mem.warnMessage += (mem.warnCount == 0 ? '' : '\n') + 'Continuing from activity "' + act2.name + '" ending ' + __WEBPACK_IMPORTED_MODULE_4__app_globals__["dateToString"](__WEBPACK_IMPORTED_MODULE_4__app_globals__["addDays"](con.startDate1, -1), true) + '.';
                            mem.warnCount++;
                        }
                        else {
                            mem.warnMessage += (mem.warnCount == 0 ? '' : '\n') + 'Moving to activity "' + act2.name + '" starting ' + __WEBPACK_IMPORTED_MODULE_4__app_globals__["dateToString"](con.startDate2, true) + '.';
                            mem.warnCount++;
                        }
                    }
                    else if (con.isContinuation) {
                        // mob/demob was set at the continuation date
                        mem.confMessage += (mem.confCount == 0 ? '' : '\n') + 'Mob\/Demob switch incorrectly set for continuing assignment ' +
                            (isContFrom ? 'from' : 'to') + ' "' + act2.name + '" on ' + __WEBPACK_IMPORTED_MODULE_4__app_globals__["dateToString"]((isContFrom ? con.endDate2 : __WEBPACK_IMPORTED_MODULE_4__app_globals__["addDays"](con.endDate1, 1)), true) + '.';
                        mem.confCount++;
                    }
                    else {
                        // date conflict assignment
                        if (con.activityId1 == con.activityId2) {
                            // same activity conflict
                            mem.confMessage += (mem.confCount == 0 ? '' : '\n') + 'Assignment period conflicts with another assigment in the current activity.';
                            mem.confCount++;
                        }
                        else {
                            // other activity
                            mem.confMessage += (mem.confCount == 0 ? '' : '\n') + 'Assignment period conflicts with another assigment in activity "' + act2.name + '".';
                            mem.confCount++;
                        }
                    }
                });
            }
        });
    };
    AppOplpanDataService.prototype.loadActivity = function (act) {
        var _this = this;
        localStorage.setItem("defActID", String(act.id));
        this.setLoading("Loading activity " + act.title + ". Please wait...");
        setTimeout(function () {
            _this.onActivityLoadAction = true;
            _this.setOtherNavGroup();
            localStorage.setItem("currentModule", "Activities");
            _this.udf.currentModule = "Activities";
        }, 100);
    };
    AppOplpanDataService.prototype.removeConfig = function (id) {
        var _this = this;
        var cfgType = id.substr(0, 4);
        var keyIndex = parseInt(id.substr(4));
        var objName;
        switch (cfgType) {
            case "SPOB":
                objName = "stdPOBLimit";
                break;
            case "UPOB":
                objName = "upmPOBLimit";
                break;
            case "UPER":
                objName = "upmanPeriods";
                break;
            case "MLIM":
                objName = "mobLimit";
                break;
            case "DLIM":
                objName = "demobLimit";
                break;
            default:
        }
        var cfg = this.config[objName];
        if (cfg == undefined)
            return undefined;
        var objArr = cfg[cfgType == 'UPER' ? 'periods' : 'limits'];
        var obj = objArr[keyIndex];
        this.isPosting = true;
        this.msgRequestPrompt = "Removing configuration entry from " + cfg['title'] + ". Please wait...";
        this.subsProcessConfig = this.processData({
            method: "deleteConfig",
            data: {
                type: cfgType,
                id: obj.id
            },
            isPosting: true,
            processMessage: this.msgRequestPrompt,
            subscription: this.subsProcessConfig,
            onSuccess: function (data) {
                _this.msgRequestPrompt = "";
                _this.setTempMessage("msgTempInfoPrompt", "Configuration setting was succesfully removed under " + cfg['title']);
                var index = objArr.map(function (e) { return e.id; }).indexOf(obj.id);
                objArr.splice(index, 1);
                _this.arrangeConfig();
                _this.isPosting = false;
            }
        });
        return this.subsProcessConfig;
    };
    AppOplpanDataService.prototype.calActivitiesInitVars = function (data) {
        //if(this.calActivities!=undefined)this.calActivities.initVars(data);
    };
    AppOplpanDataService.prototype.trackByItem = function (idx, item) {
        return idx;
    };
    AppOplpanDataService.prototype.clearSubscription = function (dataSubscription, subsName) {
        if (subsName == undefined)
            subsName = "";
        if (dataSubscription) {
            //alert('unsubscribing!!!!, ' +subsName);
            dataSubscription.unsubscribe();
        }
    };
    AppOplpanDataService.prototype.clearAllSubscription = function () {
        this.clearSubscription(this.subsActivities);
        this.clearSubscription(this.subsActivityConflicts);
        this.clearSubscription(this.subsConfig);
        this.clearSubscription(this.subsCredentials);
        this.clearSubscription(this.subsDeleteTeam);
        this.clearSubscription(this.subsPersonnel);
        this.clearSubscription(this.subsProcessConfig);
        this.clearSubscription(this.subsSaveActivity);
        this.clearSubscription(this.subsSaveActivityMember);
        this.clearSubscription(this.subsSavePersonnel);
        this.clearSubscription(this.subsSaveTeam);
        this.clearSubscription(this.subsTeams);
    };
    AppOplpanDataService = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Injectable"])(),
        __metadata("design:paramtypes", [typeof (_a = typeof __WEBPACK_IMPORTED_MODULE_2__angular_common_http__["a" /* HttpClient */] !== "undefined" && __WEBPACK_IMPORTED_MODULE_2__angular_common_http__["a" /* HttpClient */]) === "function" && _a || Object, typeof (_b = typeof __WEBPACK_IMPORTED_MODULE_1__angular_http__["a" /* Http */] !== "undefined" && __WEBPACK_IMPORTED_MODULE_1__angular_http__["a" /* Http */]) === "function" && _b || Object, typeof (_c = typeof __WEBPACK_IMPORTED_MODULE_5__app_udfs__["AppUDFService"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_5__app_udfs__["AppUDFService"]) === "function" && _c || Object])
    ], AppOplpanDataService);
    return AppOplpanDataService;
    var _a, _b, _c;
}());

//# sourceMappingURL=app.oplandata.js.map

/***/ }),

/***/ "../../../../../src/app/app.udfs.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
Object.defineProperty(__webpack_exports__, "__esModule", { value: true });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "AppUDFService", function() { return AppUDFService; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__angular_common_http__ = __webpack_require__("../../../common/@angular/common/http.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2_jquery__ = __webpack_require__("../../../../jquery/dist/jquery.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2_jquery___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_2_jquery__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3__app_globals__ = __webpack_require__("../../../../../src/app/app.globals.ts");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};




var AppUDFService = (function () {
    function AppUDFService(httpClient) {
        //calEndDate
        //calStartDate=new Date(2017,1,1)
        this.httpClient = httpClient;
        this.menuBackground = '#f7d117';
        this.selectedTextHilight = '#296192';
        //f7f7f9
        this.cardHeaderBackground = '#eceeef';
        this.cardHeaderBorder = '#d9d9d9';
        this.tableGroup1Back = '#EED674';
        this.tableGroup1Fore = 'black';
        this.defColorDanger = "#d9534f";
        this.searchText = "";
        //currentModule:string='By Name';
        //currentModule:string='Activities';
        //currentModule:string='Calendar';
        this.currentModule = 'Calendar';
        //currentModule:string='Teams';
        //currentModule:string='Settings';
        this.currGroup = '';
        this.bannerHeight = 100;
        this.bannerMenuHeight = 37;
        this.calScopeLimit = 730;
        this.calcDefaultOffset = 90;
        //calStartDate:Date=new Date(2017,1,1);
        this.profScopeLimit = this.calScopeLimit;
        //isIEOrEdge:boolean=true;
        this.isIEOrEdge = /msie\s|trident\/|edge\//i.test(window.navigator.userAgent);
        this.isIEOnly = /msie\s|trident\/|xxxx\//i.test(window.navigator.userAgent);
        //calStartDate:Date=new Date();
        //calEndDate:Date=new Date(this.calStartDate.getFullYear(),
        //                        this.calStartDate.getMonth(),
        //                        this.calStartDate.getDate() + this.calcDefaultOffset);
        this.calStartDate = new Date();
        this.calEndDate = new Date();
        this.calFilterDate = null;
        this.calFilterDateObj = null;
        this.calFilterPOB = true;
        this.calFilterMOB = true;
        this.calFilterDEMOB = true;
        this.calFilterNIGHT = true;
        this.calHoverDate = null;
        this.calActivitySite = "SWP";
        this.calActivityMonth = new Date().getMonth();
        this.calActivityYear = new Date().getFullYear();
        this.calActivitySiteMTIAP = this.calActivitySite;
        this.calActivityMTIAPYear = new Date().getFullYear();
        this.calActivityMTIAPPage = 1;
        this.calActivityMTIAPPagePerYear = 4; // 2,3,4,6,12
        this.calActivityMTIAPPeriods = [];
        this.calActivityChartScale = 1.0;
        this.calActivityChartScaleMTIAP = 1.0;
        this.calActivitySiteDisp = this.calActivitySite;
        this.calActivityMonthDisp = this.calActivityMonth;
        this.calActivityYearDisp = this.calActivityYear;
        this.calActivityMTIAPPagePerYearDisp = this.calActivityMTIAPPagePerYear;
        this.calActivityMTIAPYearDisp = this.calActivityMTIAPYear;
        this.calActivitySiteMTIAPDisp = this.calActivitySiteMTIAP;
        this.calDispStartDate = __WEBPACK_IMPORTED_MODULE_3__app_globals__["dateToString"](this.calStartDate, true);
        this.calDispEndDate = __WEBPACK_IMPORTED_MODULE_3__app_globals__["dateToString"](this.calEndDate, true);
        this.modePersonnel = "normal"; // normal, new, edit
        this.modeActivity = "normal"; // normal, new, edit, members, delalert, saving, deleting
        this.modeActivityMember = "normal"; // normal, new, edit, delalert, saving, deleting
        this.modeTeam = "normal"; // normal, new, edit, members
        this.modeTeamMember = "normal"; // normal, new, edit
        this.datePickerConfig = {
            format: "DD MMM YYYY",
            locale: "en",
            appendTo: document.body
        };
        var savedModule = localStorage.getItem("currentModule");
        var stiapSite = localStorage.getItem("stiapSite");
        var stiapYear = localStorage.getItem("stiapYear");
        var stiapMonth = localStorage.getItem("stiapMonth");
        var stiapScale = localStorage.getItem("stiapScale");
        if (stiapSite)
            this.calActivitySite = stiapSite;
        if (stiapYear)
            this.calActivityYear = Number(stiapYear);
        if (stiapMonth)
            this.calActivityMonth = Number(stiapMonth);
        if (stiapScale)
            this.calActivityChartScale = Number(stiapScale);
        this.calActivitySiteDisp = this.calActivitySite;
        this.calActivityMonthDisp = this.calActivityMonth;
        this.calActivityYearDisp = this.calActivityYear;
        var mtiapSite = localStorage.getItem("mtiapSite");
        var mtiapYear = localStorage.getItem("mtiapYear");
        var mtiapPage = localStorage.getItem("mtiapPage");
        if (mtiapSite)
            this.calActivitySiteMTIAP = mtiapSite;
        if (mtiapYear)
            this.calActivityMTIAPYear = Number(mtiapYear);
        //if(mtiapPage)this.calActivityMTIAPPage = Number(mtiapPage);
        this.calActivitySiteMTIAPDisp = this.calActivitySiteMTIAP;
        this.calActivityMTIAPYearDisp = this.calActivityMTIAPYear;
        if (savedModule) {
            if (savedModule == "Personnel")
                savedModule = "By Name";
            //if(savedModule=="Teams" || savedModule=="MTIAP")savedModule="Calendar";
            //if(savedModule=="MTIAP")savedModule="Calendar";
            this.currentModule = savedModule;
        }
        else {
            this.currentModule = 'Calendar';
        }
    }
    AppUDFService.prototype.resetPOBProfileScope = function () {
        var tmpStart = new Date();
        this.profStartDate = new Date(tmpStart.getFullYear(), tmpStart.getMonth(), 1);
        this.profEndDate = new Date(tmpStart.getFullYear(), tmpStart.getMonth() + 1, 0);
        this.profDispStartDate = __WEBPACK_IMPORTED_MODULE_3__app_globals__["dateToString"](this.profStartDate, true);
        this.profDispEndDate = __WEBPACK_IMPORTED_MODULE_3__app_globals__["dateToString"](this.profEndDate, true);
    };
    AppUDFService.prototype.isActivityNormal = function (mode) {
        if (mode != undefined) {
            if (mode)
                this.modeActivity = "normal";
        }
        return (this.modeActivity == "normal");
    };
    AppUDFService.prototype.isActivityEditing = function (mode) {
        if (mode != undefined) {
            if (mode)
                this.modeActivity = "edit";
        }
        return (this.modeActivity == "edit");
    };
    AppUDFService.prototype.isActivityNew = function (mode) {
        if (mode != undefined) {
            if (mode)
                this.modeActivity = "new";
        }
        return (this.modeActivity == "new");
    };
    AppUDFService.prototype.isActivityDelete = function (mode) {
        if (mode != undefined) {
            if (mode)
                this.modeActivity = "deleting";
        }
        return (this.modeActivity == "deleting");
    };
    AppUDFService.prototype.isActivityDeleteAlert = function (mode) {
        if (mode != undefined) {
            if (mode)
                this.modeActivity = "delalert";
        }
        return (this.modeActivity == "delalert");
    };
    AppUDFService.prototype.isActivityMemberNormal = function (mode) {
        if (mode != undefined) {
            if (mode)
                this.modeActivityMember = "normal";
        }
        return (this.modeActivityMember == "normal");
    };
    AppUDFService.prototype.isActivityMemberEdit = function (mode) {
        if (mode != undefined) {
            if (mode)
                this.modeActivityMember = "edit";
        }
        return (this.modeActivityMember == "edit");
    };
    AppUDFService.prototype.isActivityMemberNew = function (mode) {
        if (mode != undefined) {
            if (mode)
                this.modeActivityMember = "new";
        }
        return (this.modeActivityMember == "new");
    };
    AppUDFService.prototype.isActivityMemberDeleteAlert = function (mode) {
        if (mode != undefined) {
            if (mode)
                this.modeActivityMember = "delmemalert";
        }
        return (this.modeActivityMember == "delmemalert");
    };
    AppUDFService.prototype.isActivityMemberDelete = function (mode) {
        if (mode != undefined) {
            if (mode)
                this.modeActivityMember = "deleting";
        }
        return (this.modeActivityMember == "deleting");
    };
    AppUDFService.prototype.condSize = function (wdIE, wdElse) {
        return this.isIEOnly ? wdIE : wdElse;
    };
    AppUDFService.prototype.weekCount = function (year, month_number) {
        // month_number is in the range 1..12
        var firstOfMonth = new Date(year, month_number - 1, 1);
        var lastOfMonth = new Date(year, month_number, 0);
        var used = firstOfMonth.getDay() + lastOfMonth.getDate();
        return Math.ceil(used / 7);
    };
    AppUDFService.prototype.weekCountX = function (yr, mo) {
        var ret = 1;
        var curDay = 1;
        var dt = new Date(yr, mo, curDay);
        var dte = new Date(yr, mo + 1, 0);
        var dow = dt.getDay();
        var i;
        for (i = 2; i < 7; i++) {
            curDay += (7 - dow + 1);
            dow = 0; // set to sunday;
            if (curDay > dte.getDate()) {
                break;
            }
            else {
                ret++;
            }
        }
        return ret;
    };
    AppUDFService.prototype.MTIAPPeriods = function (resetPriorities) {
        //  MTIAPPeriods():Array<any>{
        var elems = 2 * this.calActivityMTIAPPagePerYear;
        //let newPeriods:Boolean=(this.calActivityMTIAPPeriods.length==0);
        var mosPerYear = 12 / this.calActivityMTIAPPagePerYear;
        //if(resetPriorities==undefined)resetPriorities=false;
        //if(newPeriods){
        // create elements {sem:number,year:number}
        this.calActivityMTIAPPeriods = null;
        this.calActivityMTIAPPeriods = [];
        for (var e = 1; e <= elems; e++) {
            this.calActivityMTIAPPeriods.push({
                sem: e,
                year: this.calActivityMTIAPYear + (e <= elems / 2 ? 0 : 1),
                priorities: []
            });
        }
        //}
        /*if(this.calActivityMTIAPYear != this.calActivityMTIAPPeriods[0].year){
          for(let e=1;e<=elems;e++){
            this.calActivityMTIAPPeriods.push({
              sem:e, year:this.calActivityMTIAPYear + (e<=elems/2 ? 0 : 1)
            });
          }
        }*/
        /*if(resetPriorities && !newPeriods){
          for(let e=1;e<=elems;e++){
            this.calActivityMTIAPPeriods[e].priorities=[];
          }
        }*/
        return this.calActivityMTIAPPeriods;
    };
    /** Teams mode */
    AppUDFService.prototype.isTeamNormal = function (mode) {
        if (mode != undefined) {
            if (mode)
                this.modeTeam = "normal";
        }
        return (this.modeTeam == "normal");
    };
    AppUDFService.prototype.isTeamEditing = function (mode) {
        if (mode != undefined) {
            if (mode)
                this.modeTeam = "edit";
        }
        return (this.modeTeam == "edit");
    };
    AppUDFService.prototype.isTeamNew = function (mode) {
        if (mode != undefined) {
            if (mode)
                this.modeTeam = "add";
        }
        return (this.modeTeam == "add");
    };
    AppUDFService.prototype.isTeamDelete = function (mode) {
        if (mode != undefined) {
            if (mode)
                this.modeTeam = "deleting";
        }
        return (this.modeTeam == "deleting");
    };
    AppUDFService.prototype.isTeamDeleteAlert = function (mode) {
        if (mode != undefined) {
            if (mode)
                this.modeTeam = "delalert";
        }
        return (this.modeTeam == "delalert");
    };
    AppUDFService.prototype.isTeamMemberDeleteAlert = function (mode) {
        if (mode != undefined) {
            if (mode)
                this.modeTeam = "delmemalert";
        }
        return (this.modeTeam == "delmemalert");
    };
    /** Personnels mode */
    AppUDFService.prototype.isPersonnelDeleteAlert = function (mode) {
        if (mode != undefined) {
            if (mode)
                this.modePersonnel = "delalert";
        }
        return (this.modePersonnel == "delalert");
    };
    AppUDFService.prototype.isPersonnelNormal = function (mode) {
        if (mode != undefined) {
            if (mode)
                this.modePersonnel = "normal";
        }
        return (this.modePersonnel == "normal");
    };
    /*
    isTeamEditing(mode?:boolean):boolean{
      if(mode!undefined){
        if(mode)
      }
    }
    isTeamMemberEdit(mode?:boolean):boolean{
  
    }
    isTeamNew(mode?:boolean):boolean{
  
    }
    isTeamMemberNew(mode?:boolean):boolean{
  
    }*/
    AppUDFService.prototype.winWd = function () {
        return __WEBPACK_IMPORTED_MODULE_2_jquery__(window).width();
    };
    AppUDFService.prototype.winHt = function () {
        return __WEBPACK_IMPORTED_MODULE_2_jquery__(window).height();
    };
    AppUDFService.prototype.htNoBan = function (px) {
        if (px) {
            return (__WEBPACK_IMPORTED_MODULE_2_jquery__(window).height() - this.bannerHeight) + 'px';
        }
        else {
            return __WEBPACK_IMPORTED_MODULE_2_jquery__(window).height() - this.bannerHeight;
        }
    };
    AppUDFService.prototype.getColWidth = function (id, pos) {
        var tbl = __WEBPACK_IMPORTED_MODULE_2_jquery__('#' + id);
        if (tbl.length) {
            var cols = __WEBPACK_IMPORTED_MODULE_2_jquery__(tbl.children('tbody').children('tr')[0]).children('td');
            return __WEBPACK_IMPORTED_MODULE_2_jquery__(cols[pos]).width();
        }
        else {
            return 50;
        }
    };
    AppUDFService.prototype.getAbsWidthVals = function (arrInput, reference, scrollWidth) {
        //const scrollWidth:number=40;
        var i;
        var curVal;
        var iCurrentValue = 0;
        var iTotalAbsolute = 0;
        var iRemainingWidth = 0;
        var iValStr;
        var withPercent = false;
        var arr = [];
        if (scrollWidth == undefined)
            scrollWidth = 0;
        // get total absolute value
        for (i = 0; i < arrInput.length; i++) {
            iValStr = String(arrInput[i]);
            if (iValStr.indexOf('%') == -1) {
                // absolute value
                arr[i] = parseFloat(iValStr);
                iTotalAbsolute += arr[i];
            }
            else {
                withPercent = true;
            }
        }
        if (withPercent) {
            iRemainingWidth = reference - iTotalAbsolute - scrollWidth;
            for (i = 0; i < arrInput.length; i++) {
                iValStr = String(arrInput[i]);
                if (iValStr.indexOf('%') != -1) {
                    arr[i] = Math.abs(iRemainingWidth * parseFloat(iValStr) / 100);
                }
            }
        }
        /*    for(i=1;i<arr.length;i++){
              arr[i]=parseFloat(arr[i])+parseFloat(arr[i-1])
            }*/
        return arr;
    };
    AppUDFService.prototype.getReactiveTableHeight = function (tblID, cardID, offBottom, minHeight) {
        if (minHeight == undefined)
            minHeight = 200;
        var card = __WEBPACK_IMPORTED_MODULE_2_jquery__('#' + cardID);
        var tbl = __WEBPACK_IMPORTED_MODULE_2_jquery__('#' + tblID);
        if (card.length == 0 || tbl.length == 0)
            return minHeight;
        var cardPos = card.position();
        var cardOff = card.offset();
        var tblPos = card.position();
        var tblOff = card.offset();
        if (offBottom == undefined)
            offBottom = cardOff.top;
        var tblHead = tbl.children('thead');
        var tblBody = tbl.children('tbody');
        var tblTrueTop = tblBody.offset().top - cardOff.top;
        //let newHeight = $(window).height() - tblTrueTop - offBottom-this.bannerHeight
        var offFoot = 0;
        var ftr = __WEBPACK_IMPORTED_MODULE_2_jquery__('.card-footer');
        if (ftr.length)
            offFoot = ftr.height();
        var newHeight = __WEBPACK_IMPORTED_MODULE_2_jquery__(window).height() - tblTrueTop - offBottom - this.bannerHeight - offFoot;
        return newHeight > minHeight ? newHeight : minHeight;
    };
    AppUDFService.prototype.clearDate = function (dateInputID) {
        var inp = __WEBPACK_IMPORTED_MODULE_2_jquery__('#' + dateInputID);
    };
    AppUDFService.prototype.acnBtnClass = function (action, disable) {
        if (disable == undefined)
            disable = false;
        var retVal = {};
        switch (action) {
            case 'delete':
            case 'edit':
            case 'add':
                if (!disable)
                    disable = this.modePersonnel != 'normal';
                retVal = {
                    "disabled": disable
                };
                break;
            case 'save':
            case 'cancel':
                retVal = {
                    "disabled": (this.modePersonnel == 'normal'),
                    "btn-primary": (this.modePersonnel != 'normal')
                };
                break;
        }
        return retVal;
    };
    AppUDFService.prototype.isDisplayDataComponent = function () {
        return this.currentModule != "Calendar" &&
            this.currentModule != "POBProfile" &&
            this.currentModule != "CalendarActivity" &&
            this.currentModule != "CalendarActivityMTIAP" &&
            this.currentModule != "Test";
        //this.udf.currentModule
    };
    AppUDFService.prototype.dummy = function (e) {
        // dummy function to trigger window dimensional parameters updates
    };
    AppUDFService = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Injectable"])(),
        __metadata("design:paramtypes", [typeof (_a = typeof __WEBPACK_IMPORTED_MODULE_1__angular_common_http__["a" /* HttpClient */] !== "undefined" && __WEBPACK_IMPORTED_MODULE_1__angular_common_http__["a" /* HttpClient */]) === "function" && _a || Object])
    ], AppUDFService);
    return AppUDFService;
    var _a;
}());

//# sourceMappingURL=app.udfs.js.map

/***/ }),

/***/ "../../../../../src/app/components/alv-table-panel/alv-table-panel.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/alv-table-panel/alv-table-panel.component.html":
/***/ (function(module, exports) {

module.exports = "<ng-container *ngTemplateOutlet=\"headerTemplate\">\n\n</ng-container>"

/***/ }),

/***/ "../../../../../src/app/components/alv-table-panel/alv-table-panel.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return AlvTablePanelComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};

var AlvTablePanelComponent = (function () {
    function AlvTablePanelComponent() {
        this.showUL = false;
        this.showLL = false;
        this.showUR = false;
        this.showLR = false;
    }
    AlvTablePanelComponent.prototype.ngOnInit = function () {
    };
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Boolean)
    ], AlvTablePanelComponent.prototype, "showUL", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Boolean)
    ], AlvTablePanelComponent.prototype, "showLL", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Boolean)
    ], AlvTablePanelComponent.prototype, "showUR", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Boolean)
    ], AlvTablePanelComponent.prototype, "showLR", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", typeof (_a = typeof __WEBPACK_IMPORTED_MODULE_0__angular_core__["TemplateRef"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_0__angular_core__["TemplateRef"]) === "function" && _a || Object)
    ], AlvTablePanelComponent.prototype, "headerTemplate", void 0);
    AlvTablePanelComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-alv-table-panel',
            template: __webpack_require__("../../../../../src/app/components/alv-table-panel/alv-table-panel.component.html"),
            styles: [__webpack_require__("../../../../../src/app/components/alv-table-panel/alv-table-panel.component.css")]
        }),
        __metadata("design:paramtypes", [])
    ], AlvTablePanelComponent);
    return AlvTablePanelComponent;
    var _a;
}());

//# sourceMappingURL=alv-table-panel.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/alv-table-template-body/alv-table-template-body.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/alv-table-template-body/alv-table-template-body.component.html":
/***/ (function(module, exports) {

module.exports = "<p>\n  alv-table-template-body works!\n</p>\n"

/***/ }),

/***/ "../../../../../src/app/components/alv-table-template-body/alv-table-template-body.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return AlvTableTemplateBodyComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};

var AlvTableTemplateBodyComponent = (function () {
    function AlvTableTemplateBodyComponent() {
    }
    AlvTableTemplateBodyComponent.prototype.ngOnInit = function () {
    };
    AlvTableTemplateBodyComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-alv-table-template-body',
            template: __webpack_require__("../../../../../src/app/components/alv-table-template-body/alv-table-template-body.component.html"),
            styles: [__webpack_require__("../../../../../src/app/components/alv-table-template-body/alv-table-template-body.component.css")]
        }),
        __metadata("design:paramtypes", [])
    ], AlvTableTemplateBodyComponent);
    return AlvTableTemplateBodyComponent;
}());

//# sourceMappingURL=alv-table-template-body.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/alv-table-template-detail-row/alv-table-template-detail-row.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return AlvTableTemplateDetailRowComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__alv_table_template_alv_table_template_component__ = __webpack_require__("../../../../../src/app/components/alv-table-template/alv-table-template.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__alv_test_component_alv_test_component_component__ = __webpack_require__("../../../../../src/app/components/alv-test-component/alv-test-component.component.ts");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};
var __param = (this && this.__param) || function (paramIndex, decorator) {
    return function (target, key) { decorator(target, key, paramIndex); }
};



var AlvTableTemplateDetailRowComponent = (function () {
    function AlvTableTemplateDetailRowComponent(parent) {
        this.parent = parent;
    }
    AlvTableTemplateDetailRowComponent.prototype.ngOnInit = function () {
    };
    AlvTableTemplateDetailRowComponent.prototype.ngAfterViewInit = function () {
        console.log('ngAfterViewInit...' + this.tds.length);
        this.tds.forEach(function (tdInstance) { return console.log(tdInstance); });
    };
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["ViewChildren"])(__WEBPACK_IMPORTED_MODULE_2__alv_test_component_alv_test_component_component__["a" /* AlvTestComponentComponent */]),
        __metadata("design:type", typeof (_a = typeof __WEBPACK_IMPORTED_MODULE_0__angular_core__["QueryList"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_0__angular_core__["QueryList"]) === "function" && _a || Object)
    ], AlvTableTemplateDetailRowComponent.prototype, "tds", void 0);
    AlvTableTemplateDetailRowComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-alv-table-template-detail-row',
            template: "\n    <tr class=\"d-flex\" [style.width.px]=\"parent.rowWidth()\"><ng-content></ng-content></tr>\n    ",
            styles: []
        })
        //app-alv-test-component
        ,
        __param(0, Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Inject"])(__WEBPACK_IMPORTED_MODULE_1__alv_table_template_alv_table_template_component__["a" /* AlvTableTemplateComponent */])),
        __metadata("design:paramtypes", [typeof (_b = typeof __WEBPACK_IMPORTED_MODULE_1__alv_table_template_alv_table_template_component__["a" /* AlvTableTemplateComponent */] !== "undefined" && __WEBPACK_IMPORTED_MODULE_1__alv_table_template_alv_table_template_component__["a" /* AlvTableTemplateComponent */]) === "function" && _b || Object])
    ], AlvTableTemplateDetailRowComponent);
    return AlvTableTemplateDetailRowComponent;
    var _a, _b;
}());

//# sourceMappingURL=alv-table-template-detail-row.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/alv-table-template-detail/alv-table-template-detail.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/alv-table-template-detail/alv-table-template-detail.component.html":
/***/ (function(module, exports) {

module.exports = "<td [style.width.px]=\"parent.options.colDefs[index].width\"\r\n    [style.border]=\"'none'\"\r\n    [style.border-bottom]=\"'1px solid #e5e5e5'\"\r\n    [style.border-left]=\"'1px solid #e5e5e5'\"\r\n    ><ng-content></ng-content></td>"

/***/ }),

/***/ "../../../../../src/app/components/alv-table-template-detail/alv-table-template-detail.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return AlvTableTemplateDetailComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__alv_table_template_alv_table_template_component__ = __webpack_require__("../../../../../src/app/components/alv-table-template/alv-table-template.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2_jquery__ = __webpack_require__("../../../../jquery/dist/jquery.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2_jquery___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_2_jquery__);
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};
var __param = (this && this.__param) || function (paramIndex, decorator) {
    return function (target, key) { decorator(target, key, paramIndex); }
};



var AlvTableTemplateDetailComponent = (function () {
    function AlvTableTemplateDetailComponent(parent, elt, renderer) {
        this.parent = parent;
        this.elt = elt;
        this.renderer = renderer;
        this.index = -1;
        this.cell = null;
    }
    AlvTableTemplateDetailComponent.prototype.ngOnInit = function () {
        //console.log(this.cell.length);
    };
    AlvTableTemplateDetailComponent.prototype.ngAfterViewInit = function () {
        //var textNode = this.elt.nativeElement.childNodes[0];
        //var textInput = textNode.nodeValue;
        //this.renderer.setText(textNode, textInput.toUpperCase());
        this.cell = __WEBPACK_IMPORTED_MODULE_2_jquery__(this.elt.nativeElement).find('td');
    };
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Number)
    ], AlvTableTemplateDetailComponent.prototype, "index", void 0);
    AlvTableTemplateDetailComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-alv-table-template-detail',
            template: __webpack_require__("../../../../../src/app/components/alv-table-template-detail/alv-table-template-detail.component.html"),
            styles: [__webpack_require__("../../../../../src/app/components/alv-table-template-detail/alv-table-template-detail.component.css")]
        }),
        __param(0, Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Inject"])(__WEBPACK_IMPORTED_MODULE_1__alv_table_template_alv_table_template_component__["a" /* AlvTableTemplateComponent */])),
        __metadata("design:paramtypes", [typeof (_a = typeof __WEBPACK_IMPORTED_MODULE_1__alv_table_template_alv_table_template_component__["a" /* AlvTableTemplateComponent */] !== "undefined" && __WEBPACK_IMPORTED_MODULE_1__alv_table_template_alv_table_template_component__["a" /* AlvTableTemplateComponent */]) === "function" && _a || Object, typeof (_b = typeof __WEBPACK_IMPORTED_MODULE_0__angular_core__["ElementRef"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_0__angular_core__["ElementRef"]) === "function" && _b || Object, typeof (_c = typeof __WEBPACK_IMPORTED_MODULE_0__angular_core__["Renderer"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_0__angular_core__["Renderer"]) === "function" && _c || Object])
    ], AlvTableTemplateDetailComponent);
    return AlvTableTemplateDetailComponent;
    var _a, _b, _c;
}());

//# sourceMappingURL=alv-table-template-detail.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/alv-table-template-filter/alv-table-template-filter.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/alv-table-template-filter/alv-table-template-filter.component.html":
/***/ (function(module, exports) {

module.exports = "<p>\n  alv-table-template-filter works!\n</p>\n"

/***/ }),

/***/ "../../../../../src/app/components/alv-table-template-filter/alv-table-template-filter.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return AlvTableTemplateFilterComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};

var AlvTableTemplateFilterComponent = (function () {
    function AlvTableTemplateFilterComponent() {
    }
    AlvTableTemplateFilterComponent.prototype.ngOnInit = function () {
    };
    AlvTableTemplateFilterComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-alv-table-template-filter',
            template: __webpack_require__("../../../../../src/app/components/alv-table-template-filter/alv-table-template-filter.component.html"),
            styles: [__webpack_require__("../../../../../src/app/components/alv-table-template-filter/alv-table-template-filter.component.css")]
        }),
        __metadata("design:paramtypes", [])
    ], AlvTableTemplateFilterComponent);
    return AlvTableTemplateFilterComponent;
}());

//# sourceMappingURL=alv-table-template-filter.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/alv-table-template-header-row/alv-table-template-header-row.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return AlvTableTemplateHeaderRowComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__alv_table_template_alv_table_template_component__ = __webpack_require__("../../../../../src/app/components/alv-table-template/alv-table-template.component.ts");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};
var __param = (this && this.__param) || function (paramIndex, decorator) {
    return function (target, key) { decorator(target, key, paramIndex); }
};


var AlvTableTemplateHeaderRowComponent = (function () {
    function AlvTableTemplateHeaderRowComponent(parent) {
        this.parent = parent;
    }
    AlvTableTemplateHeaderRowComponent.prototype.ngOnInit = function () {
    };
    AlvTableTemplateHeaderRowComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-alv-table-template-header-row',
            template: "\n  <tr class=\"d-flex p-0 m-0\" [style.width.px]=\"parent.rowWidth()\"><ng-content></ng-content></tr>\n  ",
            styles: []
        }),
        __param(0, Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Inject"])(__WEBPACK_IMPORTED_MODULE_1__alv_table_template_alv_table_template_component__["a" /* AlvTableTemplateComponent */])),
        __metadata("design:paramtypes", [typeof (_a = typeof __WEBPACK_IMPORTED_MODULE_1__alv_table_template_alv_table_template_component__["a" /* AlvTableTemplateComponent */] !== "undefined" && __WEBPACK_IMPORTED_MODULE_1__alv_table_template_alv_table_template_component__["a" /* AlvTableTemplateComponent */]) === "function" && _a || Object])
    ], AlvTableTemplateHeaderRowComponent);
    return AlvTableTemplateHeaderRowComponent;
    var _a;
}());

//# sourceMappingURL=alv-table-template-header-row.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/alv-table-template-header/alv-table-template-header.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/alv-table-template-header/alv-table-template-header.component.html":
/***/ (function(module, exports) {

module.exports = "<th [style.width.px]=\"colWidth(index)\" [attr.colspan]=\"colspan\" \n  class=\"py-1 bg-primary text-white\" [ngClass]=\"ngClass()\" [style.border]=\"'none'\" \n    [style.border-bottom]=\"'1px solid #e5e5e5'\"\n    [style.border-left]=\"'1px solid #e5e5e5'\"\n    >\n  <ng-content></ng-content>\n  <!--div class=\"container-fluid d-flex justify-content-between p-0\" \n    \n    [style.cursor]=\"sortable ? 'pointer' : 'default'\"\n    (click)=\"headerClicked($event)\"\n    >\n    <div><ng-content></ng-content></div>\n    <div *ngIf=\"sortable\" class=\"px-1 \" style=\"position: relative;\">\n      <div *ngIf=\"false\" style=\"position:absolute;background: red;height:50%;\">&nbsp;</div>\n      <i class=\"fa fa-sort\" style=\"font-size:0.9em;\"></i>\n    </div>\n  </div-->\n</th>"

/***/ }),

/***/ "../../../../../src/app/components/alv-table-template-header/alv-table-template-header.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return AlvTableTemplateHeaderComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__alv_table_template_alv_table_template_component__ = __webpack_require__("../../../../../src/app/components/alv-table-template/alv-table-template.component.ts");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};
var __param = (this && this.__param) || function (paramIndex, decorator) {
    return function (target, key) { decorator(target, key, paramIndex); }
};


var AlvTableTemplateHeaderComponent = (function () {
    function AlvTableTemplateHeaderComponent(parent) {
        this.parent = parent;
        this.startCol = -1;
        this.endCol = -1;
        this.textAlign = "";
        this.width = 80;
        this.colspan = 1;
        this.index = -1;
        //@Input() order:number=-1;
        this.dataField = "";
        this.sortable = false;
        this.sortedBy = ""; // "" - not sorted, "asc" - ascending, "desc" - descending
        parent.tblColumns.push(this);
    }
    AlvTableTemplateHeaderComponent.prototype.ngOnInit = function () {
        this.order = this.index;
    };
    AlvTableTemplateHeaderComponent.prototype.ngClass = function () {
        var isCenter = false;
        var isLeft = false;
        var isRight = false;
        var just = this.textAlign;
        if (just) {
            isCenter = (just == 'center');
            isLeft = (just == 'left');
            isRight = (just == 'right');
        }
        else {
            var colDef = this.parent.options.colDefs[this.index];
            if (colDef) {
                just = colDef.justify;
                if (just) {
                    isCenter = (just == 'center');
                    isLeft = (just == 'left');
                    isRight = (just == 'right');
                }
            }
        }
        /*return {
          'text-center':isCenter,
          'text-left':isLeft,
          'text-right':isRight,
        }*/
        if (this.textAlign)
            console.log('this.textAlign:' + this.textAlign);
        return {
            'text-left': this.textAlign == 'left',
            'text-center': this.textAlign == 'center',
            'text-right': this.textAlign == 'right',
        };
    };
    AlvTableTemplateHeaderComponent.prototype.colWidth = function (idx) {
        if (this.startCol != -1 && this.endCol != -1) {
            this.colspan = this.endCol - this.startCol + 1;
            var idx_1;
            var ret = 0;
            for (idx_1 = this.startCol; idx_1 <= this.endCol; idx_1++) {
                ret += this.parent.options.colDefs[idx_1].width;
            }
            return ret;
        }
        else {
            var colDef = this.parent.options.colDefs[idx];
            if (colDef) {
                return this.parent.options.colDefs[idx].width;
            }
            else {
                return -1;
            }
        }
    };
    AlvTableTemplateHeaderComponent.prototype.headerClicked = function (e) {
        if (!this.sortable)
            return;
        var data = this.parent.data;
        var desc = (this.sortedBy == "asc");
        var key = this.dataField;
        if (desc) {
            this.sortedBy = "desc";
            data.sort(function (a, b) {
                return (a[key] < b[key]) ? 1 : ((b[key] < a[key]) ? -1 : 0);
            });
        }
        else {
            this.sortedBy = "asc";
            data.sort(function (a, b) {
                return (a[key] > b[key]) ? 1 : ((b[key] > a[key]) ? -1 : 0);
            });
        }
    };
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Number)
    ], AlvTableTemplateHeaderComponent.prototype, "startCol", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Number)
    ], AlvTableTemplateHeaderComponent.prototype, "endCol", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", String)
    ], AlvTableTemplateHeaderComponent.prototype, "textAlign", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Number)
    ], AlvTableTemplateHeaderComponent.prototype, "width", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Number)
    ], AlvTableTemplateHeaderComponent.prototype, "colspan", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Number)
    ], AlvTableTemplateHeaderComponent.prototype, "index", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", String)
    ], AlvTableTemplateHeaderComponent.prototype, "dataField", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Boolean)
    ], AlvTableTemplateHeaderComponent.prototype, "sortable", void 0);
    AlvTableTemplateHeaderComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-alv-table-template-header',
            template: __webpack_require__("../../../../../src/app/components/alv-table-template-header/alv-table-template-header.component.html"),
            styles: [__webpack_require__("../../../../../src/app/components/alv-table-template-header/alv-table-template-header.component.css")]
        }),
        __param(0, Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Inject"])(__WEBPACK_IMPORTED_MODULE_1__alv_table_template_alv_table_template_component__["a" /* AlvTableTemplateComponent */])),
        __metadata("design:paramtypes", [typeof (_a = typeof __WEBPACK_IMPORTED_MODULE_1__alv_table_template_alv_table_template_component__["a" /* AlvTableTemplateComponent */] !== "undefined" && __WEBPACK_IMPORTED_MODULE_1__alv_table_template_alv_table_template_component__["a" /* AlvTableTemplateComponent */]) === "function" && _a || Object])
    ], AlvTableTemplateHeaderComponent);
    return AlvTableTemplateHeaderComponent;
    var _a;
}());

//# sourceMappingURL=alv-table-template-header.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/alv-table-template/alv-table-template.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/alv-table-template/alv-table-template.component.html":
/***/ (function(module, exports) {

module.exports = "<div class=\"container-fluid\">\n<div class=\"card\">\n  <div class=\"card-header\">\n    Header\n  </div>\n  <div class=\"card-block\">\n\n    <table cellspacing=\"0\" cellpadding=\"0\" [attr.id]=\"tblId\" \n      class=\"table table-sm table-bordered hover\"\n      [style.width.px]=\"tableWidth()\">\n        <thead  [attr.id]=\"tblId + '_head'\"  style=\"display: block;overflow-x:hidden;\"\n          [style.width.px]=\"tableWidth()\">\n          <ng-content select=\"app-alv-table-template-header-row\"></ng-content>    \n        </thead>\n    \n        <tbody [attr.id]=\"tblId + '_body'\" style=\"display:block; overflow:auto;position:relative;\"\n            [style.width.px]=\"tableWidth()\"\n            [style.height.px]=\"tableHeight()\"\n            (scroll)=\"tblScroll($event)\"\n            >\n            <div [attr.id]=\"tblId + '_rows'\">\n              <!--tr *ngFor=\"let row of data.slice(options.startRow,options.endRow)\"-->\n                <ng-content select=\"[tablebody]\"></ng-content>\n              <!--/tr-->\n            </div>\n        </tbody>\n    </table>\n\n  </div>\n  <div class=\"card-footer\">\n    <div class=\"d-flex justify-content-between\">\n    <button class=\"btn btn-primary\" (click)=\"testClick($event)\">\n      Test\n    </button>\n\n    <input class=\"form-control\" type=\"text\" [(ngModel)]=\"options.startRow\" name=\"startRow\" (change)=\"calcRowIndex()\"/>\n    <input class=\"form-control\" type=\"text\" [(ngModel)]=\"options.endRow\" name=\"endRow\"/>\n\n    \n    <input class=\"form-control\" type=\"text\" [(ngModel)]=\"width\" name=\"width\"/>\n    <input class=\"form-control\" type=\"text\" [(ngModel)]=\"height\" name=\"height\"/>\n    \n    <input class=\"form-control\" type=\"text\" [(ngModel)]=\"tblRowHeight\" name=\"tblRowHeight\"/>\n    <input class=\"form-control\" type=\"text\" [(ngModel)]=\"tblScrollTop\" name=\"tblScrollTop\"/>\n\n    <input class=\"form-control\" type=\"text\" [(ngModel)]=\"tblRecOff\" name=\"tblRecOff\"/>\n    <input class=\"form-control\" type=\"text\" [(ngModel)]=\"tblRecOff2\" name=\"tblRecOff2\"/>\n\n    \n  </div>\n  </div>\n</div>\n</div>"

/***/ }),

/***/ "../../../../../src/app/components/alv-table-template/alv-table-template.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return AlvTableTemplateComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1_jquery__ = __webpack_require__("../../../../jquery/dist/jquery.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1_jquery___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_1_jquery__);
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};


var AlvTableTemplateComponent = (function () {
    function AlvTableTemplateComponent() {
        this.options = {};
        this.data = {};
        this.width = 600;
        this.height = 300;
        //@Input() rowWidth:number = 0;
        this.tableDefs = [];
        this.tblGuid = this.uuidv4();
        this.tblId = "alv_tbl_" + this.tblGuid;
        this.tblTheadId = this.tblId + '_head';
        this.tblTbodyId = this.tblId + '_body';
        this.tblRowsId = this.tblId + '_rows';
        this.tblColumns = [];
        this.scrollTimeout = 0;
        this.tblRowHeight = 0;
        this.tblMaxRecords = 0;
        this.tblTotalRecords = 0;
        this.tblScrollTop = 0;
        this.tblRecOff = 0;
        this.tblRecOff2 = 0;
    }
    AlvTableTemplateComponent.prototype.ngOnInit = function () {
        var _this = this;
        if (this.options.tableDefs) {
            this.tableDefs = this.options.tableDefs;
        }
        else {
            this.tableDefs = [];
        }
        this.options.startRow = 0;
        this.options.endRow = 1;
        setTimeout(function () { _this.setupTable(); }, 10);
    };
    AlvTableTemplateComponent.prototype.columnWidth = function (idx) {
        var dc = this.tblColumns[idx];
        //console.log('width:' +  dc.width);
        return dc.width;
    };
    AlvTableTemplateComponent.prototype.rowWidth = function () {
        var ret = 0;
        this.options.colDefs.forEach(function (e) {
            if (!e.hidden) {
                ret += e.width;
            }
        });
        return ret;
    };
    AlvTableTemplateComponent.prototype.tableWidth = function () {
        return this.width;
    };
    AlvTableTemplateComponent.prototype.tableHeight = function () {
        return this.height;
    };
    AlvTableTemplateComponent.prototype.tblScroll = function (e) {
        var _this = this;
        var cont = __WEBPACK_IMPORTED_MODULE_1_jquery__(e.target);
        var head = cont.parent().find('thead');
        head.scrollLeft(cont.scrollLeft());
        var tbody = __WEBPACK_IMPORTED_MODULE_1_jquery__('#' + this.tblTbodyId);
        this.tblScrollTop = tbody.scrollTop();
        //let offFullRec:number = Number(parseInt(String(this.tblScrollTop / this.tblRowHeight)));
        //this.options.startRow  = parseInt(String(this.tblScrollTop / this.tblRowHeight));
        //this.options.endRow  = Number(this.options.startRow) + Number(this.tblMaxRecords);    
        this.tblRecOff2 = this.tblScrollTop / this.tblRowHeight;
        var mt = this.tblRowHeight * parseInt(String(this.tblScrollTop / this.tblRowHeight));
        this.tblRecOff = (this.tblRecOff2 - parseInt(String(this.tblRecOff2))) * this.tblRowHeight;
        mt = this.tblScrollTop - this.tblRecOff;
        console.log();
        var trs = __WEBPACK_IMPORTED_MODULE_1_jquery__(tbody.find('tr'));
        //$(trs[0]).css('margin-top',
        //  Math.min(this.tblScrollTop,this.tblRowHeight * (this.tblTotalRecords - this.tblMaxRecords))  + 'px')
        if (this.scrollTimeout)
            clearTimeout(this.scrollTimeout);
        this.scrollTimeout = setTimeout(function () {
            var newRecStart = parseInt(String(_this.tblRecOff2));
            if (_this.options.startRow != newRecStart) {
                _this.options.startRow = newRecStart;
                _this.options.endRow = Math.min(newRecStart + _this.tblMaxRecords, _this.tblTotalRecords);
            }
            var rowsDiv = __WEBPACK_IMPORTED_MODULE_1_jquery__('#' + _this.tblRowsId);
            if (rowsDiv.length) {
                rowsDiv.css('padding-top', Math.min(_this.tblScrollTop - _this.tblRecOff, _this.tblRowHeight * (_this.tblTotalRecords - _this.tblMaxRecords + 1)) + 'px');
            }
        }, 5); // delay scroll updates by 5ms
    };
    AlvTableTemplateComponent.prototype.calcRowIndex = function () {
        console.log('calcRowIndex....');
        this.options.endRow = Number(this.options.startRow) + Number(this.tblMaxRecords);
    };
    AlvTableTemplateComponent.prototype.testClick = function (e) {
        this.options.colDefs.forEach(function (e) {
            e.order = e.index;
            //console.log(e);
        });
        this.setupTable();
    };
    AlvTableTemplateComponent.prototype.setupTable = function () {
        var tbl = __WEBPACK_IMPORTED_MODULE_1_jquery__('#' + this.tblId);
        var tbody = __WEBPACK_IMPORTED_MODULE_1_jquery__('#' + this.tblTbodyId);
        var thead = __WEBPACK_IMPORTED_MODULE_1_jquery__('#' + this.tblTheadId);
        var htrs = thead.find('tr');
        var trs = tbody.find('tr');
        this.tblRowHeight = trs.height();
        this.tblMaxRecords = parseInt(String(tbody.height() / this.tblRowHeight)) + 2;
        if (this.data.rows) {
            this.tblTotalRecords = this.data.rows.length;
        }
        else {
            this.tblTotalRecords = this.data.length;
        }
        this.options.startRow = 0;
        this.options.endRow = this.options.startRow + this.tblMaxRecords;
        var rowsDiv = tbl.find('#' + this.tblRowsId);
        if (rowsDiv.length) {
            rowsDiv.css('height', (this.tblTotalRecords * this.tblRowHeight) + 'px');
        }
        setTimeout(function () {
            //htrs = thead.find('tr');
            //trs = tbody.find('tr');
            //htrs.addClass('d-flex').css('position','relative').css('width',this.rowWidth + 'px');
            //trs.addClass('d-flex').css('width',this.rowWidth + 'px');
        }, 50);
    };
    AlvTableTemplateComponent.prototype.uuidv4 = function () {
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
            var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
            return v.toString(16);
        });
    };
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Object)
    ], AlvTableTemplateComponent.prototype, "options", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Object)
    ], AlvTableTemplateComponent.prototype, "data", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Number)
    ], AlvTableTemplateComponent.prototype, "width", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Number)
    ], AlvTableTemplateComponent.prototype, "height", void 0);
    AlvTableTemplateComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-alv-table-template',
            template: __webpack_require__("../../../../../src/app/components/alv-table-template/alv-table-template.component.html"),
            styles: [__webpack_require__("../../../../../src/app/components/alv-table-template/alv-table-template.component.css")]
        }),
        __metadata("design:paramtypes", [])
    ], AlvTableTemplateComponent);
    return AlvTableTemplateComponent;
}());

//# sourceMappingURL=alv-table-template.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/alv-table/alv-table.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/alv-table/alv-table.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return AlvTableComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__alv_table_panel_alv_table_panel_component__ = __webpack_require__("../../../../../src/app/components/alv-table-panel/alv-table-panel.component.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2_jquery__ = __webpack_require__("../../../../jquery/dist/jquery.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2_jquery___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_2_jquery__);
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};



var AlvTableComponent = (function () {
    //set pane(v: AlvTablePanelComponent) {
    //  setTimeout(() => { this.panelULv = v; 
    //    console.log('v...');
    //    console.log(v);
    //  }, 0);
    // }
    //@ViewChild('panelLL') panelLL:AlvTablePanelComponent
    function AlvTableComponent() {
        this.data = [
            { id: 1, name: "Item 1" },
            { id: 2, name: "Item 2" },
            { id: 3, name: "Item 3" },
            { id: 4, name: "Item 4" },
            { id: 5, name: "Item 5" },
        ];
        this.dataRows = [
            { id: 1, name: "Item 1" },
            { id: 2, name: "item 2", desc: "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book" },
            { id: 3, name: "Item 3", desc: "There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humour," },
            { id: 4, name: "Item 4" },
            { id: 5, name: "Item 5", desc: "There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humour," },
            { id: 6, name: "Item 6" },
            { id: 7, name: "Item 7" },
            { id: 8, name: "Item 8" },
            { id: 9, name: "Item 9" },
            { id: 10, name: "Item 10", desc: "The quick brown fox jumps over the lazy dog." },
            { id: 11, name: "Item 11" },
            { id: 12, name: "Item 12" },
            { id: 13, name: "Item 13" },
            { id: 14, name: "Item 14" },
            { id: 15, name: "Item 15" },
            { id: 16, name: "Item 16" },
            { id: 17, name: "Item 17" },
            { id: 18, name: "Item 18" },
            { id: 19, name: "Item 19" },
            { id: 20, name: "Item 20" },
            { id: 21, name: "Item 21" },
            { id: 22, name: "Item 22" },
            { id: 23, name: "Item 23" },
            { id: 24, name: "Item 24" },
            { id: 25, name: "Item 25" },
            { id: 26, name: "Item 26" },
            { id: 27, name: "Item 27" },
            { id: 28, name: "Item 28" },
            { id: 29, name: "Item 29" },
            { id: 30, name: "Item 30" },
        ];
        /*
          dataRows:Array<object>=[
            {id:1,name:"Item 1"},
            {id:2,name:"Lorem Ipsum is simply dummy text of "},
            {id:3,name:"There are many variations of passages"},
            {id:4,name:"Item 4"},
            {id:5,name:"There are many variations of passages,"},
            {id:6,name:"Item 6"},
            {id:7,name:"Item 7"},
            {id:8,name:"Item 8"},
            {id:9,name:"Item 9"},
            {id:10,name:"The quick brown fox jumps over the lazy dog."},
            {id:11,name:"Item 11"},
            {id:12,name:"There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humour,"},
            {id:13,name:"Item 13"},
            {id:14,name:"Item 14"},
            {id:15,name:"Item 15"},
            {id:16,name:"Item 16"},
            {id:17,name:"Item 17"},
            {id:18,name:"Item 18"},
            {id:19,name:"Item 19"},
            {id:20,name:"Item 20"},
            {id:21,name:"Item 21"},
            {id:22,name:"Item 22"},
            {id:23,name:"Item 23"},
            {id:24,name:"Item 24"},
            {id:25,name:"Item 25"},
            {id:26,name:"Item 26"},
            {id:27,name:"Item 27"},
            {id:28,name:"Item 28"},
            {id:29,name:"Item 29"},
            {id:30,name:"Item 30"},
          ]
        */
        this.dataRowsSubset = [];
        this.dispDataStartIndex = 0;
        this.dispDataRows = 8;
        this.dispDataTop = 0;
        this.dispWindowTop = 0;
        this.contScrollX = 0;
        this.contScrollY = 0;
        this.dispWindow = null;
        this.dispWindowHeight = 0;
        this.dispRow1Height = 0;
        this.dispRow2Height = 0;
        this.dispRow1Top = 0;
        this.dispRow2Top = 0;
        this.scrollTimout = null;
        this.suspendScroll = false;
        this.isScrollDown = true;
        this.isScrollUp = false;
        this.prvScroll = 0;
    }
    AlvTableComponent.prototype.ngOnInit = function () {
        var _this = this;
        setTimeout(function () { console.log(_this.panelULv); }, 100);
        this.shftDataUp();
    };
    AlvTableComponent.prototype.onScrollEvent = function (e) {
        if (!this.dispWindow)
            this.dispWindow = __WEBPACK_IMPORTED_MODULE_2_jquery__(e.srcElement);
        this.contScrollY = this.dispWindow.scrollTop();
        this.isScrollDown = (this.prvScroll < this.contScrollY);
        this.isScrollUp = (this.prvScroll > this.contScrollY);
        if (this.isScrollDown)
            this.onScrollDown();
        if (this.isScrollUp)
            this.onScrollUp();
        this.prvScroll = this.contScrollY;
    };
    AlvTableComponent.prototype.onScrollUp = function () {
        if (this.dispDataStartIndex <= 0) {
            this.dispWindow.scrollTop(0);
            this.dispWindowTop = 0;
            return;
        }
        if (this.dispWindowTop > this.contScrollY) {
            // the displayed window needs another row before the current first row
            // load the new set of data starting from d(-1)-with refrence 
            // to the current d(0).
            // to fill the gap, the new first row's top must be set to the current display
            // data window top less the possible height of the new first row.
            // to determine the height of the new first row, it has to be rendered first then
            // query its height.
            this.dispDataStartIndex--; // set the new data first row index
            var pscr = this.contScrollY;
            this.shftDataUp(); // get new set of data
            // get height of the new first row
            // reposition top of the data display window
            //let rows=this.dispWindow.find('tr');
            //let r1=$(rows[0]);
            //let r1h:number=r1.height();
            var r1h = this.getRow1Height();
            this.dispWindowTop -= r1h;
        }
    };
    AlvTableComponent.prototype.getHeight = function (row) {
        if (row._rowHeight_ == undefined) {
            var r1 = __WEBPACK_IMPORTED_MODULE_2_jquery__('#_t1_' + row.id);
            var r2 = __WEBPACK_IMPORTED_MODULE_2_jquery__('#_t2_' + row.id);
            var d1 = r1.find('div').outerHeight();
            var d2 = r2.find('div').outerHeight();
            if (d1 == undefined || d2 == undefined) {
                return 33;
            }
            else {
                /*        row._rowHeight_ = Math.max(
                          d1+parseInt(r1.css('padding-top'))+parseInt(r1.css('padding-bottom')),
                          d2+parseInt(r2.css('padding-top'))+parseInt(r2.css('padding-bottom'))
                        );*/
                row._rowHeight_ = Math.max(r1.outerHeight(), r2.outerHeight());
            }
        }
        return row._rowHeight_;
    };
    AlvTableComponent.prototype.onScrollDown = function () {
        // stay on the current first row until it is hidden...
        // that, is, while current scroll "contScrollY" less height of the first row 
        // is >= container top
        // if the current starting index is more than or equal to the length of 
        // source data less the number of rows displayed then do not perform the operation
        //if(this.dispDataStartIndex>=(this.dataRows.length-(this.dispDataRows+1))) {
        //  return;
        //}
        //let rows=this.dispWindow.find('tr');
        //let r1=$(rows[0]);
        //let r1h:number=r1.height();
        var r1h = this.getRow1Height();
        if ((this.contScrollY - r1h) >= this.dispWindowTop) {
            this.dispDataStartIndex++;
            // get next set of data
            this.shftDataUp();
            // reposition displayed data
            this.dispWindowTop = this.contScrollY;
        }
    };
    AlvTableComponent.prototype.getRow1Height = function () {
        var r = this.dataRowsSubset[0];
        if (r._rowHeight_) {
            return r._rowHeight_;
        }
        else {
            var rows = this.dispWindow.find('tbody').find('tr');
            var r1 = __WEBPACK_IMPORTED_MODULE_2_jquery__(rows[0]);
            var r1h = r1.height();
            r._rowHeight_ = r1h;
            return r1h;
        }
    };
    AlvTableComponent.prototype.onScrollEventX = function (e) {
        //if(this.suspendScroll)return;
        if (!this.dispWindow)
            this.dispWindow = __WEBPACK_IMPORTED_MODULE_2_jquery__(e.srcElement);
        if (this.scrollTimout)
            clearTimeout(this.scrollTimout);
        //    this.scrollTimout=setTimeout(()=>{
        //let tbl=$(e.srcElement);
        this.dispWindowHeight = this.dispWindow.height();
        var trs = this.dispWindow.find('tr');
        this.dispRow1 = __WEBPACK_IMPORTED_MODULE_2_jquery__(trs[0]);
        this.dispRow2 = __WEBPACK_IMPORTED_MODULE_2_jquery__(trs[1]);
        this.dispRow1Height = this.dispRow1.height() + 1;
        this.dispRow2Height = this.dispRow2.height() + 1;
        this.dispRow1Top = this.dispRow1.position().top;
        this.dispRow2Top = this.dispRow2.position().top;
        this.contScrollY = this.dispWindow.scrollTop();
        if (this.prvScroll < this.contScrollY) {
            this.isScrollDown = true;
        }
        else if (this.prvScroll > this.contScrollY) {
            this.isScrollUp = true;
        }
        else {
            this.isScrollDown = false;
            this.isScrollUp = false;
        }
        this.prvScroll = this.contScrollY;
        if (this.dispDataStartIndex > 0 && this.isScrollUp) {
            if (this.contScrollY < this.dispWindowTop) {
                // load d(-1)
                this.suspendScroll = true;
                var prvScr = this.contScrollY;
                var diff = this.dispWindowTop - this.contScrollY;
                this.dispDataStartIndex--;
                this.shftDataUp();
                trs = this.dispWindow.find('tr');
                this.dispRow1 = __WEBPACK_IMPORTED_MODULE_2_jquery__(trs[0]);
                this.dispRow1Height = this.dispRow1.height() + 1;
                this.dispRow2 = __WEBPACK_IMPORTED_MODULE_2_jquery__(trs[1]);
                this.dispRow2Height = this.dispRow2.height() + 1;
                //this.dispWindowTop=this.contScrollY-this.dispRow1Height+diff;
                this.dispWindowTop = prvScr - this.dispRow1Height + diff;
                this.suspendScroll = false;
            }
            if (this.dispDataStartIndex == 0)
                this.dispWindowTop = 0;
        }
        trs = this.dispWindow.find('tr');
        this.dispRow1 = __WEBPACK_IMPORTED_MODULE_2_jquery__(trs[0]);
        this.dispRow1Height = this.dispRow1.height();
        this.dispRow1 = __WEBPACK_IMPORTED_MODULE_2_jquery__(trs[1]);
        this.dispRow2Height = this.dispRow2.height();
        this.dispRow1Top = this.dispRow1.position().top;
        this.dispRow2Top = this.dispRow2.position().top;
        if (this.dispDataStartIndex < (this.dataRows.length - this.dispDataRows) &&
            this.isScrollDown &&
            !this.suspendScroll) {
            if ((this.contScrollY - this.dispRow1Height) > this.dispWindowTop) {
                //if(this.contScrollY == this.dispRow2Top){
                this.dispWindowTop = this.contScrollY;
                this.dispDataStartIndex++;
                this.shftDataUp();
            }
        }
        //    },1);
        // dispWindowTop
    };
    AlvTableComponent.prototype.shftDataUp = function (force) {
        if (force == undefined)
            force = false;
        force = true;
        if (this.dataRowsSubset.length == 0 || force) {
            var endIndex = this.dispDataStartIndex + this.dispDataRows;
            this.dataRowsSubset = this.dataRows.slice(this.dispDataStartIndex, endIndex);
        }
        else {
        }
    };
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["ViewChild"])(__WEBPACK_IMPORTED_MODULE_1__alv_table_panel_alv_table_panel_component__["a" /* AlvTablePanelComponent */]),
        __metadata("design:type", typeof (_a = typeof __WEBPACK_IMPORTED_MODULE_1__alv_table_panel_alv_table_panel_component__["a" /* AlvTablePanelComponent */] !== "undefined" && __WEBPACK_IMPORTED_MODULE_1__alv_table_panel_alv_table_panel_component__["a" /* AlvTablePanelComponent */]) === "function" && _a || Object)
    ], AlvTableComponent.prototype, "panelULv", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", typeof (_b = typeof __WEBPACK_IMPORTED_MODULE_0__angular_core__["TemplateRef"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_0__angular_core__["TemplateRef"]) === "function" && _b || Object)
    ], AlvTableComponent.prototype, "headerTemplate", void 0);
    AlvTableComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-alv-table',
            // templateUrl: './alv-table.component.html',
            template: __webpack_require__("../../../../../src/app/components/alv-table/alv-table.component2.html"),
            styles: [__webpack_require__("../../../../../src/app/components/alv-table/alv-table.component.css")]
        }),
        __metadata("design:paramtypes", [])
    ], AlvTableComponent);
    return AlvTableComponent;
    var _a, _b;
}());

//# sourceMappingURL=alv-table.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/alv-table/alv-table.component2.html":
/***/ (function(module, exports) {

module.exports = "<!--ng-container *ngTemplateOutlet=\"headerTemplate\">\r\n\r\n</ng-container-->\r\n<app-alv-table-panel *ngTemplateOutlet=\"headerTemplate\"></app-alv-table-panel>\r\n<!--app-alv-table-panel #panelLL [showLL]=\"true\" *ngTemplateOutlet=\"headerTemplate\"></app-alv-table-panel-->\r\n"

/***/ }),

/***/ "../../../../../src/app/components/alv-test-component/alv-test-component.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/alv-test-component/alv-test-component.component.html":
/***/ (function(module, exports) {

module.exports = "<span>Test</span>"

/***/ }),

/***/ "../../../../../src/app/components/alv-test-component/alv-test-component.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return AlvTestComponentComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};

var AlvTestComponentComponent = (function () {
    function AlvTestComponentComponent() {
    }
    AlvTestComponentComponent.prototype.ngOnInit = function () {
    };
    AlvTestComponentComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-alv-test-component',
            template: __webpack_require__("../../../../../src/app/components/alv-test-component/alv-test-component.component.html"),
            styles: [__webpack_require__("../../../../../src/app/components/alv-test-component/alv-test-component.component.css")]
        }),
        __metadata("design:paramtypes", [])
    ], AlvTestComponentComponent);
    return AlvTestComponentComponent;
}());

//# sourceMappingURL=alv-test-component.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/ctl-drop-down/ctl-drop-down.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "li.item{\r\n    cursor: pointer;\r\n    white-space:nowrap;\r\n    -webkit-user-select: none;\r\n       -moz-user-select: none;\r\n        -ms-user-select: none;\r\n            user-select: none;\r\n}\r\nli.item:hover, li.itemCurrent.hover{\r\n    background: #296192;\r\n    color: white;\r\n}\r\n\r\nli.itemCurrent{\r\n    background: #051D31;\r\n    color: yellow;\r\n}\r\n\r\n::-webkit-scrollbar {\r\n    width: 15px;\r\n}\r\n\r\n::-webkit-scrollbar-track {\r\n    background: white;\r\n    -webkit-box-shadow: inset 0 0 6px rgba(0,0,0,0.3); \r\n    border-radius: 2px;\r\n}\r\n\r\n::-webkit-scrollbar-thumb {\r\n    border-radius: 2px;\r\n    -webkit-box-shadow: inset 0 0 6px rgba(0,0,0,0.5); \r\n}", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/ctl-drop-down/ctl-drop-down.component.html":
/***/ (function(module, exports) {

module.exports = "\n<!--input type=\"text\" placeholder=\"Search...\" style=\"width:100%\" class=\"form-control form-control-sm\" id=\"firstName\"  \n[(ngModel)]=\"drpdFilter\" \n-->     \n<div style=\"overflow-y:auto;border:1px solid #c0c0c0;\"\n  [style.width]=\"drpdWidth + 'px'\"\n  [style.height]=\"drpdHeight + 'px'\"\n  (scroll)=\"onScroll($event)\">\n\n  <div id=\"drpdScrollDiv\" style=\"position:relative;width:100%;\"\n      [style.height]=\"drpdScrollHeight+'px'\">\n\n      <div id=\"drpdInformationDiv\" style=\"padding:0px;width:100%;position:absolute;left:0px;overflow:hidden;\"\n          [style.top]=\"(drpdStartRow * drpdRowHeight) +'px'\"\n          [style.height]=\"drpdScrollContainerHeight() +'px'\"\n      >\n    \n          <ul class=\"list-group\">\n              <li *ngFor=\"let item of drpdDispData;let idx=index;\" \n                  class=\"list-group-item px-1 py-0\"\n                  style=\"white-space:nowrap;\"\n                  [ngClass]=\"{'itemCurrent':item.id==drpdValue, 'item':item.id!=drpdValue}\"\n                  [style.height]=\"drpdRowHeight+'px'\"\n                  [style.font-size]=\"drpdFontSize\"\n                  [title]=\"item.name\"\n                  (click)=\"itemClick(item)\"\n                  (mousedown)=\"itemMouseDown(item)\"\n                  >\n                  {{item.name}}\n              </li>\n          </ul>\n\n      </div>\n\n  </div>\n</div>"

/***/ }),

/***/ "../../../../../src/app/components/ctl-drop-down/ctl-drop-down.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return CtlDropDownComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1_jquery__ = __webpack_require__("../../../../jquery/dist/jquery.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1_jquery___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_1_jquery__);
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};


var CtlDropDownComponent = (function () {
    function CtlDropDownComponent() {
        this.itemId = 0;
        this.itemSelected = new __WEBPACK_IMPORTED_MODULE_0__angular_core__["EventEmitter"]();
        this.drpdData = [];
        this.drpdWidth = 200;
        this.drpdStartRow = 0;
        this.drpdRowHeight = 25;
        this.drpdDispRows = 20;
        this.drpdFontSize = '0.8em';
        this.drpdFilter = "";
        this.drpdRecCount = this.drpdData.length;
        this.drpdHeight = this.drpdDispRows * this.drpdRowHeight;
        this.drpdScrollHeight = Math.max(this.drpdHeight, this.drpdRecCount * this.drpdRowHeight - 1);
        this.drpdDispData = [];
        this.drpdValueChange = new __WEBPACK_IMPORTED_MODULE_0__angular_core__["EventEmitter"]();
    }
    Object.defineProperty(CtlDropDownComponent.prototype, "drpdValue", {
        get: function () {
            return this.itemId;
        },
        set: function (val) {
            this.itemId = val;
            //console.log('value:'+val);
            this.drpdValueChange.emit(this.itemId);
        },
        enumerable: true,
        configurable: true
    });
    CtlDropDownComponent.prototype.ngOnChanges = function () {
        // alert('changed!!!')
        // if bound data is changed coded in this hook will be executed
        //if(this.drpdValue)
        this.setInitialIndex();
    };
    CtlDropDownComponent.prototype.setInitialIndex = function () {
        return;
        /*let idx=this.drpdData.findIndex((item:any)=>{ return item.id==this.drpdValue});
        if(idx==-1){
          this.drpdStartRow = 0;
        }else{
          this.drpdStartRow = idx;
        }
        this.onScroll(null);*/
    };
    CtlDropDownComponent.prototype.ngOnInit = function () {
        /*
            this.drpdRecCount=this.drpdData.length;
            this.drpdHeight=this.drpdDispRows * this.drpdRowHeight;
            this.drpdScrollHeight = Math.max(this.drpdHeight, this.drpdRecCount * this.drpdRowHeight-1);
            this.drpdDispData = this.drpdData.slice(this.drpdStartRow,this.drpdStartRow+this.drpdDispRows+1);
        */
        var _this = this;
        setTimeout(function () {
            _this.drpdRecCount = _this.drpdData.length;
            _this.drpdHeight = _this.drpdDispRows * _this.drpdRowHeight;
            _this.drpdScrollHeight = Math.max(_this.drpdHeight, _this.drpdRecCount * _this.drpdRowHeight - 1);
            _this.drpdDispData = _this.drpdData.slice(_this.drpdStartRow, _this.drpdStartRow + _this.drpdDispRows + 1);
            _this.setInitialIndex();
        }, 500);
    };
    CtlDropDownComponent.prototype.drpdScrollContainerHeight = function () {
        return ((Math.min(this.drpdDispRows, this.drpdData.length)) * this.drpdRowHeight);
    };
    CtlDropDownComponent.prototype.onScroll = function (evt) {
        var det = __WEBPACK_IMPORTED_MODULE_1_jquery__('#drpdInformationDiv');
        var oldTop = det.position().top;
        var newTop = evt.target.scrollTop;
        if (newTop > (oldTop + this.drpdRowHeight) || ((oldTop - newTop + this.drpdRowHeight) >= this.drpdRowHeight)) {
            det.css('top', (newTop) + 'px');
            this.drpdStartRow = parseInt(String(newTop / this.drpdRowHeight));
            this.drpdDispData = this.drpdData.slice(this.drpdStartRow, this.drpdStartRow + this.drpdDispRows + 1);
        }
    };
    CtlDropDownComponent.prototype.itemClick = function (item) {
        //this.drpdValue=item.id;
        //this.itemValue.emit(item.id);
        //console.log(item);
        //this.itemSelected.emit(item.id);
        //alert('itemClick: '+ item.id);
    };
    CtlDropDownComponent.prototype.itemMouseDown = function (item) {
        this.drpdValue = item.id;
        //alert('itemMouseDown: '+ item.id);
    };
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Object),
        __metadata("design:paramtypes", [Object])
    ], CtlDropDownComponent.prototype, "drpdValue", null);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Output"])(),
        __metadata("design:type", typeof (_a = typeof __WEBPACK_IMPORTED_MODULE_0__angular_core__["EventEmitter"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_0__angular_core__["EventEmitter"]) === "function" && _a || Object)
    ], CtlDropDownComponent.prototype, "drpdValueChange", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Output"])(),
        __metadata("design:type", Object)
    ], CtlDropDownComponent.prototype, "itemSelected", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Object)
    ], CtlDropDownComponent.prototype, "drpdData", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Number)
    ], CtlDropDownComponent.prototype, "drpdWidth", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Number)
    ], CtlDropDownComponent.prototype, "drpdStartRow", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Number)
    ], CtlDropDownComponent.prototype, "drpdRowHeight", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Number)
    ], CtlDropDownComponent.prototype, "drpdDispRows", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", String)
    ], CtlDropDownComponent.prototype, "drpdFontSize", void 0);
    CtlDropDownComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-ctl-drop-down',
            template: __webpack_require__("../../../../../src/app/components/ctl-drop-down/ctl-drop-down.component.html"),
            styles: [__webpack_require__("../../../../../src/app/components/ctl-drop-down/ctl-drop-down.component.css")]
        }),
        __metadata("design:paramtypes", [])
    ], CtlDropDownComponent);
    return CtlDropDownComponent;
    var _a;
}());

//# sourceMappingURL=ctl-drop-down.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/ctl-user-select/ctl-user-select.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "li.item{\r\n    cursor: pointer;\r\n}\r\nli.item:hover{\r\n    background: #296192;\r\n    color: white;\r\n}\r\n\r\n::-webkit-scrollbar {\r\n    width: 12px;\r\n}\r\n\r\n::-webkit-scrollbar-track {\r\n    -webkit-box-shadow: inset 0 0 6px rgba(0,0,0,0.3); \r\n    border-radius: 10px;\r\n}\r\n\r\n::-webkit-scrollbar-thumb {\r\n    border-radius: 10px;\r\n    -webkit-box-shadow: inset 0 0 6px rgba(0,0,0,0.5); \r\n}", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/ctl-user-select/ctl-user-select.component.html":
/***/ (function(module, exports) {

module.exports = "<!-- Split dropup button -->\n<div class=\"btn-group dropdown\">\n  <input type=\"text\" \n    [style.width]=\"getTextWidth()\"\n    class=\"form-control form-control-sm\"\n    [value]=\"lookup['r'+counter][displayField]\"\n    onfocus=\"this.blur();\"\n    [disabled]=\"disableSelect\"\n    placeholder=\"select a person...\"\n    />\n  <button type=\"button\" \n    [disabled]=\"disableSelect\"\n    class=\"btn btn-secondary btn-sm\" data-toggle=\"dropdown\" \n    aria-haspopup=\"true\" aria-expanded=\"false\">\n    <!--span class=\"\">Toggle Dropdown</span-->\n    <i class=\"fa fa-user\"></i>\n  </button>\n  <div class=\"dropdown-menu\" [style.width]=\"getDropdownWidth()\"  \n    >\n    <div class=\"card\" [style.height]=\"(height)+'px'\">\n      <div class=\"card-header p-1 m-0\">\n\n          <div class=\"input-group\">\n            <input type=\"text\" style=\"font-size:0.8em;\"\n              class=\"form-control form-control-sm\" \n              [(ngModel)]=\"searchText\"\n              placeholder=\"Search for...\">\n            <span class=\"input-group-btn\">\n              <button class=\"btn btn-secondary btn-sm\" type=\"button\">\n                  <i class=\"fa fa-search\" aria-hidden=\"true\"></i>\n              </button>\n            </span>\n          </div>        \n\n      </div>\n      <div class=\"card-block p-0 m-0\" \n        style=\"overflow-y:auto\">\n          <ul class=\"list-group\">\n              <div *ngFor=\"let group of source;trackBy:trackByGroup;\" >\n              <li class=\"list-group-item p-1 m-0\" \n                *ngIf=\"itemSource(group).length>0\"\n                style=\"background:#eceeef;\">\n                <div style=\"cursor:default;\"\n                  [style.font-size]=\"fontSize\"\n                >\n                  {{ group[groupName]+':' }}\n                </div>\n                <ul class=\"list-group\">\n                  \n                  <li class=\"item list-group-item p-1\" *ngFor=\"let user of itemSource(group);trackBy:trackByUser;\" \n                    (mousedown)=\"itemClicked(user[valueColumn],user,group)\"\n                    style=\"cursor:pointer;\"\n                    [style.font-size]=\"fontSize\"\n                    >\n                    {{ itemDisplayText(user) }}\n                  </li>\n                  <!--li class=\"item list-group-item p-1\"\n                    (click)=\"itemClickTest2()\"\n                    (mousedown)=\"itemClickTest2()\">\n                    \n                    test prompt\n                  </li-->\n                </ul>\n              </li>\n              </div>\n            </ul>\n      </div>\n    </div>\n\n  </div>\n</div>  "

/***/ }),

/***/ "../../../../../src/app/components/ctl-user-select/ctl-user-select.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return CtlUserSelectComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};

var CtlUserSelectComponent = (function () {
    function CtlUserSelectComponent() {
        this.counterValue = 0;
        this.searchText = '';
        this.counterChange = new __WEBPACK_IMPORTED_MODULE_0__angular_core__["EventEmitter"]();
        this.itemSelected = new __WEBPACK_IMPORTED_MODULE_0__angular_core__["EventEmitter"]();
        this.source = [];
        this.displayColumn = '';
        this.fontSize = '0.8em';
    }
    Object.defineProperty(CtlUserSelectComponent.prototype, "counter", {
        get: function () {
            return this.counterValue;
        },
        set: function (val) {
            this.counterValue = val;
            this.counterChange.emit(this.counterValue);
        },
        enumerable: true,
        configurable: true
    });
    CtlUserSelectComponent.prototype.ngOnInit = function () {
        //alert(this.lookup['r'+this.value]['fullName']);
        //lert(this.source[0]['data'].length);
        if (this.widthUnit == undefined)
            this.widthUnit = 'px';
    };
    CtlUserSelectComponent.prototype.itemClickedTest = function (user) {
        alert('item is clicked');
    };
    CtlUserSelectComponent.prototype.itemClickTest2 = function () {
        alert('item is clicked2');
    };
    //(click)="itemClicked(user[valueColumn],user,group)" 
    CtlUserSelectComponent.prototype.itemClicked = function (userID, user, team) {
        //this.counter=user['id'];
        this.counter = userID;
        this.itemSelected.emit({ "user": user, "team": team });
    };
    CtlUserSelectComponent.prototype.xitemDisplayText = function (user) {
        return user[this.valueColumn];
    };
    CtlUserSelectComponent.prototype.itemDisplayText = function (user) {
        var retVal = '';
        //user[valueColumn]
        if (this.displayColumn == undefined)
            this.displayColumn = '';
        if (this.displayColumn.length == 0) {
            // if display column is not specified, use the lookup object as source of display,
            // using the user[valueColumn] to search for the user item
            retVal = this.lookup['r' + user[this.valueColumn]][this.displayField];
        }
        else {
            retVal = user[this.valueColumn];
        }
        return retVal;
    };
    CtlUserSelectComponent.prototype.itemSource = function (group) {
        var _this = this;
        var retVal = group[this.dataName];
        if (group[this.dataName].length > 0) {
        }
        var newArr = retVal.filter(function (item) {
            return (_this.itemDisplayText(item).toLowerCase().indexOf(_this.searchText.toLowerCase()) != -1);
        });
        return newArr;
    };
    CtlUserSelectComponent.prototype.getTextWidth = function () {
        return (this.width - 30) + this.widthUnit;
    };
    CtlUserSelectComponent.prototype.getDropdownWidth = function () {
        return (this.width) + this.widthUnit;
    };
    CtlUserSelectComponent.prototype.trackByUser = function (idx, item) {
        return idx;
    };
    CtlUserSelectComponent.prototype.trackByGroup = function (idx, item) {
        return idx;
    };
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Object),
        __metadata("design:paramtypes", [Object])
    ], CtlUserSelectComponent.prototype, "counter", null);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Output"])(),
        __metadata("design:type", Object)
    ], CtlUserSelectComponent.prototype, "counterChange", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Output"])(),
        __metadata("design:type", Object)
    ], CtlUserSelectComponent.prototype, "itemSelected", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Number)
    ], CtlUserSelectComponent.prototype, "width", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", String)
    ], CtlUserSelectComponent.prototype, "widthUnit", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Number)
    ], CtlUserSelectComponent.prototype, "height", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Number)
    ], CtlUserSelectComponent.prototype, "value", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Boolean)
    ], CtlUserSelectComponent.prototype, "disableSelect", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Object)
    ], CtlUserSelectComponent.prototype, "lookup", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", String)
    ], CtlUserSelectComponent.prototype, "displayField", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", String)
    ], CtlUserSelectComponent.prototype, "valueField", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Object)
    ], CtlUserSelectComponent.prototype, "source", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", String)
    ], CtlUserSelectComponent.prototype, "groupName", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", String)
    ], CtlUserSelectComponent.prototype, "dataName", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", String)
    ], CtlUserSelectComponent.prototype, "valueColumn", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", String)
    ], CtlUserSelectComponent.prototype, "displayColumn", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", String)
    ], CtlUserSelectComponent.prototype, "fontSize", void 0);
    CtlUserSelectComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-ctl-user-select',
            template: __webpack_require__("../../../../../src/app/components/ctl-user-select/ctl-user-select.component.html"),
            styles: [__webpack_require__("../../../../../src/app/components/ctl-user-select/ctl-user-select.component.css")]
        }),
        __metadata("design:paramtypes", [])
    ], CtlUserSelectComponent);
    return CtlUserSelectComponent;
}());

//# sourceMappingURL=ctl-user-select.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/frm-activity/frm-activity.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/frm-activity/frm-activity.component.html":
/***/ (function(module, exports) {

module.exports = "<div class=\"card\"  id=\"actCard\"  style=\"height:100%;\" (window:resize)=\"onResize($event)\">\n    <div class=\"card-header row py-0\">\n        <div class=\"col float-left py-0\">\n            <h4 class=\"card-title\"><i style=\"font-size:2em;margin-top:5px;margin-right:20px;\" \n              class=\"fa fa-list-alt text-muted\" aria-hidden=\"true\"></i>Activity Details \n              <span *ngIf=\"withConflictInTeam(0,'W',true)['withWarning']\">\n                <i class=\"fa fa-flag text-info\" \n                  title=\"With assignment continuation ...\"\n                  style=\"font-size:0.7em;cursor:pointer;\" \n                  aria-hidden=\"true\">\n                </i>\n              </span>\n              <span *ngIf=\"withConflictInTeam(0,'C',true)['withConflict']\">\n                <i class=\"fa fa-flag text-danger\" \n                  title=\"Assignment schedule conflict(s) exist ...\"\n                  style=\"font-size:0.7em;cursor:pointer;\" \n                  aria-hidden=\"true\">\n                </i>\n              </span>\n            </h4>\n        </div>\n        <div class=\" btn-group-sm col m-0 float-right text-right align-text-bottom text-info\"\n          style=\"padding-top:20px;\"\n          \n          >\n\n<!-- Button trigger modal -->\n<!--button type=\"button\" class=\"btn btn-primary\" data-toggle=\"modal\" data-target=\"#exampleModal\">\n    Launch demo modal\n  </button-->          \n\n  <!--button class=\"btn btn-link\">\n    <i class=\"fa fa-spinner fa-spin fa-3x fa-fw text-info\" style=\"font-size:20px;\"></i>\n    <i class=\"fa fa-circle-o-notch fa-spin fa-3x fa-fw text-info\" style=\"font-size:20px;\"></i>\n    <i class=\"fa fa-refresh fa-spin fa-3x fa-fw text-info\" style=\"font-size:20px;\"></i>\n    <i class=\"fa fa-cog fa-spin fa-3x fa-fw text-info\" style=\"font-size:20px;\"></i>\n    <i class=\"fa fa-spinner fa-pulse fa-3x fa-fw text-info\" style=\"font-size:20px;\"></i>\n  </button>\n\n  <span class=\"fa-stack fa-lg\">\n    <i class=\"fa fa-camera fa-stack-1x\"></i>\n    <i class=\"fa fa-ban fa-stack-2x text-danger\"></i>\n  </span-->\n  \n\n  <button *ngIf=\"oplandb.logInfo.activityAdd\"\n  style=\"position: relative;\"\n  class=\"btn btn-link\" \n  [ngClass]=\"udf.acnBtnClass('add')\"\n  (click)=\"reloadActivity()\"\n  onfocus=\"this.blur()\"\n  title=\"Reload activity record ...\"\n>\n  <i id=\"act_reload\" class=\"fa fa-refresh text-info\" style=\"font-size:20px;cursor:pointer;\"></i>\n</button>\n          <button *ngIf=\"oplandb.logInfo.activityAdd\"\n            style=\"position: relative;\"\n            class=\"btn btn-link\" \n            [ngClass]=\"udf.acnBtnClass('add')\"\n            (click)=\"toggleShowInCalendar()\"\n            onfocus=\"this.blur()\"\n            title=\"Toggle visibility in calendar ...\"\n          >\n            <i id=\"per_add\" [ngClass]=\"inCalendar()\" class=\"fa\" \n              \n              style=\"font-size:23px;cursor:pointer;\"></i>\n\n          <!--i *ngIf=\"noAccess()\" class=\"fa fa-ban text-danger\" style=\"z-index:1;position:absolute;top:12px;left:25px;font-size:12px\"></i-->\n          <i *ngIf=\"noAccess()\" class=\"fa fa-ban fa-stack-2x text-danger\" style=\"top:-1px;\"></i>\n              \n          </button>\n\n\n          <button *ngIf=\"oplandb.logInfo.activityMemAdmin\" class=\"btn btn-sm btn-link\" \n          style=\"position: relative;\"\n          data-toggle=\"modal\"\n          data-target=\"#exampleModal\"\n          (click)=\"btnClick('add')\"\n          title=\"Add new activity member ...\"\n          [disabled]=\"alertModal() || !disableControls('header')\"\n        >\n          <i class=\"fa fa-user-plus text-info\"  [ngClass]=\"{'text-muted':noAccess(),'text-info':!noAccess()}\" style=\"font-size:20px;cursor:pointer;\"></i>\n          <!--i *ngIf=\"noAccess()\" class=\"fa fa-ban text-danger\" style=\"z-index:1;position:absolute;top:12px;left:20px;font-size:12px\"></i-->\n          <i *ngIf=\"noAccess()\" class=\"fa fa-ban fa-stack-2x text-danger\" style=\"top:-1px;\"></i>\n        \n        </button>  \n        \n        \n\n          <button *ngIf=\"oplandb.logInfo.activityAdd\"\n            class=\"btn btn-link\" \n            [ngClass]=\"udf.acnBtnClass('add')\"\n            (click)=\"btnClick('ADD')\"\n            onfocus=\"this.blur()\"\n            title=\"Add new activity...\"\n            [disabled]=\"alertModal() || !disableControls('header') || !disableControls()\"\n            >\n            <i id=\"per_add\" class=\"fa fa-plus-square-o text-info\" style=\"font-size:20px;cursor:pointer;\"></i>\n          </button>\n\n          <button  *ngIf=\"oplandb.logInfo.activityEdit && (oplandb.currActivity.id!=-1 && (oplandb.datActivities.length>0))\"\n            class=\"btn btn-link\" \n            style=\"position: relative;\"\n            [ngClass]=\"udf.acnBtnClass('edit')\"\n            (click)=\"btnClick('EDIT')\"\n            onfocus=\"this.blur()\"\n            title=\"Edit activity details...\"\n            [disabled]=\"alertModal() || !disableControls('header') || !disableControls()\"\n            >\n            <i id=\"per_edit\" class=\"fa fa-pencil-square-o\"  [ngClass]=\"{'text-muted':noAccess(),'text-info':!noAccess()}\"  style=\"font-size:20px;cursor:pointer;\"></i>\n            <!--i *ngIf=\"noAccess()\" class=\"fa fa-ban text-danger\" style=\"z-index:1;position:absolute;top:12px;left:20px;font-size:12px\"></i-->\n            <i *ngIf=\"noAccess()\" class=\"fa fa-ban fa-stack-2x text-danger\" style=\"top:-1px;left:-1px;\"></i>\n\n          </button>\n          <button  *ngIf=\"oplandb.logInfo.activityDelete && (oplandb.currActivity.id!=-1 && oplandb.currActivity.id!=1 && (oplandb.datActivities.length>0))\"\n            class=\"btn btn-link\" \n            [ngClass]=\"udf.acnBtnClass('delete')\"\n            onfocus=\"this.blur()\"\n            (click)=\"btnClick('DEL')\"\n            [disabled]=\"alertModal() || !disableControls('header') || !disableControls()\"\n            title=\"Delete selected activity...\"\n            style=\"position: relative;\"            \n          >\n          \n          <i id=\"per_delete\" class=\"fa fa-trash-o\" [ngClass]=\"{'text-muted':noAccess(),'text-danger':!noAccess()}\" style=\"font-size:20px;cursor:pointer;\"></i>\n          <!--i *ngIf=\"noAccess()\" class=\"fa fa-ban text-danger\" style=\"z-index:1;position:absolute;top:12px;left:15px;font-size:12px\"></i-->\n          <i *ngIf=\"noAccess()\" class=\"fa fa-ban fa-stack-2x text-danger\" style=\"top:-1px;\"></i>\n\n        </button>\n      </div>\n    </div>\n    <div class=\"card-block py-0\" style=\"position:relative;overflow:hidden;overflow-y:auto;\">\n        <form class=\"my-0 py-0\">\n\n            <div *ngIf=\"dispAlert('delete')\"\n                class=\"alert alert-danger row my-0\" \n                role=\"alert\" \n                style=\"position:absolute;top:0px;left:15px;width:100%;z-index:3;\">\n                <strong>{{oplandb.alertYesNoMessage.label}}:</strong> &nbsp;{{oplandb.alertYesNoMessage.message}}&nbsp;\n                <div class=\"btn-group-sm\">\n                  <button class=\"btn btn-danger\" (click)=\"btnClick('goDelete')\" style=\"cursor:pointer\">{{oplandb.alertYesNoMessage.pos}}</button>\n                  <button class=\"btn btn-secondary\" (click)=\"btnClick('noDelete')\" style=\"cursor:pointer\">{{oplandb.alertYesNoMessage.neg}}</button>\n                </div>\n            </div>  \n            \n            <div *ngIf=\"udf.isActivityNew() && false\" class=\"alert alert-info row\"\n              role=\"alert\"\n              style=\"position:absolute;top:0px;left:15px;width:100%;z-index:3;\">\n              <strong>Attention: </strong>&nbsp;Activity members can only be added after the new activity basic info has been saved.\n            </div>\n\n            <div *ngIf=\"oplandb.msgTempInfoPrompt.length\" class=\"alert alert-info row\"\n              role=\"alert\"\n              style=\"position:absolute;top:0px;left:15px;width:100%;z-index:3;\">\n              <strong>Information: </strong>&nbsp;{{oplandb.msgTempInfoPrompt}}\n            </div>      \n                  \n                    \n            <div *ngIf=\"oplandb.msgTempPrompt.length\" class=\"alert alert-danger row\"\n              role=\"alert\"\n              style=\"position:absolute;top:0px;left:15px;width:100%;z-index:3;\">\n              <strong>Reminder: </strong>&nbsp;{{oplandb.msgTempPrompt}}\n            </div>\n                  \n            <div *ngIf=\"oplandb.msgRequestPrompt.length\" class=\"alert alert-info row\"\n              role=\"alert\"\n              style=\"position:absolute;top:0px;left:15px;width:100%;z-index:3;\">\n              <strong>Server Request: </strong>&nbsp;{{oplandb.msgRequestPrompt}}\n            </div>\n                            \n          <div class=\"row py-0\">\n          <div class=\"col\" >\n\n            <div class=\"row\">\n\n              <div class=\"form-group my-0 pl-3\">\n                  <label  class=\"col-form-label my-0 py-0\" for=\"exampleInputEmail1\">Site</label>\n\n                  <select class=\"form-control form-control-sm\" \n                    id=\"exampleSelect1\" \n                    [(ngModel)]=\"oplandb.currActivity.site\" \n                    [disabled]=\"disableControls('header')\"\n                    name=\"site\" >\n                    <option value=\"\"></option>\n                    <option value=\"SWP\">SWP</option>\n                    <option value=\"OGP\">OGP</option>\n                  </select>      \n\n              </div>\n\n              <div class=\"col form-group my-0\">\n                  <label  class=\"col-form-label my-0 py-0\" for=\"exampleInputEmail1\">Title / Name</label>\n                  <input type=\"text\" class=\"form-control form-control-sm my-0\"\n                    placeholder=\"Enter activity title\"\n                    [(ngModel)]=\"oplandb.currActivity.name\"\n                    [disabled]=\"disableControls('header')\"\n                    id=\"activityName\"\n                    name=\"name\"\n                  >\n              </div>\n            \n            </div>\n        \n            <div class=\"form-group\" style=\"margin-top:2px;\">\n                <label class=\"col-form-label my-0 py-0\" for=\"exampleTextarea\">Description</label>\n                <textarea id=\"desc\" class=\"form-control form-control-sm my-0\" rows=\"4\" \n                  placeholder=\"Enter activity description\"\n                  [(ngModel)]=\"oplandb.currActivity.description\"\n                  [disabled]=\"disableControls('header')\"\n                  name=\"description\"\n                ></textarea>\n            </div>\n\n          \n          </div> <!-- end of left column -->\n\n          <div class=\"mx-3\">  <!-- start of right column -->\n            <div class=\"row\">\n              <div style=\"width:180px;\">\n                  <!-- Priority / Type -->\n                  <div class=\"form-group my-0\">\n                    <label class=\"col-form-label my-0 py-0\">Priority/Type:</label>\n                    <select class=\"form-control form-control-sm\" \n                      id=\"exampleSelect1\" \n                      [(ngModel)]=\"oplandb.currActivity.type\" \n                      [disabled]=\"disableControls('header')\"\n                      name=\"type\" >\n                      <option *ngFor=\"let tp of oplandb.datPriorities\" [ngValue]=\"tp.id\">{{ tp.name }}</option>\n                    </select>      \n                  </div>\n\n\n                  <div class=\"form-group my-0\">\n                    <label class=\"col-form-label my-0 py-0\">Start Date:</label>\n                    <div class=\"input-group input-group-sm mb-2 mr-sm-2 mb-sm-0\">\n                      <input class=\"form-control text-center my-0\" type=\"text\" \n                        id=\"start-date\"\n                        [(ngModel)]=\"oplandb.currActivity.startDate\" \n                        name=\"startDate\"\n                        [dpDayPicker]=\"udf.datePickerConfig\"\n                        onfocus=\"this.blur();\" \n                        (dblclick)=\"oplandb.currActivity.startDate=''\"\n                        [disabled]=\"disableControls('header') || oplandb.currActivity.id==1\"\n                        >\n                      <div class=\"input-group-addon\"><a href=\"#\" (click)=\"openCalendar('start-date')\"><i class=\"fa fa-calendar text-primary\"></i></a></div>\n                    </div>                    \n                  </div>\n\n                  <div class=\"form-group my-0\">\n                    <label class=\"col-form-label my-0 py-0\">End Date</label>\n                    <div class=\"input-group input-group-sm mb-2 mr-sm-2 mb-sm-0\">\n                      <input class=\"form-control text-center my-0\" type=\"text\" \n                        id=\"end-date\" \n                        [(ngModel)]=\"oplandb.currActivity.endDate\" \n                        name=\"endDate\"\n                        [dpDayPicker]=\"udf.datePickerConfig\"\n                        onfocus=\"this.blur();\" \n                        (dblclick)=\"oplandb.currActivity.endDate=''\"\n                        [disabled]=\"disableControls('header')\"\n                        >\n                      <div class=\"input-group-addon\">\n                        <a href=\"#\" (click)=\"openCalendar('end-date')\"><i class=\"fa fa-calendar text-primary\"></i></a>\n                      </div>\n                    </div>                 \n                  </div>   \n\n                </div>\n\n              <div class=\"mx-3\">\n                  <div class=\"form-group my-0\">\n\n                    <label class=\"col-form-label m-0 py-0\">Readiness</label>\n                    <select class=\"form-control form-control-sm\" \n                      style=\"width:120px;\"\n                      [ngStyle]=\"readyColor()\"\n                      [(ngModel)]=\"oplandb.currActivity.ready\" \n                      [disabled]=\"disableControls('header')\"\n                      name=\"ready\" \n                      >\n                        <option style=\"background:white\" value=\"\"></option>\n                        <option style=\"color:yellow;background:green;\" value=\"Green\">Green</option>\n                        <option style=\"color:black;background:orange;\" value=\"Amber\">Amber</option>\n                        <option  style=\"color:white;background:red;\" value=\"Red\">Red</option>\n                    </select>\n                  </div>\n\n                  <div >\n\n                    <label class=\"col-form-label mt-3 m-0 py-0\" style=\"padding-right:22px;\">No POB ?</label>\n                    <button class=\"btn btn-link m-0 py-0 btn-sm\" style=\"cursor:pointer;\"\n                      onfocus=\"this.blur();\"\n                      title=\"Toggle switch for POB and non-POB activity\"\n                      [disabled]=\"disableControls('header')\"\n                      (click)=\"toggleYesNo(oplandb.currActivity,'noPOBCount')\"\n                    >\n                      <i  \n                      class=\"fa\" style=\"font-size:1.3em;\"\n                      [ngClass]=\"clsYesNo(oplandb.currActivity,'noPOBCount')\" aria-hidden=\"true\"></i>\n                    </button>\n                    <br>\n                    <label class=\"col-form-label mt-1 py-0\">Upmanning?</label>\n                    <button class=\"btn btn-link m-0 py-0 btn-sm\" style=\"cursor:pointer;\"\n                      title=\"Toggle switch for UpManning and non-UpManning activity\"\n                      onfocus=\"this.blur();\"\n                      [disabled]=\"disableControls('header')\"\n                      (click)=\"toggleYesNo(oplandb.currActivity,'upmanning')\"\n                    >\n                      <i  \n                      class=\"fa\" style=\"font-size:1.3em;\"\n                      [ngClass]=\"clsYesNo(oplandb.currActivity,'upmanning')\" aria-hidden=\"true\"></i>\n                    </button>\n                    <br>\n                    <label class=\"col-form-label m-0  mt-1 py-0\" style=\"padding-right:35px;\">Vessel ?</label>\n                    <button class=\"btn btn-link m-0 py-0 btn-sm\" style=\"cursor:pointer;\"\n                      onfocus=\"this.blur();\"\n                      title=\"Toggle switch for vessel and non-vessel activity\"\n                      [disabled]=\"disableControls('header')\"\n                      (click)=\"toggleYesNo(oplandb.currActivity,'vessel')\"\n                    >\n                      <i  \n                      class=\"fa\" style=\"font-size:1.3em;\"\n                      [ngClass]=\"clsYesNo(oplandb.currActivity,'vessel')\" aria-hidden=\"true\"></i>\n                    </button>\n                    \n                  </div>\n\n              </div>\n            </div>\n          </div>\n\n            <div *ngIf=\"false\" class=\"col\">\n\n              <div class=\"row\">\n                <div class=\"form-group col my-0\" style=\"width:350px;\">\n                  <label class=\"col-form-label my-0 py-0\">Priority/Type:</label>\n                  <select class=\"form-control form-control-sm\" \n                    id=\"exampleSelect1\" \n                    [(ngModel)]=\"oplandb.currActivity.type\" \n                    [disabled]=\"disableControls('header')\"\n                    name=\"type\" >\n                    <option *ngFor=\"let tp of oplandb.datPriorities\" [ngValue]=\"tp.id\">{{ tp.name }}</option>\n                  </select>      \n                </div>\n              </div>\n\n              <div class=\"row my-0 py-0\">\n                  <div class=\"form-group col my-0\">\n                    <label class=\"col-form-label my-0 py-0\">Start Date:</label>\n                    <div class=\"input-group input-group-sm mb-2 mr-sm-2 mb-sm-0\" style=\"width:200px;\">\n                      <input class=\"form-control text-center my-0\" type=\"text\" \n                        id=\"start-date\"\n                        [(ngModel)]=\"oplandb.currActivity.startDate\" \n                        name=\"startDate\"\n                        [dpDayPicker]=\"udf.datePickerConfig\"\n                        onfocus=\"this.blur();\" \n                        (dblclick)=\"oplandb.currActivity.startDate=''\"\n                        [disabled]=\"disableControls('header')\"\n                        >\n                      <div class=\"input-group-addon\"><a href=\"#\" (click)=\"openCalendar('start-date')\"><i class=\"fa fa-calendar text-primary\"></i></a></div>\n                    </div>                    \n                  </div>\n                  <div class=\"form-group col my-0 p-0\" >\n                    <label class=\"col-form-label my-0 py-0\" style=\"margin-right:12px;\">Vessel ?</label>\n                    <button class=\"btn btn-link m-0 py-0 btn-sm\" \n                      style=\"width:62px;cursor:pointer;\"\n                      onfocus=\"this.blur();\"\n                      [disabled]=\"disableControls('header')\"\n                      (click)=\"toggleYesNo(oplandb.currActivity,'noPOBCount')\"\n                    >\n                      <i  \n                      class=\"fa my-1\" style=\"font-size:1.3em;\"\n                      [ngClass]=\"clsYesNo(oplandb.currActivity,'noPOBCount')\" aria-hidden=\"true\"></i>\n                    </button>\n                    <br>\n\n                    <label class=\"col-form-label m-0 py-0\">No POB ?</label>\n                    <button class=\"btn btn-link m-0 py-0 btn-sm\" style=\"width:62px;cursor:pointer;\"\n                      onfocus=\"this.blur();\"\n                      [disabled]=\"disableControls('header')\"\n                      (click)=\"toggleYesNo(oplandb.currActivity,'noPOBCount')\"\n                    >\n                      <i  \n                      class=\"fa\" style=\"font-size:1.3em;\"\n                      [ngClass]=\"clsYesNo(oplandb.currActivity,'noPOBCount')\" aria-hidden=\"true\"></i>\n                    </button>\n                    \n\n                  </div>\n                </div>\n                <div class=\"row my-0 py-0\">\n                  <div class=\"form-group col my-0\">\n                      <label class=\"col-form-label my-0 py-0\">End Date</label>\n                        <div class=\"input-group input-group-sm mb-2 mr-sm-2 mb-sm-0\" style=\"width:200px;\">\n                          <input class=\"form-control text-center my-0\" type=\"text\" \n                            id=\"end-date\" \n                            [(ngModel)]=\"oplandb.currActivity.endDate\" \n                            name=\"endDate\"\n                            [dpDayPicker]=\"udf.datePickerConfig\"\n                            onfocus=\"this.blur();\" \n                            (dblclick)=\"oplandb.currActivity.endDate=''\"\n                            [disabled]=\"disableControls('header')\"\n                            >\n                          <div class=\"input-group-addon\">\n                            <a href=\"#\" (click)=\"openCalendar('end-date')\"><i class=\"fa fa-calendar text-primary\"></i></a>\n                          </div>\n                        </div>                    \n                  </div>\n                <div class=\"col px-0\" style=\"width:95px;\">\n                  <div class=\"form-group col my-0 p-0 \" >\n                    <label class=\"col-form-label m-0 py-0\">Readiness</label>\n                    <select class=\"form-control form-control-sm\" \n                      style=\"width:100px;\"\n                      [ngStyle]=\"readyColor()\"\n                      [(ngModel)]=\"oplandb.currActivity.ready\" \n                      [disabled]=\"disableControls('header')\"\n                      name=\"ready\" \n                      >\n                        <option style=\"background:white\" value=\"\"></option>\n                        <option style=\"color:yellow;background:green;\" value=\"Green\">Green</option>\n                        <option style=\"color:black;background:orange;\" value=\"Amber\">Amber</option>\n                        <option  style=\"color:white;background:red;\" value=\"Red\">Red</option>\n                    </select>\n                  </div>\n                </div>\n              </div>\n          </div>\n        </div>\n        \n            <div class=\"row py-0 my-0\" style=\"position:relative;top:-15px;\">\n                <div class=\"form-group col my-0 py-0\">\n                  <!--div class=\"form-group row\">\n                    <label class=\"col my-0 py-0\" for=\"exampleInputPassword1\" style=\"height:25px;\">Activity member personnel ({{userID}}):</label>\n                    <div class=\"form-group-col col py-0 my-0\" style=\"text-align:right;\">\n                      <a class=\"btn btn-sm  btn-link btn-info\" href=\"#\"><i class=\"fa fa-refresh\"></i></a>\n                      <a class=\"btn btn-sm btn-link btn-info\" href=\"#\"><i class=\"fa fa-plus\"></i></a>\n                    </div>\n                  </div-->\n\n      <div class=\"row\" *ngIf=\"!oplandb.logInfo.activityMemAdmin\">\n        <label class=\" form-control-label m-2\"></label>\n        Activity Members:\n      </div>\n\n\n<!--Modal section start -->\n\n\n\n<!-- Modal section end -->\n\n\n\n\n\n\n                  <table id=\"acts\" class=\"table table-sm table-hover table-bordered tbc my-1\">\n                    <thead>\n                        <tr>\n                          <th [style.width]=\"(actsColsPX[0])+'px'\">\n                            \n                            <div style=\"display:flex;justify-content:space-between;\">\n                                <span style=\"font-weight:bolder;font-size:1.2em;cursor:pointer;\">\n                                  <i id=\"allExpCol\" class=\"fa\"\n                                    [title]=\"(isExpandAll() ? 'Expand' :  'Collapse')+' all team(s)'\"\n                                    (click)=\"expAllTeamView()\"\n                                    [ngClass]=\"clsExpColAll()\" aria-hidden=\"true\"></i>\n                                </span>\n                                <span>Name</span>\n                                <span>&nbsp;</span>\n                            </div>\n                          </th>\n                          <th [style.width]=\"actsColsPX[1]+'px'\">Start</th>\n                          <th [style.width]=\"actsColsPX[2]+'px'\">End</th>\n                          <th [style.width]=\"actsColsPX[3]+'px'\">MOB</th>\n                          <th [style.width]=\"actsColsPX[4]+'px'\">DEMOB</th>\n                          <th [style.width]=\"actsColsPX[5]+'px'\">Night</th>\n                          <th *ngIf=\"oplandb.logInfo.activityMemAdmin\" [style.width]=\"actsColsPX[6]+'px'\">Del</th>\n                          <th *ngIf=\"oplandb.logInfo.activityMemAdmin\" [style.width]=\"actsColsPX[7]+'px'\">Edit</th>\n                        </tr>\n                    </thead>\n                    <tbody [style.height]=\"actsBodyHeight+'px'\">\n                        <div *ngFor=\"let mem of oplandb.currActivity.members;trackBy:oplandb.trackByItem;let perIdx=index;\">\n                        <tr *ngIf=\"showTeam(mem,perIdx,oplandb.currActivity.members)\" \n                            [style.background]=\"udf.tableGroup1Back\"\n                            [style.color]=\"udf.tableGroup1Fore\"\n                            style=\"font-weight: bold;\">\n                          <td colspan=\"8\" \n                            [style.width]=\"actsTeamPX+'px'\" >\n                            <div style=\"display:flex;justify-content:space-between;\">\n                              <div>\n                                <i \n                                  style=\"cursor:pointer;margin-left:1px;margin-right:5px;font-weight:bolder;font-size:1.2em;\"\n                                  class=\"fa py-1\" aria-hidden=\"true\"\n                                  [title]=\"(!oplandb.idxTeam['r'+mem.teamID].expView ? 'Expand ' : 'Collapse ')+oplandb.idxTeam['r'+mem.teamID].name\"\n                                  [ngClass]=\"{'fa-angle-down':!oplandb.idxTeam['r'+mem.teamID].expView,'fa-angle-up':oplandb.idxTeam['r'+mem.teamID].expView}\"\n                                  (click)=\"expTeamView(oplandb.idxTeam['r'+mem.teamID])\"\n                                  ></i>\n                                  <span>\n                                  {{ oplandb.idxTeam['r'+mem.teamID].name +' ( ' + membersInTeam(mem.teamID) +' )' }}\n                                </span>\n                              </div>\n                              <div class=\"px-1\">\n                                <i *ngIf=\"withConflictInTeam(mem.teamID,'W').withWarning\" \n                                  title=\"With assignment continuation ...\"\n                                    class=\"fa fa-flag text-info\" \n                                    style=\"cursor:pointer;font-size:1.2em\"\n                                    aria-hidden=\"true\"\n                                    ></i>\n                                <i *ngIf=\"withConflictInTeam(mem.teamID,'C').withConflict\" \n                                title=\"Assignment schedule conflict(s) exist ...\"\n                                  class=\"fa fa-flag text-danger\" \n                                  style=\"cursor:pointer;margin-left:5px;font-size:1.2em\"\n                                  aria-hidden=\"true\"></i>\n                                </div>\n                            </div>\n                          </td>\n                        </tr>\n                        <tr *ngIf=\"showTeamGroup(mem,perIdx,oplandb.currActivity.members)\">\n                          <td [style.width]=\"actsTeamPX+'px'\" colspan=\"8\"\n                            style=\"background:#cce9f1;text-transform:uppercase;font-weight:bold;font-size:0.8em;text-align:center;padding-top:2px;padding-bottom:2px;\" >{{teamGroupName(mem)}}</td>\n                        </tr>\n                        <tr *ngIf=\"oplandb.idxTeam['r'+mem.teamID].expView\">\n                          <td [style.width]=\"actsColsPX[0]+'px'\" >\n                            <div style=\"display:flex;justify-content:space-between;\">\n                            <span>{{ oplandb.idxPersonnel['r'+mem.perID].fullName }}</span>\n                            \n                            <div style=\"text-align:right;\">\n                                <span *ngIf=\"noPreMob(mem) && isActiveDuration(mem)\"\n                                  class=\"badge badge-pill badge-info\" style=\"font-weight:normal;cursor:pointer;\"\n                                  title=\"No pre-mobilization requirement set.\">\n                                  No Pre-Mob\n                                </span>\n                                <span *ngIf=\"mem.compWarnCompliant && isActiveDuration(mem)\"\n                                  class=\"badge badge-pill badge-success\" style=\"font-weight:normal;cursor:pointer;\"\n                                  [title]=\"mem.compWarnCompliantTitle\">\n                                  C-{{mem.compWarnCompliant}}\n                                </span>\n                                <span *ngIf=\"mem.compWarnLastDay && isActiveDuration(mem)\"\n                                  class=\"badge badge-pill badge-success\" style=\"font-weight:normal;cursor:pointer;\"\n                                  [title]=\"mem.compWarnLastDayTitle\">\n                                  L-{{mem.compWarnLastDay}}\n                                </span>\n                                <span *ngIf=\"mem.compWarnExpired && isActiveDuration(mem)\"\n                                  class=\"badge badge-pill badge-danger\" style=\"font-weight:normal;cursor:pointer;\"\n                                  [title]=\"mem.compWarnExpiredTitle\">\n                                  E-{{mem.compWarnExpired}}\n                                </span>\n                                <span *ngIf=\"mem.compWarnNotTaken && isActiveDuration(mem)\"\n                                  class=\"badge badge-pill badge-danger\" style=\"font-weight:normal;cursor:pointer;\"\n                                  [title]=\"mem.compWarnNotTakenTitle\">\n\n                                  N-{{mem.compWarnNotTaken}}\n                                </span>\n                                <span *ngIf=\"mem.compWarnExp7Days && isActiveDuration(mem)\"\n                                  class=\"badge badge-pill badge-warning\" style=\"font-weight:normal;cursor:pointer;\"\n                                  [title]=\"mem.compWarnExp7DaysTitle\">\n                                  W-{{mem.compWarnExp7Days}}\n                                </span>\n                                <span *ngIf=\"mem.compWarnExp30Days && isActiveDuration(mem)\"\n                                  class=\"badge badge-pill badge-warning\" style=\"font-weight:normal;cursor:pointer;\"\n                                  [title]=\"mem.compWarnExp30DaysTitle\">\n                                  M-{{mem.compWarnExp30Days}}\n                                </span>\n                            <!--/div> \n                            <div style=\"text-align:right;\" *ngIf=\"mem.conflicts!=null\"-->\n\n                              <span *ngIf=\"mem.warnCount>0\"\n                                [title]=\"mem.warnMessage\"\n                                class=\"badge badge-pill badge-info\" style=\"font-weight:normal;cursor:pointer;\" >\n                                {{mem.warnCount}}\n                              </span>                              \n                              <span *ngIf=\"mem.confCount>0\"\n                                [title]=\"mem.confMessage\"\n                                class=\"badge badge-pill badge-danger\" style=\"font-weight:normal;cursor:pointer;\" >\n                                {{mem.confCount}}\n                              </span>\n                            </div>\n                            </div>\n                          </td>\n                          <td *ngIf=\"false\" [style.width]=\"actsColsPX[1]+'px'\" \n                            style=\"text-align:center;\">\n                            <i *ngIf=\"withConflict(mem)\"\n                              class=\"fa fa-exclamation-triangle text-danger\" \n                              title=\"conflict!\" \n                              aria-hidden=\"true\">\n                            </i>&nbsp;\n                          </td>\n                          <td [style.width]=\"actsColsPX[1]+'px'\" style=\"text-align:center;\">\n                              {{ mem.startDate }}\n                              &nbsp;<i *ngIf=\"oplandb.isActMemNoFlight(mem,false).noFlight\" \n                              class=\"fa fa-plane text-danger p-0 m-0\" \n                              [title]=\"'No flight available on mob date ' + oplandb.isActMemNoFlight(mem,false).date +'.'\" \n                              style=\"font-size:1.2em;cursor:pointer;\"></i>\n                          </td>\n                          <td [style.width]=\"actsColsPX[2]+'px'\" style=\"text-align:center;\">\n                            {{ mem.endDate }}\n                            &nbsp;<i *ngIf=\"oplandb.isActMemNoFlight(mem,true).noFlight\" \n                              class=\"fa fa-plane text-danger p-0 m-0\" \n                              [title]=\"'No flight available on demob date ' + oplandb.isActMemNoFlight(mem,true).date +'.'\"\n                              style=\"font-size:1.2em;cursor:pointer;\"></i>\n                          </td>\n                          <td [style.width]=\"actsColsPX[3]+'px'\" \n                             style=\"text-align:center;\">\n                              <i class=\"fa\"\n                                [class.fa-check-circle]=\"mem.mob\"\n                                [class.text-success]=\"mem.mob\"\n                                [class.fa-times-circle]=\"!mem.mob\"\n                                [class.text-muted]=\"!mem.mob\"\n                                >\n                              </i>\n                              \n                        </td>\n                          <td [style.width]=\"actsColsPX[4]+'px'\" style=\"text-align:center;\">\n                              <i class=\"fa\"\n                                [class.fa-check-circle]=\"mem.demob\"\n                                [class.text-success]=\"mem.demob\"\n                                [class.fa-times-circle]=\"!mem.demob\"\n                                [class.text-muted]=\"!mem.demob\"\n                                >\n                              </i>\n                              <i class=\"fa fa-sun-o\" [style.color]=\"udf.menuBackground\" *ngIf=\"mem.isDay\"></i>\n                          </td>\n                          <td [style.width]=\"actsColsPX[5]+'px'\" style=\"text-align:center;\">\n                              <i class=\"fa\"\n                                [class.fa-check-circle]=\"mem.night\"\n                                [class.text-success]=\"mem.night\"\n                                [class.fa-times-circle]=\"!mem.night\"\n                                [class.text-muted]=\"!mem.night\"\n                                >\n                              </i>\n                          </td>\n                          <td *ngIf=\"oplandb.logInfo.activityMemAdmin\"\n                           [style.width]=\"actsColsPX[6]+'px'\" style=\"text-align:center;\" class=\"acn\">\n                              <button \n                                [disabled]=\"alertModal() || !disableControls('header')\"\n                                class=\"btn btn-link btn-sm text-danger\"\n                                (click)=\"btnClick('del',mem)\"\n                                style=\"cursor:pointer;\"\n                                >\n                                \n                                  <i class=\"fa fa-times\"></i></button>\n                          </td>\n                          <td *ngIf=\"oplandb.logInfo.activityMemAdmin\"\n                              [style.width]=\"actsColsPX[7]+'px'\" style=\"text-align:center;\" class=\"acn\">\n\n                              <button class=\"btn btn-sm btn-link\" \n                              data-toggle=\"modal\"\n                              data-target=\"#exampleModal\"\n                              (click)=\"btnClick('edit',mem)\"\n                              title=\"Edit activity member ...\"\n                              style=\"cursor:pointer;\"\n                              [disabled]=\"alertModal() || !disableControls('header')\"\n                            >\n                            <i class=\"fa fa-pencil\" style=\"color:#CAAD1B;\"></i></button>  \n                                                          \n\n\n                              <!--button class=\"btn btn-link btn-sm m-0\" \n                              data-toggle=\"modal\"\n                              data-target=\"#exampleModal\"\n                              [disabled]=\"alertModal() || !disableControls('header') || !disableControls()\"\n                              (click)=\"btnClick('edit',mem)\"\n                              style=\"cursor:pointer;\"\n                              >\n                                <i class=\"fa fa-pencil\" style=\"color:#CAAD1B;\"></i></button-->\n\n\n\n                          </td>\n                        </tr>\n                      </div>\n                    </tbody>\n                  </table>\n                  <!--div style=\"background:yellow;\">New div</div-->\n                </div>\n              </div>\n\n        </form>\n    </div>\n    <div class=\"card-footer\" *ngIf=\"!disableControls('header')\">\n      <div class=\"form-buttons align-bottom\"  >\n          <button type=\"button\" class=\"btn btn-secondary\"\n          [disabled]=\"disableControls('header')\"\n          (click)=\"btnClick('CANCEL')\">Cancel</button>\n        <button type=\"button\" class=\"btn btn-primary\"\n          [disabled]=\"disableControls('header')\"\n          (click)=\"btnClick('SAVE')\" >Save</button>\n      </div>\n  </div>    \n  </div>\n\n  <!-- Modal -->\n<div class=\"modal fade\" id=\"exampleModal\" tabindex=\"-1\" role=\"dialog\" aria-labelledby=\"exampleModalLabel\" aria-hidden=\"true\">\n    <div class=\"modal-dialog\" role=\"document\">\n      <div class=\"modal-content\">\n        <div class=\"modal-header\">\n          <h5 class=\"modal-title\" id=\"exampleModalLabel\">Activity Member</h5>\n          <button type=\"button\" \n            class=\"close\" data-dismiss=\"modal\" \n            (click)=\"btnClick('cancel')\" aria-label=\"Close\">\n            <span aria-hidden=\"true\">&times;</span>\n          </button>\n        </div>\n        <div class=\"modal-body\" style=\"position:relative;font-size:0.9em\">\n\n          <div *ngIf=\"oplandb.msgTempPromptModal.length>0\"class=\"alert alert-danger row\"\n            role=\"alert\"\n            style=\"position: absolute;z-index:20;top:0px;left:15px;width:100%;\">\n            <strong>Reminder:</strong>  {{oplandb.msgTempPromptModal}}\n          </div>\n\n\n            <div class=\"row px-3\">\n                <div class=\"form-group\" >\n                  <label class=\"col-form-label my-0 py-0\">\n                    {{displayLabel()}}\n                  </label><br>\n        \n                  <app-ctl-user-select \n                    [width]=\"468\"\n                    [height]=\"250\" \n                    [(counter)]=\"oplandb.newActivityMember.perID\" \n        \n                    [lookup]=\"oplandb.idxPersonnel\"\n                    [displayField]=\"'fullName'\"\n                    [valueField]=\"'id'\"\n        \n                    \n                    \n                    [source]=\"oplandb.datTeams\"\n                    [groupName]=\"'name'\"\n                    [dataName]=\"'members'\"\n                    [valueColumn]=\"'perID'\"\n                    [displayColumn]=\"''\"\n                    (counterChange)=\"onUserChange($event)\"\n                    (itemSelected)=\"onUserItemSelected($event)\"\n                    >\n                  </app-ctl-user-select>\n                </div>  \n            </div>\n<!--[disableSelect]=\"disableControls() || udf.isActivityMemberEdit()\"-->\n        \n            <div class=\"row px-3\">\n\n              <label class=\"col-form-label\" style=\"margin-top:0px;padding-top:0px;\">Functional Sponsor / Company</label>\n\n              <div class=\"input-group input-group\" >\n                  <input type=\"text\" class=\"form-control form-control-sm\" id=\"firstName\"  \n                    [value]=\"oplandb.getCompanyName(oplandb.newActivityMember.coyID)\" \n                    onfocus=\"this.blur()\"\n                  >           \n      \n                <!-- custom dropdown use \"btn-group dropdown\" class [start] -->\n                <div class=\"btn-group dropdown\">\n                  <button type=\"button\" \n                    class=\"btn p-0 px-2 text-primary\" data-toggle=\"dropdown\" \n                    title=\"Click to select company...\"\n                    style=\"cursor: pointer;border-radius:2px;border:1px solid #c0c0c0;\"\n                    aria-haspopup=\"true\" aria-expanded=\"false\"\n                    >\n                    <i class=\"fa fa-list-alt\"></i>\n                  </button>\n  \n                  <app-ctl-drop-down *ngIf=\"this.loadLookups\" class=\"dropdown-menu p-0\"\n                    [drpdData]=\"oplandb.datCompany\"\n                    [(drpdValue)]=\"oplandb.newActivityMember.coyID\" \n                    [drpdDispRows]=\"15\"\n                    \n                    style=\"z-index:100;left:-150px;\"\n                  ></app-ctl-drop-down>          \n                  \n                </div> \n                <!-- custom dropdown use \"btn-group dropdown\" class [end] -->\n      \n              </div>\n\n              <!--select *ngIf=\"loadLookups\" class=\"form-control form-control-sm\"  (onload)=\"selectOnload()\"\n                id=\"exampleSelect1\" \n                [disabled]=\"isCoreTeam(oplandb.newActivityMember.teamID)\"\n                [(ngModel)]=\"oplandb.newActivityMember.coyID\" \n                name=\"coyID\" >\n                <option *ngFor=\"let coy of oplandb.datCompany\" [ngValue]=\"coy.id\">{{ coy.name }}</option>\n              </select-->      \n\n            </div>\n\n            <div class=\"row px-3\">\n              \n              <label class=\"col-form-label\">Position</label>\n\n              <div class=\"input-group input-group\" >\n                  <input type=\"text\" class=\"form-control form-control-sm\" id=\"firstName\"  \n                    onfocus=\"this.blur()\"\n                    [value]=\"oplandb.getPositionName(oplandb.newActivityMember.posID)\" \n                  >           \n      \n                <!-- custom dropdown use \"btn-group dropdown\" class [start] -->\n                <div class=\"btn-group dropdown\">\n                  <button type=\"button\" \n                    class=\"btn p-0 px-2 text-primary\" data-toggle=\"dropdown\" \n                    title=\"Click to select position...\"\n                    style=\"cursor: pointer;border-radius:2px;border:1px solid #c0c0c0;\"\n                    aria-haspopup=\"true\" aria-expanded=\"false\"\n                    >\n                    <i class=\"fa fa-list-alt\"></i>\n                  </button>\n  \n                  <app-ctl-drop-down *ngIf=\"this.loadLookups\" class=\"dropdown-menu p-0\"\n                    [drpdData]=\"oplandb.datPosition\"\n                    [(drpdValue)]=\"oplandb.newActivityMember.posID\" \n                    [drpdDispRows]=\"15\"\n                    \n                    style=\"z-index:100;left:-150px;\"\n                  ></app-ctl-drop-down>          \n                  \n                </div> \n                <!-- custom dropdown use \"btn-group dropdown\" class [end] -->\n      \n              </div>\n              \n\n              <!--select *ngIf=\"loadLookups\" class=\"form-control form-control-sm\" \n                id=\"exampleSelect1\" \n                [disabled]=\"isCoreTeam(oplandb.newActivityMember.teamID)\"\n                [(ngModel)]=\"oplandb.newActivityMember.posID\" \n                name=\"posID\" >\n                <option *ngFor=\"let pos of oplandb.datPosition\" [ngValue]=\"pos.id\">{{ pos.name }}</option>\n              </select-->      \n\n            </div>\n             <!-- start of activity member fields 8888 -->\n              <div class=\"row px-3\" style=\"display:flex;justify-content:space-between;padding-top:5px;\">\n                <div class=\"form-group\">\n                  <label class=\"col-form-label my-0 py-0\">Start</label>\n                  <i *ngIf=\"oplandb.isActMemNoFlight(oplandb.newActivityMember,false).noFlight\" \n                  class=\"fa fa-plane text-danger p-0 m-0\" \n                  [title]=\"'No flight available on mob date ' + oplandb.isActMemNoFlight(oplandb.newActivityMember,false).date +'.'\" \n                  style=\"font-size:1.2em;cursor:pointer;margin-left:10px;\"></i>\n                  \n                  <!--td [style.width]=\"actsColsPX[1]+'px'\" style=\"text-align:center;\">\n                      {{ mem.startDate }}\n                      &nbsp;<i *ngIf=\"oplandb.isActMemNoFlight(mem,false).noFlight\" \n                      class=\"fa fa-plane text-danger p-0 m-0\" \n                      [title]=\"'No flight available on mob date ' + oplandb.isActMemNoFlight(mem,false).date +'.'\" \n                      style=\"font-size:1.2em;cursor:pointer;\"></i-->\n\n                    <div class=\"input-group input-group-sm\" >\n                      <input class=\"form-control text-center\" \n                        [dpDayPicker]=\"udf.datePickerConfig\"\n                        [(ngModel)]=\"oplandb.newActivityMember.startDate\"\n                        id=\"memStart\"\n                        onfocus=\"this.blur()\"\n                        name=\"assigmentStartDate\"\n                        type=\"text\"\n                        style=\"width:170px;\"\n                        >\n                        <div class=\"input-group-addon mx-0\">\n                            <a href=\"#\" (click)=\"openCalendar('memStart')\"><i class=\"fa fa-calendar text-primary\"></i></a>\n                        </div>\n              </div>                    \n                </div> \n                \n                <div class=\"form-group\">\n                  <label class=\"col-form-label py-0 my-0\">End</label>\n                  <i *ngIf=\"oplandb.isActMemNoFlight(oplandb.newActivityMember).noFlight\" \n                  class=\"fa fa-plane text-danger p-0 m-0\" \n                  [title]=\"'No flight available on demob date ' + oplandb.isActMemNoFlight(oplandb.newActivityMember).date +'.'\" \n                  style=\"font-size:1.2em;cursor:pointer;margin-left:10px;\"></i>\n                <div class=\"input-group input-group-sm\" >\n                      <input class=\"form-control text-center\"\n                        [dpDayPicker]=\"udf.datePickerConfig\"\n                        [(ngModel)]=\"oplandb.newActivityMember.endDate\"\n                        name=\"assigmentEndDate\"\n                        onfocus=\"this.blur()\"\n                        id=\"memEnd\"\n                        type=\"text\"\n                        style=\"width:170px;\"\n                        >\n                      <div class=\"input-group-addon mx-0\">\n                          <a href=\"#\" (click)=\"openCalendar('memEnd')\"><i class=\"fa fa-calendar text-primary\"></i></a>\n                      </div>\n                    \n                    </div>                    \n                </div>\n              </div>\n\n              <div class=\"row px-3\">\n                <!--*ngIf=\"isCoreTeam(oplandb.newActivityMember.teamID)\"-->\n                <div class=\"form-group\" style=\"width:200px;\">\n\n                    <label class=\"col-form-label my-0 py-0\">Group</label>\n                    <select class=\"form-control form-control-sm\" \n                      id=\"selGrpID\" \n                      [(ngModel)]=\"oplandb.newActivityMember.grpID\" \n                      name=\"grpID\" >\n                      <option *ngFor=\"let grp of oplandb.config.teamGroupsArr\" [ngValue]=\"grp.id\">{{ grp.name }}</option>\n                    </select>    \n\n                </div>\n                <div class=\"col\">\n                  <div class=\"row\">\n                      <label class=\"col-3 text-center my-0\">Mob?</label>\n                      <label class=\"col-3 text-center my-0\">Demob?</label>\n                      <label class=\"col-3 text-center my-0\">Night Shift?</label>\n                      <label class=\"col-3 text-center my-0\"  *ngIf=\"isDayVisible()\">Day Trip?</label>\n                    </div>\n                  <div class=\"row\">\n                      <label class=\"col-3 text-center\">\n                        <i (click)=\"toggleYesNo(oplandb.newActivityMember,'mob')\" \n                        class=\"fa my-1\" style=\"font-size:1.3em;cursor:pointer;\"\n                        [ngClass]=\"clsYesNo(oplandb.newActivityMember,'mob')\" aria-hidden=\"true\"></i>\n                      </label>\n                      <label class=\"col-3 text-center\">\n                          <i (click)=\"toggleYesNo(oplandb.newActivityMember,'demob')\" \n                          class=\"fa my-1\" style=\"font-size:1.3em;cursor:pointer;\"\n                          [ngClass]=\"clsYesNo(oplandb.newActivityMember,'demob')\" aria-hidden=\"true\"></i>\n                      </label>\n                      <label class=\"col-3 text-center\">\n                          <i (click)=\"toggleYesNo(oplandb.newActivityMember,'night')\" \n                          class=\"fa my-1\" style=\"font-size:1.3em;cursor:pointer;\"\n                          [ngClass]=\"clsYesNo(oplandb.newActivityMember,'night')\" aria-hidden=\"true\"></i>\n                      </label>\n                      <label class=\"col-3 text-center\" *ngIf=\"isDayVisible()\">\n                        <i (click)=\"toggleYesNo(oplandb.newActivityMember,'isDay')\" \n                        class=\"fa my-1\" style=\"font-size:1.3em;cursor:pointer;\"\n                        [ngClass]=\"clsYesNo(oplandb.newActivityMember,'isDay')\" aria-hidden=\"true\"></i>\n                      </label>\n                </div>                    \n                    \n                </div>\n\n            </div>\n\n        </div>\n        <div class=\"modal-footer\">\n          <button \n            id=\"btnClose\"\n            type=\"button\" \n            class=\"btn btn-secondary\" \n            data-dismiss=\"modal\" \n            (click)=\"btnClick('cancel')\">Close</button>\n          <button type=\"button\" class=\"btn btn-primary\" \n            (click)=\"btnClick('ok')\">Save changes</button>\n        </div>\n      </div>\n    </div>\n  </div>"

/***/ }),

/***/ "../../../../../src/app/components/frm-activity/frm-activity.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return FrmActivityComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__models_data_models__ = __webpack_require__("../../../../../src/app/models/data-models.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2_jquery__ = __webpack_require__("../../../../jquery/dist/jquery.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2_jquery___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_2_jquery__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3_ng2_date_picker__ = __webpack_require__("../../../../ng2-date-picker/index.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3_ng2_date_picker___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_3_ng2_date_picker__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_4__app_globals__ = __webpack_require__("../../../../../src/app/app.globals.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_5__app_udfs__ = __webpack_require__("../../../../../src/app/app.udfs.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_6__app_oplandata__ = __webpack_require__("../../../../../src/app/app.oplandata.ts");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};







var FrmActivityComponent = (function () {
    function FrmActivityComponent(udf, oplandb) {
        this.udf = udf;
        this.oplandb = oplandb;
        this.testDate = "02/10/2017";
        this.loadLookups = false;
        //pikaday:Pikaday;
        this.actAddFldSpc = 5;
        this.actsCols = ["100%", 140, 140, 60, 60, 90, 40, 40];
        this.actsTeamCol = ["100%"];
        this.actsBodyHeight = 150;
        this.userID = 5;
        this.currActivity = new __WEBPACK_IMPORTED_MODULE_1__models_data_models__["c" /* dataActivity */]();
    }
    FrmActivityComponent.prototype.open = function () {
        this.datePicker.api.open();
        alert('open');
    };
    FrmActivityComponent.prototype.close = function () {
        this.datePicker.api.close();
        alert('close');
    };
    FrmActivityComponent.prototype.ngOnDestroy = function () {
        this.oplandb.clearSubscription(this.dataSubscription);
    };
    FrmActivityComponent.prototype.ngOnInit = function () {
        var _this = this;
        var sss = "hello 5/31/2018 3:08:28 PM keyu";
        var reg = new RegExp("[^0-9a-zA-Z]+");
        var res = sss.replace("[^0-9a-zA-Z]+", "");
        var newString = sss.replace(/[^A-Z0-9]+/ig, "");
        this.udf.isActivityMemberNormal(true);
        this.udf.isActivityNormal(true);
        this.loadLookups = false;
        this.oplandb.clearAllSubscription();
        // the following lines seem to do nothing yet???
        //$(document).on('show.bs.modal','#exampleModal', function () {
        //  alert('hi');
        //})
        __WEBPACK_IMPORTED_MODULE_2_jquery__(window).on('shown.bs.modal', function () {
            //$('#exampleModal').modal('show');
            alert('shown');
        });
        // this.oplandb.activityModifiedInCalendar
        /*
            let noforce:boolean=false;
            if(this.oplandb.activityModifiedInCalendar && !noforce){
        
              this.oplandb.activityModifiedInCalendar=false;
        
              this.oplandb.clearSubscription(this.oplandb.subsActivities )
        
              this.oplandb.subsActivities = this.oplandb.getActivities(true,data=>{
                this.oplandb.setOtherNavGroup();
              });
        
        
            }else{
              this.oplandb.setOtherNavGroup();
            }*/
        this.oplandb.setOtherNavGroup();
        setTimeout(function () { return _this.setTableColWidths(); }, 100);
        this.setTableColWidths();
        this.dataSubscription = this.oplandb.getCredentials();
        this.udf.isActivityNormal(true);
        // set this flag to force a refresh on the personnel active activity details
        this.oplandb.getActiveActivities = true;
    };
    FrmActivityComponent.prototype.parseDate = function (dateString) {
        if (dateString) {
            return new Date(dateString);
        }
        else {
            return null;
        }
    };
    FrmActivityComponent.prototype.onResize = function (e) {
        this.setTableColWidths();
    };
    FrmActivityComponent.prototype.onUserChange = function (e) {
    };
    FrmActivityComponent.prototype.onUserItemSelected = function (e) {
        var mem = this.oplandb.newActivityMember;
        mem.teamID = e['team']['id'];
        //if(this.oplandb.newActivityMember.id!=-1)return;  // no action if item is being modified
        var per = e['user'];
        var per2 = this.oplandb.idxPersonnel['r' + per.perID];
        mem.coyID = String(per2.companyID);
        mem.posID = String(per2.positionID);
    };
    FrmActivityComponent.prototype.btnShowHide = function (type) {
        if (this.udf.isActivityMemberEdit() || this.udf.isActivityMemberNew()) {
            return (type == 'ok' || type == 'cancel');
        }
        else {
            return (type == 'add');
        }
    };
    FrmActivityComponent.prototype.readyColor = function () {
        var color = "White";
        var background = "White";
        switch (this.oplandb.currActivity.ready) {
            case "Red":
                background = "Red";
                break;
            case "Green":
                background = "Green";
                break;
            case "Amber":
                background = "Orange";
                color = "Black";
                break;
            default:
        }
        return { "color": color, "background": background, "font-weight": "bold" };
    };
    FrmActivityComponent.prototype.isDayVisible = function () {
        var ret = false;
        if (this.oplandb.newActivityMember.startDate.length > 0 && this.oplandb.newActivityMember.endDate.length > 0) {
            ret = this.oplandb.newActivityMember.mob && this.oplandb.newActivityMember.demob && (this.oplandb.newActivityMember.startDate == this.oplandb.newActivityMember.endDate);
        }
        return ret;
    };
    FrmActivityComponent.prototype.btnClick = function (type, item) {
        var _this = this;
        switch (type) {
            case 'ok':
                this.oplandb.saveActivityMember(function (data) {
                    _this.setTableColWidths();
                });
                break;
            case 'cancel':
                this.oplandb.cancelActivityMember();
                break;
            case 'edit':
                //this.oplandb.setLoading("Loading activity member details form. Please wait...");
                if (!this.noAccess()) {
                    setTimeout(function () {
                        _this.loadLookups = true;
                        _this.oplandb.editActivityMember(item);
                    }, 100);
                }
                break;
            case 'del':
                if (!this.noAccess())
                    this.oplandb.deleteActivityMember(item);
                break;
            case 'add':
                //this.oplandb.setLoading("Loading activity member details form. Please wait...");
                if (!this.noAccess()) {
                    setTimeout(function () {
                        _this.loadLookups = true;
                        _this.udf.isActivityMemberNew(true);
                    }, 100);
                }
                break;
            case 'SAVE':
                if (this.oplandb.saveActivity()) {
                    setTimeout(function () { _this.setTableColWidths(); }, 100);
                }
                break;
            case 'CANCEL':
                this.oplandb.cancelActivity();
                setTimeout(function () { _this.setTableColWidths(); }, 50);
                break;
            case 'ADD':
                if (!this.noAccess()) {
                    this.oplandb.newActivity();
                    __WEBPACK_IMPORTED_MODULE_4__app_globals__["setFocus"]('activityName');
                    setTimeout(function () { _this.setTableColWidths(); }, 50);
                }
                break;
            case 'DEL':
                if (!this.noAccess())
                    this.oplandb.deleteActivity();
                break;
            case 'EDIT':
                if (!this.noAccess()) {
                    this.udf.isActivityEditing(true);
                    __WEBPACK_IMPORTED_MODULE_4__app_globals__["setFocus"]('activityName');
                    setTimeout(function () { _this.setTableColWidths(); }, 50);
                }
                break;
            case 'goDelete':
                // put delete action here then normalize mode
                if (this.udf.isActivityDeleteAlert()) {
                    this.oplandb.deleteActivity(true);
                }
                else if (this.udf.isActivityMemberDeleteAlert()) {
                    this.oplandb.deleteActivityMember(undefined, true);
                }
                break;
            case 'noDelete':
                // cancel delete request
                this.udf.isActivityNormal(true);
                this.udf.isActivityMemberNormal(true);
                break;
            default:
        }
    };
    FrmActivityComponent.prototype.selectOnload = function () {
        alert('ngAfterContentInit');
    };
    FrmActivityComponent.prototype.alertModal = function () {
        return this.udf.isActivityMemberDeleteAlert() ||
            this.udf.isActivityDeleteAlert();
    };
    FrmActivityComponent.prototype.dispAlert = function (mode) {
        /*if(mode=='delete'){
    
        }*/
        if (this.udf.isActivityMemberDeleteAlert() || this.udf.isActivityDeleteAlert()) {
            return (mode == 'delete');
        }
        return false;
    };
    FrmActivityComponent.prototype.toggleYesNo = function (obj, fld) {
        obj[fld] = !obj[fld];
    };
    FrmActivityComponent.prototype.clsYesNo = function (obj, fld) {
        return {
            'fa-check-circle': obj[fld],
            'fa-times-circle': !obj[fld],
            'text-success': obj[fld],
            'text-muted': !obj[fld]
        };
    };
    FrmActivityComponent.prototype.disableControls = function (mode) {
        var retVal = false;
        if (mode == undefined)
            mode = '';
        if (mode == 'header') {
            retVal = this.udf.isActivityNormal() || this.alertModal();
        }
        else if (mode == 'all') {
            retVal = this.udf.isActivityNormal() || this.udf.isActivityMemberNormal() || this.alertModal();
        }
        else {
            retVal = this.udf.isActivityMemberNormal() || this.alertModal();
        }
        return retVal || this.oplandb.isRequesting();
    };
    FrmActivityComponent.prototype.setTableColWidths = function () {
        var ftr = __WEBPACK_IMPORTED_MODULE_2_jquery__('.card-footer');
        var footOff = 0;
        if (ftr.length)
            footOff = ftr.height();
        this.actsBodyHeight = this.udf.getReactiveTableHeight('acts', 'actCard', 25, 90) - footOff;
        var body = __WEBPACK_IMPORTED_MODULE_2_jquery__('#acts').children('tbody');
        var isOverFlow = (body.height() < body.prop('scrollHeight'));
        var refWidth = __WEBPACK_IMPORTED_MODULE_2_jquery__('#acts').width() - 20; //-(isOverFlow ? 22 : 0);
        this.actsTeamPX = refWidth - this.actsCols.length - 1;
        var inpArr = this.oplandb.cloneObject(this.actsCols);
        if (!this.oplandb.logInfo.activityMemAdmin) {
            inpArr[inpArr.length - 1] = 0;
            inpArr[inpArr.length - 2] = 0;
        }
        this.actsColsPX = this.udf.getAbsWidthVals(inpArr, refWidth, 0);
    };
    FrmActivityComponent.prototype.showTeam = function (item, idx, group) {
        var retVal = (idx == 0);
        if (!retVal)
            retVal = group[idx - 1]['teamID'] != item['teamID'];
        return retVal;
    };
    FrmActivityComponent.prototype.showTeamGroup = function (item, idx, group) {
        if (item != undefined) {
            var teamExp = this.oplandb.idxTeam['r' + item['teamID']].expView;
            var retVal = (idx == 0) && item['grpID'] != 0;
            if (!retVal) {
                if (group[idx - 1] != undefined) {
                    retVal = group[idx - 1]['grpID'] != item['grpID'];
                }
            }
            //return this.oplandb.idxTeam['r'+item['teamID']].expView && item['grpID']!=0;
            retVal = retVal && teamExp;
            if (retVal)
                retVal = retVal && (this.teamGroupName(item).length != 0);
            return retVal;
        }
        else {
            return false;
        }
    };
    FrmActivityComponent.prototype.noPreMob = function (mem) {
        return mem.compWarnCompliant == 0 &&
            mem.compWarnExp30Days == 0 &&
            mem.compWarnExp7Days == 0 &&
            mem.compWarnExpired == 0 &&
            mem.compWarnLastDay == 0 &&
            mem.compWarnNotTaken == 0 &&
            this.oplandb.idxPersonnel['r' + mem.perID].fullName.substr(0, 3).toLowerCase() != 'tba';
    };
    FrmActivityComponent.prototype.isActiveDuration = function (mem) {
        var ret;
        var dt = new Date();
        var edt = __WEBPACK_IMPORTED_MODULE_4__app_globals__["longToDate"](mem.endDate);
        ret = (edt >= dt);
        //endDate
        return ret;
    };
    FrmActivityComponent.prototype.teamGroupName = function (mem) {
        var grpName = "";
        if (mem != undefined) {
            var grp = this.oplandb.config.teamGroupsArr.find(function (e) { return e.id == mem.grpID; });
            if (grp)
                grpName = grp.name;
        }
        return grpName;
    };
    FrmActivityComponent.prototype.membersInTeam = function (memberTeamID) {
        var arr = this.oplandb.currActivity.members
            .filter(function (item) { return item.teamID == memberTeamID; });
        return arr.length;
    };
    FrmActivityComponent.prototype.saveActivity = function () {
    };
    FrmActivityComponent.prototype.cancelUpdate = function () {
    };
    FrmActivityComponent.prototype.clickMe = function (e) {
        var params = { start: "20170401", end: "20170525", activityID: 1,
            personParams: [
                { personID: 1277, start: "20170405", end: "20170410" },
                { personID: 494, start: "20170401", end: "20170525" }
            ]
        };
        //let obj:Observable:Array<Buffer>;
        var subs = this.oplandb.http.post(__WEBPACK_IMPORTED_MODULE_4__app_globals__["getUrlRequest"]("oplanData", "getActivityValidation"), JSON.stringify(params))
            .subscribe(function (data) {
            subs.unsubscribe();
        }, function (err) {
        });
    };
    /*  isWithTeamSeleced():boolean{
        return (this.oplandb.newActivityMember.teamID!=0)
      }*/
    FrmActivityComponent.prototype.displayLabel = function () {
        var mem = this.oplandb.newActivityMember;
        var retVal = "Select a member...";
        if (mem.teamID)
            retVal = this.oplandb.idxTeam['r' + mem.teamID]['name'];
        return retVal;
    };
    FrmActivityComponent.prototype.addNewMember = function () {
        this.udf.isActivityMemberNew(true);
    };
    FrmActivityComponent.prototype.teamClicked = function (item) {
    };
    FrmActivityComponent.prototype.expTeamView = function (team) {
        team.expView = !team.expView;
    };
    FrmActivityComponent.prototype.isExpandAll = function () {
        var fa = __WEBPACK_IMPORTED_MODULE_2_jquery__('#allExpCol');
        if (fa.length == 0)
            return false;
        return fa.hasClass('fa-angle-double-down');
    };
    FrmActivityComponent.prototype.expAllTeamView = function () {
        var _this = this;
        var fa = __WEBPACK_IMPORTED_MODULE_2_jquery__('#allExpCol');
        if (fa.length == 0)
            return;
        //alert(fa.children('.fa-angle-double-down').length);
        var mems = this.oplandb.currActivity.members;
        var teams = __WEBPACK_IMPORTED_MODULE_4__app_globals__["groupBy"](mems, "teamID");
        var isExpandAll = fa.hasClass('fa-angle-double-down');
        teams.forEach(function (element) {
            var team = _this.oplandb.idxTeam['r' + element.type];
            if (!team.expView)
                team.expView = true;
            if (isExpandAll && !team.expView) {
                team.expView = true;
            }
            else if (!isExpandAll && team.expView) {
                team.expView = false;
            }
        });
    };
    FrmActivityComponent.prototype.clsExpColAll = function () {
        var _this = this;
        //oplandb.currActivity.members
        var mems = this.oplandb.currActivity.members;
        var teams = __WEBPACK_IMPORTED_MODULE_4__app_globals__["groupBy"](mems, "teamID");
        var expCnt = 0;
        teams.forEach(function (element) {
            var team = _this.oplandb.idxTeam['r' + element.type];
            if (!team.expView)
                expCnt++;
        });
        var half = teams.length / 2;
        var isHalf = (expCnt <= parseInt(String(half)));
        return {
            'fa-angle-double-down': !isHalf,
            'fa-angle-double-up': isHalf
        };
    };
    FrmActivityComponent.prototype.withConflict = function (mem) {
        //oplan.currActivity.conflicts['r'+mem.id]!=undefined
        //return true;
        if (this.oplandb.currActivity.conflicts == null) {
            return false;
        }
        else {
            return this.oplandb.currActivity.conflicts['r' + mem.id] != null;
        }
    };
    FrmActivityComponent.prototype.withConflictInTeam = function (currTeamID, retType, noTeam) {
        var ret = null;
        var act = this.oplandb.currActivity;
        var m;
        var confArr = [];
        var warnArr = [];
        if (retType == undefined)
            retType = "";
        if (noTeam == undefined)
            noTeam = false;
        if (retType == "" || retType == "C") {
            confArr = act.members.filter(function (mem) {
                return (mem.teamID == currTeamID || noTeam) && mem.confCount > 0;
            });
        }
        if (retType == "" || retType == "W") {
            warnArr = act.members.filter(function (mem) {
                return (mem.teamID == currTeamID || noTeam) && mem.warnCount > 0;
            });
        }
        return {
            W: warnArr, withWarning: (warnArr.length > 0),
            C: confArr, withConflict: (confArr.length > 0),
            E: (warnArr.length > 0 || confArr.length > 0)
        };
    };
    FrmActivityComponent.prototype.openCalendar = function (id) {
        __WEBPACK_IMPORTED_MODULE_4__app_globals__["setFocus"](id);
    };
    FrmActivityComponent.prototype.reloadActivity = function () {
    };
    FrmActivityComponent.prototype.toggleShowInCalendar = function () {
        if (this.noAccess())
            return;
        this.oplandb.currActivity.showInChart = !this.oplandb.currActivity.showInChart;
        this.oplandb.currActivityObj.showInChart = this.oplandb.currActivity.showInChart;
        // post changes
        this.dataSubscription = this.oplandb.processData({
            method: "saveActivityCalendarVisibility",
            data: {
                "id": this.oplandb.currActivity.id,
                "showInChart": this.oplandb.currActivity.showInChart,
            },
            isPosting: true,
            processMessage: "Saving activity visibility status. Please wait...",
            onSuccess: function (data) {
            },
            subscription: this.dataSubscription
        });
    };
    FrmActivityComponent.prototype.isCoreTeam = function (teamID) {
        var team = this.oplandb.idxTeam['r' + teamID];
        if (team == undefined) {
            return true;
        }
        else {
            return team.core;
        }
    };
    FrmActivityComponent.prototype.inCalendar = function () {
        //[ngClass]="{'text-muted':noAccess(),'text-info':!noAccess()}"
        return {
            "fa-toggle-on": this.oplandb.currActivity.showInChart,
            "fa-toggle-off": !this.oplandb.currActivity.showInChart,
            "text-info": !this.noAccess(),
            "text-muted": this.noAccess(),
        };
    };
    FrmActivityComponent.prototype.noAccess = function () {
        return false;
        //return (this.oplandb.currActivity.obsolete || this.oplandb.currActivity.removed); //this.oplandb.currActivity.obsolete || this.oplandb.currActivity.removed;
        //return true;
    };
    FrmActivityComponent.prototype.calOpen = function () {
        alert('hello');
    };
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["ViewChild"])('datepicker'),
        __metadata("design:type", typeof (_a = typeof __WEBPACK_IMPORTED_MODULE_0__angular_core__["ElementRef"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_0__angular_core__["ElementRef"]) === "function" && _a || Object)
    ], FrmActivityComponent.prototype, "datepicker", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["ViewChild"])('dayPicker'),
        __metadata("design:type", typeof (_b = typeof __WEBPACK_IMPORTED_MODULE_3_ng2_date_picker__["DatePickerComponent"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_3_ng2_date_picker__["DatePickerComponent"]) === "function" && _b || Object)
    ], FrmActivityComponent.prototype, "datePicker", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", typeof (_c = typeof __WEBPACK_IMPORTED_MODULE_1__models_data_models__["c" /* dataActivity */] !== "undefined" && __WEBPACK_IMPORTED_MODULE_1__models_data_models__["c" /* dataActivity */]) === "function" && _c || Object)
    ], FrmActivityComponent.prototype, "currActivity", void 0);
    FrmActivityComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-frm-activity',
            template: __webpack_require__("../../../../../src/app/components/frm-activity/frm-activity.component.html"),
            styles: [__webpack_require__("../../../../../src/app/components/frm-activity/frm-activity.component.css")]
        }),
        __metadata("design:paramtypes", [typeof (_d = (typeof __WEBPACK_IMPORTED_MODULE_5__app_udfs__ !== "undefined" && __WEBPACK_IMPORTED_MODULE_5__app_udfs__).AppUDFService) === "function" && _d || Object, typeof (_e = typeof __WEBPACK_IMPORTED_MODULE_6__app_oplandata__["a" /* AppOplpanDataService */] !== "undefined" && __WEBPACK_IMPORTED_MODULE_6__app_oplandata__["a" /* AppOplpanDataService */]) === "function" && _e || Object])
    ], FrmActivityComponent);
    return FrmActivityComponent;
    var _a, _b, _c, _d, _e;
}());

//# sourceMappingURL=frm-activity.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/frm-chart-activities-mtiap/frm-chart-activities-mtiap.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "\r\n@media print{\r\n    .card{\r\n        height:auto;\r\n    }\r\n    .noprint{\r\n        display: none;\r\n    }\r\n    tbody{\r\n        overflow: hidden;\r\n    }\r\n    .noleftmargin{\r\n        margin-left:0px;\r\n    }\r\n\r\n\r\n}\r\n\r\n\r\n@media screen{\r\n    .card{\r\n        height:100%;\r\n    }\r\n    tbody{\r\n        overflow-y: auto;\r\n    }\r\n    .noscreen{\r\n        display:none;\r\n    }\r\n}\r\n\r\ntable th{\r\n    border:1px solid gray;\r\n}\r\n\r\n\r\ntable td{\r\n    background:white;\r\n    text-align: center;\r\n    font-size: 11px;\r\n    border:1px solid gray;\r\n\r\n    -ms-user-select: none;\r\n\r\n        user-select: none;\r\n    -moz-user-select: none;\r\n    -khtml-user-select: none;\r\n    -webkit-user-select: none;\r\n    -o-user-select: none;\r\n}\r\n\r\n.div{\r\n    -ms-user-select: none;\r\n        user-select: none;\r\n    -moz-user-select: none;\r\n    -khtml-user-select: none;\r\n    -webkit-user-select: none;\r\n    -o-user-select: none;\r\n}\r\n\r\ntable tr.sm td{\r\n    font-size:9px;\r\n}", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/frm-chart-activities-mtiap/frm-chart-activities-mtiap.component.html":
/***/ (function(module, exports) {

module.exports = "<div style=\"position:absolute;\" \n  (window:resize)=\"onResize($event)\"\n  [style.height]=\"(windowHeight-udf.bannerHeight)+'px'\"\n  [style.width]=\"(windowWidth)+'px'\"\n  [style.top]=\"(printMode ? 0 : udf.bannerHeight) + 'px'\"\n>\n\n \n<div class=\"card\" \n    [style.height]=\"printMode ? 'auto' : '100%'\"\n    [style.border]=\"printMode ? 'none' : '1px'\"\n    >\n<app-report-header *ngIf=\"printMode\" [headerText]=\"getTitle()\"></app-report-header>\n<div class=\"card-header py-0\" *ngIf=\"!printMode\">\n  <h5 class=\"my-1\">{{getTitle()}}</h5>\n</div>      \n<div class=\" card-block p-0 pb-0 my-1\" style=\"overflow-x:auto;overflow-y:hidden;\">\n  <div class=\"container-fluid p-0\" >\n    <!--div style=\"font-size:0.75em\">\n    <div id=\"testDIV\" \n      style=\"position: absolute;z-index:5;\n        background:red;color:white;\n        line-height:1.1;display:none;\"\n\n      [style.width.px]=\"testWidth\"\n      [style.font-size.em]=\"testHeight\"\n      [style.opacity]=\"0.5\"\n    >{{testString}}</div>\n  </div-->\n\n    <div *ngFor=\"let sem of sems;let idx=index;\"\n     [style.padding-left.px]=\"tableLeftMargin(sem.sem)\"  > \n      <table  *ngIf=\"(sem.sem==udf.calActivityMTIAPPage)\"\n        class=\"m-0\"\n        [ngClass]=\"{'tbc':!printMode}\"\n        \n        [style.font-size.em]=\"0.75 * udf.calActivityChartScaleMTIAP\"\n        [style.width.px]=\"chartWidth(sem.sem)+20\" border=\"1\"\n        >\n          <thead style=\"border:none; background:transparent;\">\n\n            <tr  >\n              <th [style.width.px]=\"chartRowHeaderWidth()\" class=\"text-center\">{{sem.year}}</th>\n              <th *ngFor=\"let col of moHeader(sem.sem);let cidx=index\" \n                class=\"text-center\" \n                [attr.colspan]=\"monthColSpan(sem.year,cidx)\"\n                [style.width.px]=\"monthColSpan(sem.year,cidx) * chartColWidth()\">{{col.toUpperCase()}}</th>\n                <!--th *ngIf=\"printMode\" \n                [style.visibility]=\"printMode ? 'hidden' : 'visible'\"\n                style=\"width:18px;background:#cdcdcd;\">ww</th-->\n                <th \n                [style.visibility]=\"printMode ? 'hidden' : 'visible'\"\n                style=\"width:17px;background:transparent;border:1px solid white;\">&nbsp;</th>\n              </tr>\n\n            <!--tr *ngIf=\"(sem.sem % 2)==0\" >\n              <th [style.width.px]=\"chartRowHeaderWidth()\" class=\"text-center\">{{sem.year}}</th>\n              <th class=\"text-center\"  [attr.colspan]=\"udf.weekCount(sem.year,cidx+7)\"\n              [style.width.px]=\"udf.weekCount(sem.year,cidx+7) * chartColWidth()\"\n                *ngFor=\"let col of cols2;let cidx=index\">{{col.toUpperCase()}}</th>\n                <th *ngIf=\"printMode\" \n                [style.visibility]=\"printMode ? 'hidden' : 'visible'\"\n                  style=\"width:18px;background:#cdcdcd;\">&nbsp;</th>\n            </tr-->\n\n            <tr>\n                <td [style.width.px]=\"chartRowHeaderWidth()\" class=\"text-center py-0\">Week</td>\n                <td [style.width.px]=\"chartColWidth()\" class=\"text-center py-0\" \n                    *ngFor=\"let wk of getWeeks(sem.sem)\">{{wk}}</td>\n                    <!--td  *ngIf=\"\" \n                      [style.visibility]=\"printMode ? 'hidden' : 'visible'\"\n                      style=\"width:18px;background:#cdcdcd;\">&nbsp;</td-->\n\n                    <td\n                      [style.visibility]=\"printMode ? 'hidden' : 'visible'\"\n                      style=\"width:17px;background:none;border: 1px solid white;\">&nbsp;</td>\n\n                    </tr>\n            \n          </thead>\n\n\n        <tbody *ngIf=\"false\" id=\"stiap_body\" style=\"overflow-x:hidden;position: relative;\" \n          (scroll)=\"bodyScroll($event)\"\n          \n        >\n\n            <tr *ngFor=\"let pri of sem.priorities;let pidx=index\">\n              <td [style.width.px]=\"chartRowHeaderWidth()\" class=\"text-center\">{{pri.name}}</td>\n              <td [style.width.px]=\"chartRowHeaderWidth()\" class=\"text-center\">{{pri.name}}</td>\n            </tr>\n\n          </tbody>\n\n          <tbody [style.height.px]=\"printMode ? 25 : (windowHeight-udf.bannerHeight-70-(rowHeaderHeight*udf.calActivityChartScaleMTIAP*3))\">\n          <!--totalBodyHeight(sem.sem) : (windowHeight-udf.bannerHeight-70-(rowHeaderHeight*udf.calActivityChartScaleMTIAP*3))-->\n          <!--tr>\n            <td [style.width.px]=\"chartRowHeaderWidth()\" class=\"text-center py-0\">&nbsp;</td>\n            <td [style.width.px]=\"chartColWidth()\" class=\"text-center py-0\" \n                *ngFor=\"let wk of getWeeks(sem.sem)\">{{wk}}</td><td style=\"width:18px;background:#cdcdcd;\">&nbsp;</td>\n          </tr-->\n      \n            <tr *ngFor=\"let pri of sem.priorities;let pidx=index\">\n              <td [style.width.px]=\"chartRowHeaderWidth()\" class=\"text-center\">\n                <span>{{pri.name}}</span>\n              </td>\n              \n              <td style=\"background: yellow;\" [attr.colspan]=\"getColCount(sem.sem) \" class=\"p-0\"\n                [style.width.px]=\"chartDetailsWidth(sem.sem)\" \n                style=\"position:relative;\">\n\n                <div *ngIf=\"pri.rows.length==0\" style=\"height:35px;\">&nbsp;</div>\n\n                <div *ngFor=\"let row of pri.rows;let ridx=index\" \n                  style=\"width:100%;position: relative;\"\n                  [style.height.px]=\"(pri.name=='MilestoneX' ? 25 : chartRowHeight())-2\"\n                  >\n\n                  <div  *ngFor=\"let dat1 of row.data;\"  \n                  [attr.id]=\"'testDIV_'+sem.sem+'_'+dat1.id\" \n                  style=\"position: absolute;z-index:6;\n                    background:red;color:white;\n                    line-height:1.1;top:0px;left:0px;visibility: hidden;\"\n                  [style.font-size.em]=\"1.3\"\n                  [style.width.px]=\"dat1['wdSem'+sem.sem]\"\n                  [style.opacity]=\"0.5\"\n                  >{{dat1.title}}\n                </div>\n                              \n\n\n                  <div *ngFor=\"let dat of row.data;\" \n                    style=\"position: absolute;top:0px;cursor: pointer;\n                    border-style:solid;border-width:1px;\n                    text-overflow: ellipsis;\n                    text-align:center;\n                    overflow:hidden;line-height:1.1;\"\n                    (mousemove)=\"onMouseMove($event,sem.sem,dat)\"\n                    (mouseout)=\"onMouseOut($event,sem.sem,dat)\"\n                    (dblclick)=\"oplandb.loadActivity(dat)\"\n                    [style.font-size.em]=\"chartFontScale(sem.sem,dat)\"\n                    [style.height.px]=\"pri.name=='MilestoneX' ? 25 : chartRowHeight()\"\n                    [style.left.px]=\"actLeft(sem.sem,dat,ridx)\"\n                    [style.width.px]=\"dat['wdSem'+sem.sem]\"\n                    [style.color]=\"dat.ready=='Amber' ? 'rgb(255,127,39)': dat.ready\"\n                    [style.border-color]=\"dat.isVessel ? 'rgb(33,10,255)' : 'rgb(159,150,159)'\"\n                    [attr.title]=\"actTitle(sem.sem,dat)\">\n                    {{actTitle(sem.sem,dat,true)}}\n                  </div>\n                </div>\n                <div *ngIf=\"false\" style=\"position:absolute;z-index:100;top:0;left:0px;border-top:1px solid black;width:100%;height:1px;\"></div>\n                <div style=\"position:absolute;z-index:1;top:100%;left:0px;border-bottom:2px solid black;width:100%;height:1px;\"></div>\n                <div style=\"position:absolute;z-index:1;top:0px;left:100%;border-left:1px solid black;width:1px;height:100%;\"></div>\n                <div style=\"position:absolute;z-index:1;top:0px;left:0px;border-right:1px solid black;width:1px;height:100%;\"></div>\n              </td>\n            </tr>\n          </tbody>\n      </table>\n    </div>\n  </div>\n</div> <!-- card block -->\n<div class=\"noprint card-footer p-1 text-center\"\n[style.font-size]=\"'0.8em'\">\n  {{getTitle()}}\n</div>\n\n\n</div>\n\n</div>"

/***/ }),

/***/ "../../../../../src/app/components/frm-chart-activities-mtiap/frm-chart-activities-mtiap.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return FrmChartActivitiesMtiapComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__app_globals__ = __webpack_require__("../../../../../src/app/app.globals.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__app_udfs__ = __webpack_require__("../../../../../src/app/app.udfs.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3_jquery__ = __webpack_require__("../../../../jquery/dist/jquery.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3_jquery___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_3_jquery__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_4__app_oplandata__ = __webpack_require__("../../../../../src/app/app.oplandata.ts");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};





var FrmChartActivitiesMtiapComponent = (function () {
    function FrmChartActivitiesMtiapComponent(udf, oplandb) {
        this.udf = udf;
        this.oplandb = oplandb;
        this.gbSW = 0;
        this.gbEW = 0;
        this.gbWD = 0;
        this.gbDebug = "";
        this.testString = "Test DIV";
        this.testWidth = 100;
        this.testHeight = 1.2;
        this.windowWidth = 1024;
        this.windowHeight = 768;
        this.printMode = false;
        this.rowHeaderWidth = 100;
        this.rowHeaderHeight = 20;
        this.rowHeight = 65;
        this.chartDefBaseWidth = 35; // base on 6-month period per page/2-page per year
        this.chartDefWidth = this.chartDefBaseWidth * this.udf.calActivityMTIAPPagePerYear / 2;
        this.chartPriorities = [
            { id: 7, name: "Milestones" },
            { id: 1, name: "HSSE Critical" },
            { id: 2, name: "MIE" },
            { id: 3, name: "Production Critical inc. DCP" },
            { id: 4, name: "Projects" },
            { id: 5, name: "Fabric Maintenance" },
            { id: 6, name: "Opportunity Surveys/Visits" },
        ];
        this.years = [
            { year: 2018, priorities: [
                    { id: 1,
                        rows: [
                            { activities: [
                                    {
                                        id: 1,
                                        title: "activity 1",
                                        ready: "G",
                                        start: "",
                                        end: ""
                                    },
                                    {
                                        id: 2,
                                        title: "activity 2",
                                        ready: "G",
                                        start: "",
                                        end: ""
                                    },
                                    {
                                        id: 3,
                                        title: "activity 3",
                                        ready: "R",
                                        start: "",
                                        end: ""
                                    }
                                ] },
                            { activities: [
                                    {
                                        id: 4,
                                        title: "activity 4",
                                        start: "",
                                        ready: "A",
                                        end: ""
                                    },
                                    {
                                        id: 5,
                                        title: "activity 5",
                                        start: "",
                                        ready: "A",
                                        end: ""
                                    },
                                    {
                                        id: 6,
                                        title: "activity 6",
                                        start: "",
                                        ready: "G",
                                        end: ""
                                    }
                                ] }
                        ] }
                ] },
            { year: 2019 }
        ];
        this.sems = [
            { sem: 1, year: 2018 },
            { sem: 2, year: 2018 },
            { sem: 3, year: 2019 },
            { sem: 4, year: 2019 },
        ];
        this.periods = [1, 2];
        this.chartCols = [
            "January", "February", "March",
            "April", "May", "June",
            "July", "August", "September",
            "October", "November", "December"
        ];
        this.cols1 = [
            "January", "February", "March",
            "April", "May", "June",
        ];
        this.cols2 = [
            "July", "August", "September",
            "October", "November", "December"
        ];
    }
    FrmChartActivitiesMtiapComponent.prototype.ngOnInit = function () {
        this.oplandb.clearAllSubscription();
        this.onResize(null);
        this.getSummaryData();
    };
    FrmChartActivitiesMtiapComponent.prototype.ngOnDestroy = function () {
        this.oplandb.clearSubscription(this.subsSummary);
    };
    FrmChartActivitiesMtiapComponent.prototype.onResize = function (e) {
        this.setWindowSizeVars();
    };
    FrmChartActivitiesMtiapComponent.prototype.setWindowSizeVars = function () {
        this.windowWidth = __WEBPACK_IMPORTED_MODULE_3_jquery__(window).width();
        this.windowHeight = __WEBPACK_IMPORTED_MODULE_3_jquery__(window).height();
    };
    FrmChartActivitiesMtiapComponent.prototype.periodInYear = function (per) {
        var spy = this.udf.calActivityMTIAPPagePerYear;
        return (per <= spy ? per : per - spy);
    };
    FrmChartActivitiesMtiapComponent.prototype.moHeader = function (per) {
        var spy = this.udf.calActivityMTIAPPagePerYear;
        var mos = 12 / spy;
        var yPer = (per <= spy ? per : per - spy);
        var arr = [];
        var sIdx = (yPer - 1) * mos;
        return this.chartCols.slice(sIdx, sIdx + mos);
    };
    FrmChartActivitiesMtiapComponent.prototype.chartFontScale = function (sem, dat) {
        var div = __WEBPACK_IMPORTED_MODULE_3_jquery__('#testDIV_' + sem + '_' + dat.id);
        var ht = this.chartRowHeight();
        var sc = div.height() / ht;
        if (sc < 1) {
            return 1.3;
        }
        else {
            return Math.max(1.3 / sc, 0.82);
        }
        //return 1;
        /*if(this.udf.calActivityChartScaleMTIAP > 1){
          return Math.max(1,this.udf.calActivityChartScaleMTIAP*0.85);  // apply font scale ceiling to prevent super large text
        }else{
          return Math.max(this.udf.calActivityChartScaleMTIAP,0.7);
        }*/
        /*let actFontSize:number=dat['fsizeSem'+sem];
        if(actFontSize!=-1){
    
          return actFontSize;
    
        }else{*/
        /*     let ht:number=this.chartRowHeight();
             let div=$('#testDIV_'+sem+'_'+dat.id);
       
             if(div.height()>ht){
               //this.setFontSize(sem,dat);
               return 1;
             }else{
               return 1.3;
             }
         */
        /*   setTimeout(()=>{
             this.setFontSize(sem,dat);
           },10);
           return 1.3;
         }*/
        //this.testWidth=this.actWidth(sem,dat);
        //this.testString=dat.title;
        //this.testHeight=1;
        /*let div=$('#testDIV_'+pidx+'_'+ridx);
        if(div.height()>ht){
          this.testHeight=1.1;
          if(div.height()>ht){
            this.testHeight=1;
            if(div.height()>ht){
              this.testHeight=0.9;
              if(div.height()>ht){
                this.testHeight=0.8;
              }
            }
          }
        }
        if(div.length){
          //console.log(div);
        }*/
        //this.testHeight = 1;
        //return this.testHeight;
    };
    FrmChartActivitiesMtiapComponent.prototype.chartSeparatorHeight = function () {
        return 3 * this.udf.calActivityChartScaleMTIAP;
    };
    FrmChartActivitiesMtiapComponent.prototype.chartColumnHeaderHeight = function () {
        return 48 * this.udf.calActivityChartScaleMTIAP;
    };
    FrmChartActivitiesMtiapComponent.prototype.chartColWidth = function () {
        return this.chartDefWidth * this.udf.calActivityChartScaleMTIAP;
    };
    FrmChartActivitiesMtiapComponent.prototype.chartRowHeight = function () {
        return this.rowHeight * this.udf.calActivityChartScaleMTIAP;
    };
    FrmChartActivitiesMtiapComponent.prototype.chartRowHeaderWidth = function () {
        return this.rowHeaderWidth * this.udf.calActivityChartScaleMTIAP;
    };
    FrmChartActivitiesMtiapComponent.prototype.chartDetailsWidth = function (sem) {
        return this.getColCount(sem) * (this.chartDefWidth) * this.udf.calActivityChartScaleMTIAP;
    };
    FrmChartActivitiesMtiapComponent.prototype.chartWidth = function (sem) {
        return this.rowHeaderWidth * this.udf.calActivityChartScaleMTIAP + this.chartDetailsWidth(sem);
    };
    FrmChartActivitiesMtiapComponent.prototype.tableLeftMargin = function (sem) {
        if (this.printMode) {
            return 0;
        }
        else {
            return Math.max(0, (this.windowWidth - this.chartDetailsWidth(sem) - this.chartRowHeaderWidth()) / 2);
        }
    };
    FrmChartActivitiesMtiapComponent.prototype.totalBodyHeight = function (sem) {
        var _this = this;
        var ht = 0;
        this.sems[sem]['priorities'].forEach(function (pri) {
            ht += (1 + pri.rows.length) * _this.chartRowHeight();
        });
        return ht + (this.rowHeaderHeight * this.udf.calActivityChartScaleMTIAP) * 5;
    };
    FrmChartActivitiesMtiapComponent.prototype.getBosyStyle = function () {
        //return [style.height.px] = "printMode ? totalBodyHeight(sem.sem) : (windowHeight-udf.bannerHeight-70-(rowHeaderHeight*udf.calActivityChartScaleMTIAP*3))"
        /*if(this.printMode){
          return {}
        }else{
          return {
            "height": (this.windowHeight-this.udf.bannerHeight-70-
                (this.rowHeaderHeight*this.udf.calActivityChartScaleMTIAP*3))
          }
        }*/
        return { 'height': 500 };
    };
    FrmChartActivitiesMtiapComponent.prototype.actLeft = function (sem, dat) {
        var ret = (dat['sem' + sem + 'SW'] - 1) * (this.chartColWidth()) - 1;
        return ret;
    };
    FrmChartActivitiesMtiapComponent.prototype.actWidth = function (sem, dat) {
        var ret = (dat['sem' + sem + 'EW'] - dat['sem' + sem + 'SW'] + 1) * (this.chartColWidth());
        return ret;
    };
    FrmChartActivitiesMtiapComponent.prototype.actTitle = function (sem, dat, titleOnly) {
        if (titleOnly == undefined)
            titleOnly = false;
        var title = "[" + (__WEBPACK_IMPORTED_MODULE_1__app_globals__["daysDiff"](dat.startDate, dat.endDate) + 1) + "d] " + dat.title;
        if (titleOnly) {
            return title;
        }
        else {
            return title +
                '\nStart: ' + __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](dat.startDate, true) +
                '\nEnd: ' + __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](dat.endDate, true) +
                '\nId: ' + dat.id +
                this.gbDebug;
        }
    };
    //getColCount(sem.sem) * (chartColWidth()+1) + chartRowHeaderWidth()
    FrmChartActivitiesMtiapComponent.prototype.getTitle = function () {
        //let ret:string="MTIAP View for " + G.monthNames[this.udf.calActivityMonth] + ' ' + this.udf.calActivityYear;
        var ret = this.udf.calActivitySiteMTIAP + " MTIAP for Years " + this.udf.calActivityMTIAPYear + "-" + this.nextYear();
        return ret;
    };
    FrmChartActivitiesMtiapComponent.prototype.nextYear = function () {
        return parseInt(String(this.udf.calActivityMTIAPYear)) + 1;
    };
    FrmChartActivitiesMtiapComponent.prototype.getColCount = function (sem) {
        var yr = this.udf.calActivityMTIAPYear;
        var ret = 0;
        var spy = this.udf.calActivityMTIAPPagePerYear;
        var mos = 12 / spy;
        var smo = mos * (sem - 1) + 1;
        for (var i = 0; i < mos; i++) {
            ret += this.udf.weekCount(yr, smo + i);
        }
        return ret;
        /*switch(sem){
          case 1:
            return this.udf.weekCount(yr,1) + this.udf.weekCount(yr,2) + this.udf.weekCount(yr,3) +
                   this.udf.weekCount(yr,4) + this.udf.weekCount(yr,5) + this.udf.weekCount(yr,6)
          case 2:
            return this.udf.weekCount(yr,7) + this.udf.weekCount(yr,8) + this.udf.weekCount(yr,9) +
                   this.udf.weekCount(yr,10) + this.udf.weekCount(yr,11) + this.udf.weekCount(yr,12)
          case 3:
            return this.udf.weekCount(yr+1,1) + this.udf.weekCount(yr+1,2) + this.udf.weekCount(yr+1,3) +
                   this.udf.weekCount(yr+1,4) + this.udf.weekCount(yr+1,5) + this.udf.weekCount(yr+1,6)
          case 4:
            return this.udf.weekCount(yr+1,7) + this.udf.weekCount(yr+1,8) + this.udf.weekCount(yr+1,9) +
            this.udf.weekCount(yr+1,10) + this.udf.weekCount(yr+1,11) + this.udf.weekCount(yr+1,12)
        }*/
    };
    FrmChartActivitiesMtiapComponent.prototype.getWeeks = function (sem) {
        var ret = [];
        var i;
        var j;
        var yr = this.udf.calActivityMTIAPYear;
        var pg = this.udf.calActivityMTIAPPage;
        var spy = this.udf.calActivityMTIAPPagePerYear;
        var y2 = (yr > this.udf.calActivityMTIAPYear);
        var per = 12 / spy;
        var off = (pg - 1 - (pg <= spy ? 0 : spy)) * per;
        var mos = 12 / this.udf.calActivityMTIAPPagePerYear;
        var semObj = this.sems[sem - 1];
        for (i = 1; i <= mos; i++) {
            var mo = i + off;
            var wks = this.udf.weekCount(semObj.year, mo);
            for (j = 1; j <= wks; j++) {
                ret.push(String(j));
            }
        }
        return ret;
    };
    FrmChartActivitiesMtiapComponent.prototype.getWeeksBeforeDate = function (d) {
        // this returns the total number of weeks starting from
        // the first month of the year to the month prior to the
        // month of date specified
        var ret = 0;
        var yr = d.getFullYear();
        var mo = d.getMonth();
        var wc = this.udf.weekCount;
        var imo;
        this.udf.weekCount(yr, mo);
        for (imo = 0; imo < mo; imo++) {
            ret += wc(yr, imo + 1); // wc function is base 1 for month
        }
        return ret;
    };
    FrmChartActivitiesMtiapComponent.prototype.getSummaryData = function () {
        var _this = this;
        var y2 = parseInt(String(this.udf.calActivityMTIAPYear)) + 1;
        this.subsSummary = this.oplandb.processData({
            method: "getCalendarMTIAPSummary",
            data: {
                site: this.udf.calActivitySiteMTIAP,
                startDate: __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](new Date(this.udf.calActivityMTIAPYear, 0, 1)),
                endDate: __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](new Date(y2, 11, 31)),
                months: 12 / this.udf.calActivityMTIAPPagePerYear
            },
            isPosting: false,
            processMessage: "Loading activity MTIAP view data. Please wait...",
            subscription: this.subsSummary,
            onSuccess: function (data) {
                //if(this.calActivities!=undefined)this.calActivities.initVars(data);
                _this.initVars(data);
                _this.oplandb.isLoading = false;
            }
        });
    };
    FrmChartActivitiesMtiapComponent.prototype.onMouseMove = function (e, sem, dat) {
        //let div=$('#testDIV_'+sem+'_'+dat.id);
        //div.css('visibility','visible');
    };
    FrmChartActivitiesMtiapComponent.prototype.onMouseOut = function (e, sem, dat) {
        //let div=$('#testDIV_'+sem+'_'+dat.id);
        //div.css('visibility','hidden');
    };
    FrmChartActivitiesMtiapComponent.prototype.monthColSpan = function (yr, idx) {
        var pg = this.udf.calActivityMTIAPPage;
        var spy = this.udf.calActivityMTIAPPagePerYear;
        var y2 = (yr > this.udf.calActivityMTIAPYear);
        var per = 12 / spy;
        var off = (pg - 1 - (pg <= spy ? 0 : spy)) * per;
        var mo = idx + 1 + off;
        return this.udf.weekCount(yr, mo);
    };
    FrmChartActivitiesMtiapComponent.prototype.initVars = function (data) {
        var _this = this;
        this.chartDefWidth = this.chartDefBaseWidth * this.udf.calActivityMTIAPPagePerYear / 2;
        this.sems = this.udf.MTIAPPeriods();
        var spy = this.udf.calActivityMTIAPPagePerYear; // sem per year, or period per year
        // initialize sem/period properties
        this.sems.forEach(function (s) {
            // initialize priorities
            s.priorities = [];
            // calculate and assign year
            s.year = parseInt(String(_this.udf.calActivityMTIAPYear)) + (s.sem <= spy ? 0 : 1);
        });
        var acts = [];
        data.ACTS.forEach(function (e) {
            var row = new Object();
            var incSwIdx = 6; // start index of period inclusion switches
            row.id = Number(e[0]);
            row.priority = parseInt(e[1]);
            for (var per = 1; per <= _this.sems.length; per++) {
                row['plotRow' + per] = -1; // row assignment, -1 if not yet assigned
                row['wdSem' + per] = -1; // actual width of activity in the period
                row['sem' + per + 'SW'] = -1; // activity start week column in the period
                row['sem' + per + 'EW'] = -1; // activity end week column in the period
            }
            row.title = e[2];
            row.startDate = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](e[3]);
            row.endDate = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](e[4]);
            row.ready = e[5];
            row.sem1 = e[6];
            row.sem2 = e[7];
            row.sem3 = e[8];
            row.sem4 = e[9];
            row.sem5 = e[10];
            row.sem6 = e[11];
            row.sem7 = e[12];
            row.sem8 = e[13];
            row.sem9 = e[14];
            row.sem10 = e[15];
            row.sem11 = e[16];
            row.sem12 = e[17];
            row.startWeekOfMonth = e[18];
            row.endWeekOfMonth = e[19];
            row.isVessel = e[20];
            row.bndSD = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](e[21]);
            row.bndED = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](e[22]);
            row.sem13 = e[23];
            row.sem14 = e[24];
            row.sem15 = e[25];
            row.sem16 = e[26];
            row.sem17 = e[27];
            row.sem18 = e[28];
            row.sem19 = e[29];
            row.sem20 = e[30];
            row.sem21 = e[31];
            row.sem22 = e[32];
            row.sem23 = e[33];
            row.sem24 = e[34];
            _this.sems.forEach(function (per) {
                var perIdx = per.sem;
                if (row['sem' + perIdx]) {
                    var onPrev = false;
                    var onNext = false;
                    var yr = _this.udf.calActivityMTIAPYear + (perIdx <= spy ? 0 : 1);
                    var off = 0;
                    var prevPerOnYear = perIdx - 1 - (perIdx <= spy ? 0 : spy);
                    if (prevPerOnYear != 0) {
                        // iterate through all the previous period and get the total weeks per period
                        for (var prevPer = 1; prevPer <= prevPerOnYear; prevPer++) {
                            // get column count of previous period within the year of perIdx
                            off += _this.getColCount(prevPer + (perIdx <= spy ? 0 : spy));
                        }
                    }
                    if (perIdx > 1)
                        onPrev = row['sem' + (perIdx - 1)];
                    if (perIdx < _this.sems.length)
                        onNext = row['sem' + (perIdx + 1)];
                    if (row.id == 1020) {
                        //onPrev=true;
                        console.log("Act1020:", perIdx, onPrev, onNext, row['sem' + (perIdx)]);
                        console.log(perIdx, _this.sems.length);
                        console.log("row.sem4:", row.sem3);
                    }
                    row['sem' + (perIdx) + 'SW'] = onPrev ? 1 : _this.getWeeksBeforeDate(row.bndSD) - off + row.startWeekOfMonth;
                    row['sem' + (perIdx) + 'EW'] = onNext ? _this.getColCount(perIdx) : _this.getWeeksBeforeDate(row.bndED) - off + row.endWeekOfMonth;
                    row['wdSem' + perIdx] = _this.actWidth(perIdx, row);
                } // if on current period!!!
            });
            acts.push(row); // append row to activities collection
        }); // for eacj data.ACTS
        // start positioning activities without overlapping ....
        this.sems.forEach(function (s) {
            // sem ordinal index
            var si = s.sem;
            // get all activities covered within a sem/period
            var sem = acts.filter(function (elem) { return elem['sem' + si]; });
            // iterate through activity types/priorities
            data.TYPES.forEach(function (e) {
                if (e[0] != 0) {
                    var pri_1 = e[0]; // priority key field
                    // get all activities that fall under the current type/priority
                    var priArr = sem.filter(function (elem) { return elem.priority == pri_1; });
                    var ctr = 0;
                    var curRow_1 = 0; // initialize plot row number
                    // get all remaining elements that are not yet positioned
                    var remRows = priArr.filter(function (elem) { return elem['plotRow' + si] == -1; });
                    var _loop_1 = function () {
                        var endCol = 0; // initialize for each row
                        remRows.forEach(function (r) {
                            if (r['sem' + si + 'SW'] > endCol) {
                                r['plotRow' + si] = curRow_1; // assign row number to the activity 
                                endCol = r['sem' + si + 'EW']; // mark the end column of the activity as the new reference end column
                            }
                        });
                        // reinitilize remaining activity array 
                        remRows = remRows.filter(function (elem) { return elem['plotRow' + si] == -1; });
                        curRow_1++; // increment row number
                        ctr++; // increment row ceiling counter    
                    };
                    // iterate through activities while number of rows is under 100 (arbitrary number which is assumed to cover all activities within the period)
                    while (remRows.length != 0 && ctr < 100) {
                        _loop_1();
                    }
                    // arrange activities in a sem/period->type/priority by row assignment
                    __WEBPACK_IMPORTED_MODULE_1__app_globals__["sortBy"](priArr, 'plotRow' + si);
                    // create new element to the sem/period priorities array property
                    s.priorities.push({
                        id: e[0],
                        name: e[1],
                        rows: __WEBPACK_IMPORTED_MODULE_1__app_globals__["groupBy"](priArr, 'plotRow' + si) // group elements by assigned row number
                    });
                } // is priority id is not 0
            }); // for each priority/type      
        }); // for each sem
    };
    FrmChartActivitiesMtiapComponent.prototype.onBeforePrint = function (event) {
        // 1052.16:0.96 px for A4
        this.printMode = true;
        //this.oldScale=this.udf.calActivityChartScale;
        this.udf.calActivityChartScale = 1;
    };
    FrmChartActivitiesMtiapComponent.prototype.onAfterPrint = function (event) {
        this.printMode = false;
        //this.udf.calActivityChartScale=this.oldScale;
    };
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["HostListener"])('window:beforeprint', ['$event']),
        __metadata("design:type", Function),
        __metadata("design:paramtypes", [Object]),
        __metadata("design:returntype", void 0)
    ], FrmChartActivitiesMtiapComponent.prototype, "onBeforePrint", null);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["HostListener"])('window:afterprint', ['$event']),
        __metadata("design:type", Function),
        __metadata("design:paramtypes", [Object]),
        __metadata("design:returntype", void 0)
    ], FrmChartActivitiesMtiapComponent.prototype, "onAfterPrint", null);
    FrmChartActivitiesMtiapComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-frm-chart-activities-mtiap',
            template: __webpack_require__("../../../../../src/app/components/frm-chart-activities-mtiap/frm-chart-activities-mtiap.component.html"),
            styles: [__webpack_require__("../../../../../src/app/components/frm-chart-activities-mtiap/frm-chart-activities-mtiap.component.css")]
        }),
        __metadata("design:paramtypes", [typeof (_a = (typeof __WEBPACK_IMPORTED_MODULE_2__app_udfs__ !== "undefined" && __WEBPACK_IMPORTED_MODULE_2__app_udfs__).AppUDFService) === "function" && _a || Object, typeof (_b = typeof __WEBPACK_IMPORTED_MODULE_4__app_oplandata__["a" /* AppOplpanDataService */] !== "undefined" && __WEBPACK_IMPORTED_MODULE_4__app_oplandata__["a" /* AppOplpanDataService */]) === "function" && _b || Object])
    ], FrmChartActivitiesMtiapComponent);
    return FrmChartActivitiesMtiapComponent;
    var _a, _b;
}());

//# sourceMappingURL=frm-chart-activities-mtiap.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/frm-chart-activities/frm-chart-activities.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "\r\n@media print{\r\n    .card{\r\n        height:auto;\r\n    }\r\n    .noprint{\r\n        display: none;\r\n    }\r\n    tbody{\r\n        overflow: hidden;\r\n    }\r\n    .noleftmargin{\r\n        margin-left:0px;\r\n    }\r\n\r\n\r\n}\r\n\r\n\r\n@media screen{\r\n    .card{\r\n        height:100%;\r\n    }\r\n    tbody{\r\n        overflow-y: auto;\r\n    }\r\n    .noscreen{\r\n        display:none;\r\n    }\r\n}\r\n\r\ntable th{\r\n    border:1px solid gray;\r\n}\r\n\r\n\r\ntable td{\r\n    background:white;\r\n    text-align: center;\r\n    font-size: 11px;\r\n    border:1px solid gray;\r\n\r\n    -ms-user-select: none;\r\n\r\n        user-select: none;\r\n    -moz-user-select: none;\r\n    -khtml-user-select: none;\r\n    -webkit-user-select: none;\r\n    -o-user-select: none;\r\n}\r\n\r\n.div{\r\n    -ms-user-select: none;\r\n        user-select: none;\r\n    -moz-user-select: none;\r\n    -khtml-user-select: none;\r\n    -webkit-user-select: none;\r\n    -o-user-select: none;\r\n}\r\n\r\ntable tr.sm td{\r\n    font-size:9px;\r\n}\r\n\r\n.blink{\r\n    font-size: 1em;\r\n    color: rgb(217,83,66);\r\n    \r\n    -webkit-animation: blink 1s infinite;\r\n    \r\n            animation: blink 1s infinite;\r\n }\r\n \r\n @-webkit-keyframes blink{\r\n    0%{opacity: 1;}\r\n    75%{opacity: 1;}\r\n    76%{ opacity: 0;}\r\n    100%{opacity: 0;}\r\n }\r\n \r\n @keyframes blink{\r\n    0%{opacity: 1;}\r\n    75%{opacity: 1;}\r\n    76%{ opacity: 0;}\r\n    100%{opacity: 0;}\r\n }", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/frm-chart-activities/frm-chart-activities.component.html":
/***/ (function(module, exports) {

module.exports = "<div [innerHTML]=\"debugText\" id=\"debugDiv\" \n  *ngIf=\"debugOn\"\n  style=\"font-size:9pt;padding-left:5px;padding-right:5px;background:beige;position: absolute;z-index:100;top:0px;left:0px;height:65px;width:100%;overflow:auto\"\n  >\n</div>\n\n<div style=\"position:absolute;z-index:50;left:740px;\"\n  [style.top.px]=\"(udf.bannerHeight-udf.bannerMenuHeight+6)\"\n>\n\n  <button \n  style=\"cursor:pointer;text-decoration: none;font-size:1.1em;\" \n  class=\"btn btn-link p-1 noprint \"\n  onfocus=\"this.blur()\"\n  data-toggle=\"tooltip\" \n  (click)=\"toggleEditMode(dragEditing ? 'save' : 'initialize')\"\n  [title]=\"(dragEditing ? 'Save changes and exit drag & drop editing mode...' : 'Turn on drag & drop editing mode...')\"\n  >\n    <i class=\"fa\" [ngClass]=\"{'fa-save':dragEditing,'text-success':dragEditing,'fa-pencil':!dragEditing}\"\n      aria-hidden=\"true\"></i>\n    <span *ngIf=\"!dragEditing\">...</span>\n  </button>\n\n  <button class=\"btn btn-link p-1 text-danger\" *ngIf=\"dragEditing\" \n    (click)=\"toggleEditMode('cancel',true)\"\n    style=\"cursor: pointer;\"\n    title=\"Cancel changes and exit drag & drop editing mode...\"\n    >\n    <i class=\"fa fa-ban\"  aria-hidden=\"true\"></i>\n  </button>\n  <span class=\"blink\" *ngIf=\"dragEditing\" style=\"cursor: default;\"> Drag & Drop is ON !</span>\n\n\n</div>\n\n<div style=\"position:absolute;\" (window:resize)=\"onResize($event)\"\n[style.height]=\"(windowHeight-udf.bannerHeight)+'px'\"\n[style.width]=\"(windowWidth)+'px'\"\n[style.top]=\"(printMode ? 0 : udf.bannerHeight) + 'px'\"\n  >\n\n \n    <div class=\"card\" \n      [style.height]=\"printMode ? 'auto' : '100%'\"\n      [style.border]=\"printMode ? 'auto' : '1px'\"\n      >\n      <app-report-header *ngIf=\"printMode\" [headerText]=\"getTitle()\"></app-report-header>\n      <div class=\"card-header py-0\" *ngIf=\"!printMode\">\n        <h5 class=\"my-1\">{{getTitle()}}</h5>\n      </div>      \n      <div class=\" card-block p-2 my-1\" style=\"overflow-x:auto;overflow-y:hidden;\" onselectstart=\"return false;\" >\n        \n\n          <table class=\"tbc\" border=\"0\" cellspacing=\"2\" cellpadding=\"1\" \n            [style.margin-left]=\"printMode ? '0px' : (chartWidth() < windowWidth ? (windowWidth-chartWidth())/2 :0)+'px'\"\n            [style.width]=\"chartWidth()+'px'\"\n            >\n          <thead>\n            <!--tr>\n                <td [attr.colspan]=\"chartCols.length + 1\" style=\"font-weight:bold;background:brown;color:white;\">{{getTitle()}}</td>\n            </tr-->            \n            \n            <tr >\n                <td  [style.height]=\"(rowHeaderHeight*udf.calActivityChartScale)+'px'\"\n                [style.width]=\"chartRowHeaderWidth() + 'px'\"><span *ngIf=\"hovObjSrc\">&nbsp;</span>&nbsp;</td>\n                  <td [style.height]=\"(rowHeaderHeight*udf.calActivityChartScale)+'px'\"\n                  *ngFor=\"let wk of getWeeks();let widx=index;\" \n                    [style.font-size]=\"(0.75 * udf.calActivityChartScale) + 'em'\"\n                    [attr.colspan]=\"wk\" >{{(wk>1 ?'WEEK ' : 'WK ')+(widx+1)}}</td>\n            </tr>\n            <tr ><td [style.width]=\"chartRowHeaderWidth() + 'px'\" >&nbsp;</td>\n              <td [style.font-size]=\"(0.6 * udf.calActivityChartScale) + 'em'\"\n                [style.height]=\"(rowHeaderHeight*udf.calActivityChartScale)+'px'\"\n                [style.width.px]=\"chartColWidth()\" \n                *ngFor=\"let col of chartCols;trackBy:oplandb.trackByItem;let cidx=index;\" \n                [title] = \"oplandb.isNoFlight(getColDate(cidx+1)) ? 'No flight available.' : ''\"\n                [style.background]=\"oplandb.isNoFlight(getColDate(cidx+1)) ? '#c2e7ed' : ''\"\n                >{{cidx+1}}</td>\n            </tr>\n            <tr><td [attr.colspan]=\"chartCols.length + 1\" \n                  [style.height]=\"chartSeparatorHeight()+'px'\" \n                  style=\"font-size:1px;\" >&nbsp;</td></tr>\n            <tr *ngIf=\"udf.calActivitySite=='SWP'\" ><td [style.width]=\"chartRowHeaderWidth() + 'px'\"\n                    [style.height]=\"(rowHeaderHeight*udf.calActivityChartScale)+'px'\"\n                    [style.font-size]=\"(0.75 * udf.calActivityChartScale) + 'em'\">SWP TOTAL POB</td>\n                <td [style.width]=\"chartColWidth() + 'px'\" \n                    [style.font-size]=\"(0.75 * udf.calActivityChartScale) + 'em'\"\n                  *ngFor=\"let col of chartCols;trackBy:oplandb.trackByItem;\" >{{col.POB}}</td>\n            </tr>\n            \n            <!--tr class=\"sm\" ><td>DCP OPERATION</td><td *ngFor=\"let col of chartCols\" >&nbsp;</td></tr>\n            <tr class=\"sm\" ><td>FLOWLINE CONFIG</td><td *ngFor=\"let col of chartCols\" >&nbsp;</td></tr>\n            <tr class=\"sm\" ><td>LOT / WELL TEST</td><td *ngFor=\"let col of chartCols\" >&nbsp;</td></tr>\n            <tr class=\"sm\" ><td>CONDY LOADING</td><td *ngFor=\"let col of chartCols\" >&nbsp;</td></tr-->\n          </thead>\n          <tbody id=\"stiap_body\" style=\"overflow-x:hidden;position: relative;\"\n            (scroll)=\"bodyScroll($event)\"\n            [style.height]=\"printMode ? (totalBodyHeight()) +'px' : (windowHeight-udf.bannerHeight-85-(rowHeaderHeight*udf.calActivityChartScale*3 + chartSeparatorHeight()))+'px'\"\n          >\n\n          \n            <div id=\"dragHandle\" *ngIf=\"dragEditing\"\n              class=\"d-flex justify-content-between p-0\" \n              style=\"color:white;z-index:73;background: blue;opacity: 0.2;position: absolute;\"\n              onselectstart=\"return false;\" \n              [style.cursor]=\"curMove()\"\n              [style.visibility]=\"hovObjSrc ? 'visible' : 'hidden'\"\n\n              [style.top.px]=\"handTop()\"\n              [style.left.px]=\"handLeft()\"\n              \n              [style.height.px]=\"handHeight()\"\n              [style.width.px]=\"handWidth()\"\n              title=\"Make activity earlier or later...\"\n              (mousedown)=\"mouseEvent($event)\"\n            >\n              <div id=\"dragLeft\" (mousedown)=\"mouseEvent($event)\"\n                title=\"Make activity earlier and longer...\"\n                [style.cursor]=\"curLeft()\"\n                onselectstart=\"return false;\" \n                style=\"background:black;color:white;width:8px;\">&nbsp;</div>\n              <div id=\"dragRight\" (mousedown)=\"mouseEvent($event)\"\n                title=\"Make activity longer...\"\n                [style.cursor]=\"curLeft()\"\n                onselectstart=\"return false;\" \n                style=\"background:black;color:white;width:8px;\">&nbsp;</div>\n                \n            </div>\n\n            <div id=\"dragMask\" *ngIf=\"hovObj\"\n              style=\"z-index:75;position:absolute;background:black;white-space:normal;opacity: 0.2;\"\n              [style.margin-top.px]=\"maskTop()\"\n              [style.margin-left.px]=\"chartRowHeaderWidth()\"\n              [style.width.px]=\"maskWidth()\"\n              [style.height.px]=\"maskHeight()\"\n              [style.cursor]=\"!dragValid() ? 'no-drop' : (dragMode=='move' ? 'move' : (dragMode=='e-resize' || dragMode=='w-resize' ? 'ew-resize' : 'default'))\"\n\n              onselectstart=\"return false;\" \n\n              (mousemove)=\"mouseEvent($event)\"\n              (mouseup)=\"mouseEvent($event)\"\n\n              >\n              &nbsp;\n            </div>\n\n            <tr *ngFor=\"let pri of chartPriorities;let pidx=index;\" style=\"border-bottom:1px solid gray;\" >\n              <td [style.width]=\"chartRowHeaderWidth() + 'px'\" \n                style=\"border-right:2px solid gray;border-bottom:1px solid gray;\"\n                [style.font-size]=\"(0.8 * udf.calActivityChartScale) + 'em'\"\n                [style.height]=\"((pri.id==777 ? 0.5 : 1) * chartRowHeight() + 5) + 'px'\">{{(pri.id==7?'': pidx +'. ') + pri.name}}</td>\n              <!--[style.background]=\"(pidx % 2 ? 'white' : udf.cardHeaderBackground )\"-->\n              <td \n                [attr.colspan]=\"chartCols.length\"\n                style=\"padding:0px;position:relative;border:1px solid gray;border-bottom:3px solid gray;border-right:3px solid gray;\"\n                [style.height]=\"(pri['actRows'].length *  (pri.id==777 ? 0.5 : 1) * chartRowHeight() + 5) + 'px'\"\n                [style.width]=\"(chartColWidth()*chartCols.length) + 'px'\"\n              >\n\n              <div *ngIf=\"pri['actRows'].length==0\"\n                style=\"position:absolute;background: transparent; border:none;\"\n                [style.height]=\"chartColumnHeaderHeight()+'px'\"\n                [style.top]=\"(1+(pri.id==777 ? 0.5 : 1) * idx * chartRowHeight())+'px'\"\n                [style.left]=\"0 + 'px'\"\n                [style.width]=\"(chartColWidth()*chartCols.length) + 'px'\"\n                onselectstart=\"return false;\" \n              >\n              <div>&nbsp;</div>\n              <div *ngIf=\"printMode\"  \n                [style.height.px] = \"(pri.id==777 ? 0.5 : 1) * chartRowHeight() + 10\" \n                style=\"position:absolute;z-index:200;top:-30px;left:-1px;border-right:1px solid gray;width:1px;\"></div>\n\n              </div>\n              \n              <div *ngFor=\"let row of pri['actRows'];let idx=index;\"\n                style=\"position:absolute;background: transparent; border:none;\"\n                [style.height]=\"chartColumnHeaderHeight()+'px'\"\n                [style.top]=\"(1+(pri.id==777 ? 0.5 : 1) * idx * chartRowHeight())+'px'\"\n                [style.left]=\"0 + 'px'\"\n                [style.width]=\"(chartColWidth()*chartCols.length) + 'px'\"\n                onselectstart=\"return false;\" \n              >\n                <div *ngFor=\"let act of row;trackBy:oplandb.trackByItem;let aidx=index;\"\n                  style=\"background:#c0c0c0;position:absolute;text-overflow: ellipsis;overflow: hidden;\"\n                  [style.left]=\"(activityStart(act)*udf.calActivityChartScale) + 'px'\"\n                  [style.width]=\"((activityWidth(act)) * udf.calActivityChartScale) + 'px'\"\n                  [style.height]=\"((pri.id==777 ? 0.5 : 1) * chartRowHeight())+'px'\"\n                  onselectstart=\"return false;\" \n                >\n\n                <!--Actual visible cell displaying activity title-->\n                <div  onselectstart=\"return false;\" style=\"height:100%;cursor:pointer;border:none;\"\n                  (dblclick)=\"!dragEditing ? oplandb.loadActivity(act) : null\"\n                  [ngStyle]=\"readyColor(act)\"\n                  [style.margin]=\"'1px'\"\n                  [style.font-size]=\"(0.9 * chartFontScale()) + 'em'\" \n                  [title]=\"hoverText(act)\"\n                  (mouseenter)=\"mouseEvent($event,act,pri,row,idx)\"\n                  (mouseleave)=\"mouseEvent($event)\"\n                  \n                  (mouseup)=\"mouseEvent($event)\">\n                  {{act.title}}\n                </div>\n                </div>\n              </div>\n              \n              <div *ngIf=\"printMode && pidx==0\" \n                [style.top.px]=\"-1\"\n                style=\"position:absolute;z-index:100;left:0px;border-bottom:1px solid gray;width:100%;height:1px;\"></div>\n              <div *ngIf=\"printMode || udf.isIEOnly\" \n                [style.top.px]=\"((pri['actRows'].length?pri['actRows'].length:1) *  (pri.id==777 ? 0.5 : 1) * chartRowHeight() + 3)\"\n                style=\"position:absolute;z-index:100;left:0px;border-bottom:1px solid gray;width:100%;height:1px;\"></div>\n              <div *ngIf=\"printMode || udf.isIEOnly\" style=\"position:absolute;z-index:100;top:0px;left:100%;border-left:1px solid black;width:1px;height:100%;\"></div>\n              <div *ngIf=\"printMode\"  \n                [style.height.px] = \"((pri['actRows'].length?pri['actRows'].length:1) *  (pri.id==777 ? 0.5 : 1) * chartRowHeight() + 5)\" \n                style=\"position:absolute;z-index:200;top:0px;left:-1px;border-right:1px solid gray;width:1px;\"></div>\n          \n            </td>\n            </tr>\n            <tr><td [attr.colspan]=\"chartCols.length + 1\" \n              [style.height]=\"chartSeparatorHeight()+'px'\" style=\"font-size:1px;\" >&nbsp;</td></tr>\n          </tbody>\n      </table> \n              \n          </div>\n          <div class=\"noprint card-footer p-1 text-center\"\n          [style.font-size]=\"'0.8em'\">\n            \n            {{getTitle()}}\n          </div>\n   \n    \n  </div>\n</div>"

/***/ }),

/***/ "../../../../../src/app/components/frm-chart-activities/frm-chart-activities.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return FrmChartActivitiesComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__app_globals__ = __webpack_require__("../../../../../src/app/app.globals.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__app_udfs__ = __webpack_require__("../../../../../src/app/app.udfs.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3_jquery__ = __webpack_require__("../../../../jquery/dist/jquery.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3_jquery___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_3_jquery__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_4__app_oplandata__ = __webpack_require__("../../../../../src/app/app.oplandata.ts");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};





var FrmChartActivitiesComponent = (function () {
    function FrmChartActivitiesComponent(udf, oplandb) {
        this.udf = udf;
        this.oplandb = oplandb;
        this.debugOn = false;
        this.dragEditing = false;
        this.dragMode = "";
        this.dragHandleId = "";
        this.debugText = "test";
        this.noMove = false;
        this.noLeft = false;
        this.noRight = false;
        this.hovObj = null;
        this.hovAct = null;
        this.hovPri = null;
        this.hovRow = null;
        this.hovRowIdx = -1;
        this.hovOldStart = "";
        this.hovOldEnd = "";
        this.hovNewStart = "";
        this.hovNewEnd = "";
        this.hovNewRow = -1;
        this.hovNewRowIdx = -1;
        this.hovObjSrc = null;
        this.windowWidth = 1024;
        this.windowHeight = 768;
        this.rowHeaderWidth = 100;
        this.rowHeaderHeight = 20;
        this.rowHeight = 50;
        this.chartDefWidth = 32;
        this.chartWeeks = [];
        this.chartCols = [
            { POB: 56 },
            { POB: 56 },
            { POB: 56 },
            { POB: 58 },
            { POB: 58 },
            { POB: 57 },
            { POB: 57 },
            { POB: 56 },
            { POB: 56 },
            { POB: 56 },
            { POB: 57 },
            { POB: 58 },
            { POB: 58 },
            { POB: 58 },
            { POB: 59 },
            { POB: 59 },
            { POB: 59 },
            { POB: 59 },
            { POB: 59 },
            { POB: 62 },
            { POB: 62 },
            { POB: 73 },
            { width: 120, POB: 73 },
            { POB: 73 },
            { POB: 67 },
            { POB: 66 },
            { POB: 78 },
            { POB: 77 },
            { POB: 75 },
            { POB: 75 },
            { POB: 75 },
        ];
        this.chartPriorities = [];
        this.chartPrioritiesx = [
            {
                id: 1,
                order: 1,
                name: "HSSE Critical (control, monitor, intervene)",
                activities: [],
            },
            {
                id: 2,
                order: 2,
                name: "MIE (PM and CM including P1 FM)",
                activities: []
            },
            {
                id: 3,
                order: 3,
                name: "Production Critical inc. DCP, CMs and PMs",
                activities: []
            },
            {
                id: 4,
                order: 4,
                name: "Projects",
                activities: []
            },
            {
                id: 5,
                order: 5,
                name: "Fabric Maintenance",
                activities: []
            },
            {
                id: 6,
                order: 6,
                name: "Opportunity Survey/Visits",
                activities: []
            }
        ];
        this.handPrevActivityId = -1;
        this.handWidthValue = -1;
        this.handHeightValue = -1;
        this.handTopValue = -1;
        this.handLeftValue = -1;
        this.modActs = [];
        this.printMode = false;
        this.oldScale = this.udf.calActivityChartScale;
    }
    FrmChartActivitiesComponent.prototype.ngOnInit = function () {
        this.oplandb.clearAllSubscription();
        this.onResize(null);
        this.getSummaryData();
        //getSummaryData
        //this.initVars();
    };
    FrmChartActivitiesComponent.prototype.ngOnDestroy = function () {
        this.oplandb.clearSubscription(this.subsSummary);
    };
    FrmChartActivitiesComponent.prototype.dragValid = function () {
        return true;
    };
    FrmChartActivitiesComponent.prototype.maskTop = function () {
        var body = __WEBPACK_IMPORTED_MODULE_3_jquery__('#stiap_body');
        return body.scrollTop();
    };
    FrmChartActivitiesComponent.prototype.bodyScroll = function (e) {
        // created and assigned this function to (scroll) event to trigger update to 
        // dragMask parameters...
    };
    FrmChartActivitiesComponent.prototype.maskWidth = function () {
        var body = __WEBPACK_IMPORTED_MODULE_3_jquery__('#stiap_body');
        return body.width();
    };
    FrmChartActivitiesComponent.prototype.maskHeight = function () {
        var body = __WEBPACK_IMPORTED_MODULE_3_jquery__('#stiap_body');
        return body.height();
    };
    FrmChartActivitiesComponent.prototype.curMove = function () {
        return this.noMove ? "no-drop" : "move";
    };
    FrmChartActivitiesComponent.prototype.curLeft = function () {
        return this.noLeft ? "no-drop" : "ew-resize";
    };
    FrmChartActivitiesComponent.prototype.curRight = function () {
        return this.noRight ? "no-drop" : "ew-resize";
    };
    FrmChartActivitiesComponent.prototype.handWidth = function () {
        if (this.handWidthValue == -1) {
            var obj = __WEBPACK_IMPORTED_MODULE_3_jquery__(this.hovObjSrc);
            //if(obj.length)this.handWidthValue = obj.width()+6;
            var act = this.hovAct;
            if (act)
                this.handWidthValue = this.dateSpanToPx(act.sd, act.ed);
            //dateSpanToPx
        }
        return this.handWidthValue;
    };
    FrmChartActivitiesComponent.prototype.handHeight = function () {
        if (this.handHeightValue == -1) {
            var obj = __WEBPACK_IMPORTED_MODULE_3_jquery__(this.hovObjSrc);
            if (obj.length)
                this.handHeightValue = obj.height() + 6;
        }
        return this.handHeightValue;
    };
    FrmChartActivitiesComponent.prototype.handTop = function () {
        if (this.handTopValue == -1) {
            var obj = __WEBPACK_IMPORTED_MODULE_3_jquery__(this.hovObjSrc);
            if (obj.length) {
                var body = __WEBPACK_IMPORTED_MODULE_3_jquery__('#stiap_body');
                var par = obj.parent().parent();
                var par2 = par.parent();
                this.handTopValue = body.scrollTop() + par.position().top + par2.position().top;
                //$$$debug
            }
            else {
                //this.handTopValue = 0;
            }
        }
        return this.handTopValue;
    };
    FrmChartActivitiesComponent.prototype.handLeft = function () {
        if (this.handLeftValue == -1) {
            var obj = __WEBPACK_IMPORTED_MODULE_3_jquery__(this.hovObjSrc);
            if (obj.length) {
                var par = obj.parent();
                //this.handLeftValue = this.chartRowHeaderWidth() + par.position().left;
                var act = this.hovAct;
                this.handLeftValue = this.dateToPx(act.sd);
            }
            else {
                //this.handLeftValue = this.chartRowHeaderWidth();
            }
        }
        return this.handLeftValue;
    };
    FrmChartActivitiesComponent.prototype.validateChange = function () {
        var act = this.hovAct;
        // set newStart and/or newEnd to original activity startan
        if (this.hovNewStart.length == 0)
            this.hovNewStart = act.sd;
        if (this.hovNewEnd.length == 0)
            this.hovNewEnd = act.ed;
        var sd = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](act.sd);
        var ed = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](act.ed);
        var mind = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](act.minDate);
        var maxd = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](act.maxDate);
        var nsd = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](this.hovNewStart);
        var ned = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](this.hovNewEnd);
        var diffOld = __WEBPACK_IMPORTED_MODULE_1__app_globals__["daysDiff"](sd, ed); // duration between the old start and end dates
        var diffNew = __WEBPACK_IMPORTED_MODULE_1__app_globals__["daysDiff"](nsd, ned); // duration between the new start and end dates
        var diffStartMin = __WEBPACK_IMPORTED_MODULE_1__app_globals__["daysDiff"](sd, mind); // days from activity start to earliest member assignment
        var diffMinMax = __WEBPACK_IMPORTED_MODULE_1__app_globals__["daysDiff"](mind, maxd); // days between the minimum and maximum assignment dates
        if (diffOld != diffNew && !act.block) {
            // if activity duration is changed, not a block activity, trigger validation
            // to make sure all members will still fit inside the new duration
            // required duration for the activity to be able to accommodate all members
            var reqDuration = diffStartMin + diffMinMax;
            // new duration new date span less take off offset
            var newDuration = diffNew;
            //this.debugText +=  ", reqDuration:" + reqDuration + ", newDuration:" + newDuration
            // if newDuration is less than the required duration, changes made will not be allowed
            if (newDuration < reqDuration)
                return false; // new duration will not be able to accommodate all member assinments
        }
        return true;
    };
    FrmChartActivitiesComponent.prototype.mouseEvent = function (e, obj, pri, row, rowIndex) {
        var _this = this;
        if (!this.dragEditing)
            return; // fire up mouse event only if dragEditing is true
        if (!obj)
            obj = null;
        var dw = this.chartColWidth(); // day column width in pixel
        if (e.type == "mouseenter") {
            // activity event
            this.debugText = e.type + "(" + pri.id + ")";
            if (pri.id == 7)
                return;
            this.debugText = e.type + "(1)";
            //if(obj.memCount==0)return;  // obj = activity, memcount=0??
            this.debugText = e.type + "(2)";
            this.hovObjSrc = e.target;
            this.hovAct = obj;
            this.hovPri = pri;
            this.hovRow = row;
            this.hovRowIdx = rowIndex;
            if (obj) {
                if (this.handPrevActivityId != obj.id) {
                    // reset handle dimensions and reset previous activity id
                    this.handPrevActivityId = obj.id;
                    this.handTopValue = -1;
                    this.handLeftValue = -1;
                    this.handHeightValue = -1;
                    this.handWidthValue = -1;
                    this.hovNewStart = "";
                    this.hovNewEnd = "";
                    this.hovNewRow = -1;
                    this.hovNewRowIdx = -1;
                }
                //this.debugText = JSON.stringify(obj) + ", " + this.handWidthValue;
            }
        }
        else if (e.type == 'mousemove') {
            // mask event
            var newWidth = void 0;
            var off = void 0;
            var offY = void 0;
            var h = __WEBPACK_IMPORTED_MODULE_3_jquery__("#dragHandle");
            //this.debugText = this.chartRowHeaderWidth() + ", " +  e.x + ", " + h.offset().left + ", " + h.position().left + ", " + this.dragMode;
            if (this.dragMode == "e-resize") {
                var body = __WEBPACK_IMPORTED_MODULE_3_jquery__('#stiap_body');
                var eX = (this.udf.isIEOnly ? body.offset().left : 0) + e.x;
                var obj_1 = __WEBPACK_IMPORTED_MODULE_3_jquery__(this.hovObjSrc);
                var par = obj_1.parent().parent();
                var parX = par.offset().left;
                //let pickedPx:number = e.x-parX;
                var pickedPx = eX - parX;
                var dateOnClicked = this.pxToDate(pickedPx + dw / 2);
                var dateStringOnClicked = __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](dateOnClicked);
                newWidth = (this.dateToPx(dateStringOnClicked) - h.position().left);
                if (newWidth >= dw) {
                    this.hovNewEnd = __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](__WEBPACK_IMPORTED_MODULE_1__app_globals__["addDays"](dateOnClicked, -1));
                    this.handWidthValue = newWidth;
                }
                this.debugText = "e-resize!!!(newWidth): " + newWidth + ", e.x:" + e.x + ", " + eX;
            }
            else if (this.dragMode == "w-resize") {
                var body = __WEBPACK_IMPORTED_MODULE_3_jquery__('#stiap_body');
                var eX = (this.udf.isIEOnly ? body.offset().left : 0) + e.x;
                var obj_2 = __WEBPACK_IMPORTED_MODULE_3_jquery__(this.hovObjSrc);
                var par = obj_2.parent().parent();
                var parX = par.offset().left;
                //let pickedPx:number = e.x-parX;
                var pickedPx = eX - parX;
                var dateOnClicked = this.pxToDate(pickedPx + dw / 2.5);
                var dateStringOnClicked = __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](dateOnClicked);
                var newLeft = this.dateToPx(dateStringOnClicked);
                var newWidth_1 = this.handWidthValue - (newLeft - h.position().left);
                if (newWidth_1 >= dw) {
                    this.handWidthValue = newWidth_1;
                    this.handLeftValue = newLeft;
                    this.hovNewStart = dateStringOnClicked;
                }
                //1this.noLeft = !this.validateChange(); // temporarily removed, not working as desired
            }
            else if (this.dragMode == "move_x") {
            }
            else if (this.dragMode == "move") {
                // this entiree setion perfectly works on chrome but not on IE
                var body = __WEBPACK_IMPORTED_MODULE_3_jquery__('#stiap_body');
                var bodyOffestTop = this.udf.isIEOnly ? body.offset().top : 0;
                var eY = e.y + bodyOffestTop;
                var obj_3 = __WEBPACK_IMPORTED_MODULE_3_jquery__(this.hovObjSrc);
                var par = obj_3.parent().parent();
                var par2 = par.parent();
                var parY = par2.position().top; //+ par2.position().top;
                var parH = parY + par2.height();
                var parX = par.offset().left;
                var pickedPx = e.x - parX;
                var dateOnClicked = this.pxToDate(pickedPx - this.mouseDownXOff + dw / 2);
                var dateStringOnClicked = __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](dateOnClicked);
                this.handLeftValue = this.dateToPx(dateStringOnClicked);
                var dateOff = __WEBPACK_IMPORTED_MODULE_1__app_globals__["daysDiff"](dateOnClicked, __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](this.hovAct.sd));
                var newEndDate = __WEBPACK_IMPORTED_MODULE_1__app_globals__["addDays"](__WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](this.hovAct.ed), dateStringOnClicked < this.hovAct.sd ? -dateOff : dateOff);
                this.hovNewStart = dateStringOnClicked;
                this.hovNewEnd = __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](newEndDate);
                //offY = (e.y+bodyOffestTop)-(this.mouseDownY+bodyOffestTop);
                offY = 0;
                var newTop = this.handTopValue + offY;
                var barHeight = this.chartRowHeight() * (this.hovPri.id == 7 ? 0.5 : 1);
                var topLimit = parY + body.scrollTop();
                var botLimit = parH + body.scrollTop(); //-obj.height();
                if (newTop >= topLimit) {
                    var intTop = parseInt(String(((eY - body.offset().top + body.scrollTop()) - topLimit) / barHeight));
                    var intPx = intTop * barHeight;
                    if ((intPx + topLimit) > topLimit) {
                        newTop = intPx + topLimit;
                    }
                    else {
                        newTop = topLimit;
                    }
                    if (intTop <= this.hovPri.actRows.length) {
                        this.hovNewRowIdx = intTop;
                        this.handTopValue = newTop;
                        this.mouseDownY = eY;
                        if (intTop == this.hovPri.actRows.length) {
                            this.hovPri.actRows[this.hovPri.actRows.length] = [];
                        }
                    }
                } // if(newTop>=topLimit)
                /*
                this.debugText = " this.handWidthValue:"+this.handWidthValue +
                " $('#stiap_body').scrollTop():" + $('#stiap_body').scrollTop() +
                " $(this.hovObjSrc).parent().parent().position().top:" + $(this.hovObjSrc).parent().parent().position().top+
                " $(this.hovObjSrc).parent().parent().parent().position().top:" + $(this.hovObjSrc).parent().parent().parent().position().top+
                " $(this.hovObjSrc).offset().top:" + $(this.hovObjSrc).offset().top +
                " this.bannerHeight:" + this.udf.bannerHeight +
                " par2.id: " + $(par2).attr('id')+
                " par2.tagName: " + $(par2).prop('tagName')+
                " newTop: " + newTop+
                " barHeight: " + barHeight+
                " $('#stiap_body')offset().top: " + $('#stiap_body').offset().top+
                " this.mouseDownY: " + this.mouseDownY+
                " e.y: " + e.y +" ,  "+ eY+
                " offY: " + offY+
                " parY: " + parY+
                " topLimit: " + topLimit+
                " isIEOnly: " + this.udf.isIEOnly +
                " this.handTopValue: " + this.handTopValue+
                " this.handLeftValue: " + this.handLeftValue;
                */
            }
        }
        else if (e.type == 'mouseup_bypass') {
            this.hovObj = null;
            this.hovObjSrc = null;
            var act = this.hovAct;
        }
        else if (e.type == 'mouseup') {
            // mask event
            // this entiree setion perfectly works on chrome but not on IE
            this.hovObj = null;
            this.hovObjSrc = null;
            var act_1 = this.hovAct;
            if (this.hovRowIdx != this.hovNewRowIdx && this.hovNewRowIdx != -1) {
                // if activity is moved to a different row
                var rIdx = void 0;
                for (rIdx = 0; rIdx < this.hovRow.length; rIdx++) {
                    if (act_1.id == this.hovRow[rIdx].id) {
                        // remove activity from original row
                        this.hovRow.splice(rIdx, 1);
                        break;
                    }
                }
                // add activity to the new row
                this.hovPri.actRows[this.hovNewRowIdx].push(act_1);
                this.hovRowIdx = this.hovNewRowIdx;
                this.hovRow = this.hovPri.actRows[this.hovRowIdx];
                // clean all priority rows that are empty
                var withEmpty = true;
                while (withEmpty) {
                    withEmpty = false;
                    for (rIdx = 0; rIdx < this.hovPri.actRows.length; rIdx++) {
                        if (this.hovPri.actRows[rIdx].length == 0) {
                            // purge row, flag withEmpty
                            withEmpty = true;
                            this.hovPri.actRows.splice(rIdx, 1);
                            break;
                        }
                    }
                }
            } // activity is moved to a different row
            this.debugText =
                "@sd:" + act_1.sd +
                    ", ed:" + act_1.ed +
                    ", this.hovRowIdx:" + this.hovRowIdx +
                    ", this.hovNewRowIdx:" + this.hovNewRowIdx +
                    "<br/>act:" + JSON.stringify(this.hovAct) +
                    //"<br/>actBak:" + JSON.stringify(actBak);
                    //"<br/> Priority: " + JSON.stringify(this.hovPri)+
                    //"<br/> Row: " + JSON.stringify(this.hovRow)+
                    "";
            // set newStart and/or newEnd to original activity startan
            if (this.hovNewStart.length == 0)
                this.hovNewStart = act_1.sd;
            if (this.hovNewEnd.length == 0)
                this.hovNewEnd = act_1.ed;
            // modified 2019/01/23 to accommodate movememnt when members=0;
            //if((this.hovNewStart != act.sd || this.hovNewEnd != act.ed) && act.memCount!=0){
            if ((this.hovNewStart != act_1.sd || this.hovNewEnd != act_1.ed)) {
                // if start or end date is changed and members exist under the activity
                var sd_1 = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](act_1.sd);
                var nsd_1 = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](this.hovNewStart);
                var mind_1 = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](act_1.minDate);
                var maxd_1 = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](act_1.maxDate);
                // modified 2019/01/23 to accommodate movememnt when members=0;
                //if(this.validateChange()){
                if (this.validateChange() || act_1.memCount == 0) {
                    // remember old start and end dates before posting ...
                    this.hovOldStart = act_1.sd;
                    this.hovOldEnd = act_1.ed;
                    this.postActivityDateChange("update", "Posting activity date changes. Please wait...", function (data) {
                        if (data.success) {
                            // if postin
                            // backup old dates is not yet existing, before assigning new values
                            if (!act_1.dateBackup)
                                act_1.dateBackup = { sd: act_1.sd, ed: act_1.ed, minDate: act_1.minDate, maxDate: act_1.maxDate };
                            // if validation is successful, assign new date values and 
                            act_1.sd = _this.hovNewStart;
                            act_1.ed = _this.hovNewEnd;
                            if (act_1.block) {
                                // if activity is a block activity, make minDate and maxDate similar to the new start and end dates
                                act_1.minDate = _this.hovNewStart;
                                act_1.maxDate = _this.hovNewEnd;
                            }
                            else {
                                // if not a block activity, set minDate and maxDate to original dates +/- diffStartMin
                                // offset value is negative id new start date is earlier
                                var offDate = (nsd_1 < sd_1 ? -1 : 1) * __WEBPACK_IMPORTED_MODULE_1__app_globals__["daysDiff"](sd_1, nsd_1); // offset days between old start date and new start date
                                act_1.minDate = __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](__WEBPACK_IMPORTED_MODULE_1__app_globals__["addDays"](mind_1, offDate));
                                act_1.maxDate = __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](__WEBPACK_IMPORTED_MODULE_1__app_globals__["addDays"](maxd_1, offDate));
                            }
                            _this.updatePOBCount(data);
                        }
                        _this.debugText = JSON.stringify(data) + "<br/>" + JSON.stringify(_this.hovAct);
                    });
                    // test pob count alteration
                    //this.chartCols[1]['POB'] = Number(this.chartCols[1]['POB']) + 2;
                }
                else {
                    //alert('invalid!');
                    // reset measurements
                    this.handLeftValue = -1;
                    this.handWidthValue = -1;
                    alert("Change to activity period will not be able to accommodate all member assignments.");
                }
            }
            this.handTopValue = -1;
            this.handLeftValue = -1;
            this.handHeightValue = -1;
            this.handWidthValue = -1;
        }
        else if (e.type == 'mousedown') {
            // activity event
            if (this.hovObjSrc) {
                var h = __WEBPACK_IMPORTED_MODULE_3_jquery__("#dragHandle");
                this.hovObj = this.hovObjSrc;
                //this.debugText = e.target.id;
                var dragId = e.target.id;
                this.mouseDownX = e.x;
                this.mouseDownY = e.y;
                this.mouseDownXOff = this.mouseDownX - h.offset().left;
                this.dragHandleId = dragId;
                this.dragMode = (dragId == "dragLeft" ? "w-resize" : (dragId == "dragRight" ? "e-resize" : (dragId == "dragHandle" ? "move" : "")));
            }
        }
    };
    FrmChartActivitiesComponent.prototype.dateSpanToPx = function (sd, ed) {
        var dw = this.chartColWidth(); // day column width in pixel
        var ds = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](sd);
        var de = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](ed);
        var ddiff = __WEBPACK_IMPORTED_MODULE_1__app_globals__["daysDiff"](ds, de);
        var ret = (ddiff + 1) * dw;
        return ret;
    };
    FrmChartActivitiesComponent.prototype.pxToDate = function (px) {
        var rhw = this.chartRowHeaderWidth(); // row header width offset before day 1
        var dw = this.chartColWidth(); // day column width in pixel
        var selStartDate = new Date(this.udf.calActivityYear, this.udf.calActivityMonth, 1);
        var pastDays = parseInt(String(px / dw));
        var ret = __WEBPACK_IMPORTED_MODULE_1__app_globals__["addDays"](selStartDate, pastDays);
        return ret;
    };
    FrmChartActivitiesComponent.prototype.dateToPx = function (dt) {
        var rhw = this.chartRowHeaderWidth(); // row header width offset before day 1
        var dw = this.chartColWidth(); // day column width in pixel
        var cdt = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](dt);
        // calculate span from the start date to srart date of the selected date
        var selStartDate = new Date(this.udf.calActivityYear, this.udf.calActivityMonth, 1);
        var ddiff = __WEBPACK_IMPORTED_MODULE_1__app_globals__["daysDiff"](cdt, selStartDate);
        //this.debugText ="Start date: " +  dt + ", selStartDate:" + selStartDate + ", ddiff:"+ddiff;
        var ret = rhw + (ddiff * dw * (cdt > selStartDate ? 1 : -1));
        return ret;
    };
    FrmChartActivitiesComponent.prototype.onResize = function (e) {
        this.setWindowSizeVars();
    };
    FrmChartActivitiesComponent.prototype.setWindowSizeVars = function () {
        this.windowWidth = __WEBPACK_IMPORTED_MODULE_3_jquery__(window).width();
        this.windowHeight = __WEBPACK_IMPORTED_MODULE_3_jquery__(window).height();
    };
    FrmChartActivitiesComponent.prototype.hoverText = function (act) {
        //return act.id
        return act.title +
            "\nStart: " + __WEBPACK_IMPORTED_MODULE_1__app_globals__["shortToLongDate"](act.sd) +
            "\nEnd: " + __WEBPACK_IMPORTED_MODULE_1__app_globals__["shortToLongDate"](act.ed) +
            (this.dragEditing ? "" : "\n\n(Double click to go to activity record...)");
    };
    FrmChartActivitiesComponent.prototype.chartFontScale = function () {
        if (this.udf.calActivityChartScale > 1) {
            return Math.max(1, this.udf.calActivityChartScale * 0.85); // apply font scale ceiling to prevent super large text
        }
        else {
            return Math.max(this.udf.calActivityChartScale, 0.7);
        }
    };
    FrmChartActivitiesComponent.prototype.chartSeparatorHeight = function () {
        return 3 * this.udf.calActivityChartScale;
    };
    FrmChartActivitiesComponent.prototype.chartColumnHeaderHeight = function () {
        return 48 * this.udf.calActivityChartScale;
    };
    FrmChartActivitiesComponent.prototype.chartColWidth = function () {
        return this.chartDefWidth * this.udf.calActivityChartScale;
    };
    FrmChartActivitiesComponent.prototype.chartRowHeight = function () {
        return this.rowHeight * this.udf.calActivityChartScale;
    };
    FrmChartActivitiesComponent.prototype.chartRowHeaderWidth = function () {
        return this.rowHeaderWidth * this.udf.calActivityChartScale;
    };
    FrmChartActivitiesComponent.prototype.chartWidth = function () {
        return (this.rowHeaderWidth + this.chartCols.length * this.chartDefWidth) * this.udf.calActivityChartScale;
    };
    FrmChartActivitiesComponent.prototype.calDays = function () {
        var d1 = new Date(this.udf.calActivityYear, this.udf.calActivityMonth, 1);
        var d2 = new Date(this.udf.calActivityYear, Number(this.udf.calActivityMonth) + 1, 1);
        var days = __WEBPACK_IMPORTED_MODULE_1__app_globals__["daysDiff"](d1, d2);
        //20170906,20170914
        this.arrangeActivities();
        //let act:object=this.chartPriorities[0]['actRows'][0][1];
        return days;
    };
    FrmChartActivitiesComponent.prototype.getColDate = function (idx) {
        var dt = new Date(this.udf.calActivityYear, this.udf.calActivityMonth, idx);
        return __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](dt);
    };
    FrmChartActivitiesComponent.prototype.activityStart = function (act) {
        var ret = 0;
        var d1 = new Date(this.udf.calActivityYear, this.udf.calActivityMonth, 1);
        var d2 = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](act.sd);
        if (d2 < d1)
            d2 = d1;
        ret = __WEBPACK_IMPORTED_MODULE_1__app_globals__["daysDiff"](d1, d2) * this.chartDefWidth;
        return ret;
    };
    FrmChartActivitiesComponent.prototype.activityWidth = function (act) {
        var ret = 0;
        var d0 = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](act.sd);
        var d1 = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](act.ed);
        var d2 = new Date(this.udf.calActivityYear, Number(this.udf.calActivityMonth) + 1, 0);
        var d3 = new Date(this.udf.calActivityYear, this.udf.calActivityMonth, 1);
        if (d1 > d2)
            d1 = d2;
        if (d3 > d0)
            d0 = d3;
        ret = (__WEBPACK_IMPORTED_MODULE_1__app_globals__["daysDiff"](d0, d1) + 1) * this.chartDefWidth;
        return ret;
    };
    FrmChartActivitiesComponent.prototype.readyColor = function (act) {
        var clr = act.ready;
        var color = "White";
        var back = "white";
        var border = act.vessel ? "3px solid blue" : "1px solid #c0c0c0";
        var modified = (act.dateBackup ? true : false);
        switch (clr) {
            case "Red":
                back = modified ? "Pink" : "Red";
                color = modified ? "black" : "white";
                break;
            case "Green":
                back = modified ? "Lime" : "Green";
                color = modified ? "black" : "white";
                break;
            case "Amber":
                back = modified ? "Yellow" : "orange";
                color = "black";
                break;
            default:
        }
        return { "color": color, "background": back, "border": border };
    };
    FrmChartActivitiesComponent.prototype.arrangeActivities = function () {
        var _this = this;
        this.chartPriorities.forEach(function (pri) {
            var actRows = [];
            var actsClone = _this.oplandb.cloneObject(pri['activities']);
            __WEBPACK_IMPORTED_MODULE_1__app_globals__["sortBy"](actsClone, "sd");
            var actRowObj = _this.getActivityRow(actsClone);
            while (actsClone.length > 0) {
                actRowObj = _this.getActivityRow(actsClone);
                actsClone = actRowObj['newSource'];
                actRows.push(actRowObj['actRow']);
            }
            pri['actRows'] = actRows;
        });
        /*let actsNew:Array<object>=[];
        let actsTemp:Array<object>=[];
        let actsComp:Array<object>=[];
    
    
        
        let newRef:object=actsClone[0];
    
    
        actsTemp.push(newRef);
        for(let i=1;i<actsClone.length;i++){
          actsComp.push(actsClone[i]);
        }
    
        for(let i=0;i<actsComp.length;i++){
          let obj:object=actsComp[i];
          if(obj['sd']>newRef['ed']){
            actsTemp.push(obj);
            newRef=obj;
          }else{
            actsNew.push(obj);
          }
        }
    
    
        return;
          
        /*do{
          newRef=actsClone[0];
          actsNew.push(newRef);
          for(let i=0;i<actsComp.length;i++){
    
          }
          return;
        } while(actsClone.length>1)*/
    };
    FrmChartActivitiesComponent.prototype.getActivityRow = function (actRow) {
        var actsNew = [];
        var actsTemp = [];
        var actsComp = [];
        var newRef = actRow[0];
        actsTemp.push(newRef); // use the first element as starting point
        for (var i = 1; i < actRow.length; i++) {
            actsComp.push(actRow[i]);
        }
        for (var i = 0; i < actsComp.length; i++) {
            var obj = actsComp[i];
            if (obj['sd'] > newRef['ed']) {
                actsTemp.push(obj);
                newRef = obj;
            }
            else {
                actsNew.push(obj);
            }
        }
        return {
            newSource: this.oplandb.cloneObject(actsNew),
            actRow: this.oplandb.cloneObject(actsTemp)
        };
    };
    FrmChartActivitiesComponent.prototype.getTitle = function () {
        var ret = this.udf.calActivitySite + " STIAP View for " + __WEBPACK_IMPORTED_MODULE_1__app_globals__["monthNames"][this.udf.calActivityMonth] + ' ' + this.udf.calActivityYear;
        //SWP - {{G.monthNames[udf.calActivityMonth]}} {{udf.calActivityYear}}
        return ret;
    };
    FrmChartActivitiesComponent.prototype.getWeeks = function () {
        return this.chartWeeks;
    };
    FrmChartActivitiesComponent.prototype.initVars = function (dat) {
        var _this = this;
        var days = this.calDays();
        this.chartCols = [];
        this.chartWeeks = [];
        var ctr = 0;
        for (var i = 1; i <= days; i++) {
            var dt = new Date(this.udf.calActivityYear, this.udf.calActivityMonth, i);
            var dy = dt.getDay();
            ctr++;
            if (dy == 6) {
                // reset counter
                this.chartWeeks[this.chartWeeks.length] = ctr;
                ctr = 0;
            }
            this.chartCols[this.chartCols.length] = { POB: 0 };
        }
        if (dat != undefined)
            this.updatePOBCount(dat);
        if (ctr != 0)
            this.chartWeeks[this.chartWeeks.length] = ctr;
        ///chartPriorities
        if (dat != undefined) {
            this.chartPriorities = [];
            dat['TYPES'].forEach(function (type) {
                //if(type[0]!=0 && type[2]!=999){   // exclude blank type and priority 999
                if (type[0] != 0) {
                    _this.chartPriorities.push({ id: type[0], order: type[2],
                        name: type[1], activities: [], actRows: [] });
                }
            });
            dat['ACTS'].forEach(function (act) {
                var aType = Number(act[1]);
                var pri = _this.chartPriorities.find(function (e) {
                    return (e.id == aType);
                });
                if (pri != undefined) {
                    pri['activities'].push({ id: act[0], sd: act[3], ed: act[4], title: act[2], ready: act[5], vessel: act[6], block: act[7], memCount: act[8], minDate: act[9], maxDate: act[10] });
                }
            });
            this.arrangeActivities();
        }
    };
    FrmChartActivitiesComponent.prototype.onBeforePrint = function (event) {
        // 1052.16:0.96 px for A4
        this.printMode = true;
        this.oldScale = this.udf.calActivityChartScale;
        this.udf.calActivityChartScale = 1;
        //let tbl=document.getElementById('cardDetailsTable');
        //document.body.appendChild(tbl);
    };
    FrmChartActivitiesComponent.prototype.onAfterPrint = function (event) {
        this.printMode = false;
        this.udf.calActivityChartScale = this.oldScale;
    };
    FrmChartActivitiesComponent.prototype.totalBodyHeight = function () {
        var _this = this;
        var ht = 0;
        this.chartPriorities.forEach(function (pri) {
            ht += Math.max(pri.actRows.length, 1) * _this.chartRowHeight();
        });
        return ht + (this.rowHeaderHeight * this.udf.calActivityChartScale) * 3 +
            this.chartSeparatorHeight() * 5;
    };
    FrmChartActivitiesComponent.prototype.getSummaryData = function () {
        var _this = this;
        this.subsSummary = this.oplandb.processData({
            method: "getCalendarSummary",
            data: {
                site: this.udf.calActivitySite,
                startDate: __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](new Date(this.udf.calActivityYear, this.udf.calActivityMonth, 1)),
                endDate: __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](new Date(this.udf.calActivityYear, Number(this.udf.calActivityMonth) + 1, 0))
            },
            isPosting: false,
            processMessage: "Loading activity integrated view data. Please wait...",
            subscription: this.subsSummary,
            onSuccess: function (data) {
                //if(this.calActivities!=undefined)this.calActivities.initVars(data);
                _this.initVars(data);
                _this.oplandb.isLoading = false;
            }
        });
    };
    FrmChartActivitiesComponent.prototype.toggleEditMode = function (mode, conf) {
        var _this = this;
        var modActs = this.getModifiedActivities();
        if ((mode == "cancel" || mode == "save") && modActs.length == 0) {
            // check if there is anything to cancel or save. if none, no need to proceed
            // to posting data, just turn off dragEditing switch
            alert("No modified activities to " + (mode == "save" ? "save changes." : "rollback changes on."));
            this.dragEditing = false;
            return;
        }
        if (mode == "cancel") {
            if (!confirm("You are about to CANCEL all changes made to activity periods.\nClick OK to continue or CANCEL to abort."))
                return;
        }
        if (mode == "save") {
            if (!confirm("You are about to COMMIT/SAVE all changes made to activity periods.\nClick OK to continue or CANCEL to abort."))
                return;
        }
        this.postActivityDateChange(mode, mode == "initialize" ?
            "Initialing Drag and Drop mode. Please wait..." :
            mode == "cancel" ?
                "Rolling back activity updates. Please wait..." :
                "Saving all modified activity periods. Please wait...", function (data) {
            if (mode == "cancel") {
                // roll back all changes made on activity dates....
                _this.clearBackupDates(modActs, true);
                if (modActs.length == 1) {
                    alert("Rolled back changes made to an activity!");
                }
                else {
                    alert("Rolled back changes made to " + modActs.length + " activites!");
                }
            }
            else if (mode == "save") {
                _this.clearBackupDates(modActs);
                _this.oplandb.datActivities = null;
                if (modActs.length == 1) {
                    alert("Saved changes made to an activity!");
                }
                else {
                    alert("Saved changes made to " + modActs.length + "activites!");
                }
            }
            _this.updatePOBCount(data);
        });
        this.dragEditing = !this.dragEditing;
    };
    FrmChartActivitiesComponent.prototype.clearBackupDates = function (acts, rollback) {
        if (rollback == undefined)
            rollback = false;
        acts.forEach(function (act) {
            if (rollback) {
                act.sd = act.dateBackup.sd;
                act.ed = act.dateBackup.ed;
                act.minDate = act.dateBackup.minDate;
                act.maxDate = act.dateBackup.maxDate;
            }
            act.dateBackup = null;
        });
    };
    FrmChartActivitiesComponent.prototype.getModifiedActivities = function (clearBackup) {
        if (clearBackup == undefined)
            clearBackup = false;
        var ret = [];
        var isModified;
        this.chartPriorities.forEach(function (pri) {
            pri['actRows'].forEach(function (row) {
                row.forEach(function (act) {
                    isModified = (act.dateBackup ? true : false);
                    if (isModified)
                        ret.push(act);
                });
            });
        });
        return ret;
    };
    FrmChartActivitiesComponent.prototype.updatePOBCount = function (data) {
        var _this = this;
        // update daily POB numbers!
        if (data['POB'] != undefined) {
            if (data['POB'].length) {
                var idx_1 = 0;
                data['POB'].forEach(function (val) {
                    _this.chartCols[idx_1]['POB'] = val;
                    idx_1++;
                });
            }
        } // end of update POB
    };
    FrmChartActivitiesComponent.prototype.postActivityDateChange = function (postMode, postMessage, afterPost) {
        var postData = {
            site: this.udf.calActivitySite,
            startDate: __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](new Date(this.udf.calActivityYear, this.udf.calActivityMonth, 1)),
            endDate: __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](new Date(this.udf.calActivityYear, Number(this.udf.calActivityMonth) + 1, 0)),
            mode: postMode,
            uid: this.oplandb.logInfo.uid,
        };
        if (postMode == "update") {
            postData.actId = this.hovAct.id;
            postData.oldStartDate = this.hovAct.sd;
            postData.oldEndDate = this.hovAct.ed;
            postData.block = this.hovAct.block;
            postData.newStartDate = this.hovNewStart;
            postData.newEndDate = this.hovNewEnd;
        }
        /*
        newStartDate:this.hovNewStart,
        newEndDate:this.hovNewEnd,
        actId:this.hovAct.id,
        //"Calculating daily POB. Please wait..."
        */
        this.subsSummary = this.oplandb.processData({
            method: "setActivityDate",
            data: postData,
            isPosting: false,
            processMessage: postMessage,
            subscription: this.subsSummary,
            onSuccess: function (data) {
                if (afterPost != undefined) {
                    afterPost(data);
                }
                //if(this.calActivities!=undefined)this.calActivities.initVars(data);
                //this.initVars(data);
                //this.oplandb.isLoading=false;
                //this.chartCols[1]['POB'] = Number(this.chartCols[1]['POB']) + 2;
            }
        });
    };
    FrmChartActivitiesComponent.prototype.saveActivityDateChanges = function () {
    };
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["HostListener"])('window:beforeprint', ['$event']),
        __metadata("design:type", Function),
        __metadata("design:paramtypes", [Object]),
        __metadata("design:returntype", void 0)
    ], FrmChartActivitiesComponent.prototype, "onBeforePrint", null);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["HostListener"])('window:afterprint', ['$event']),
        __metadata("design:type", Function),
        __metadata("design:paramtypes", [Object]),
        __metadata("design:returntype", void 0)
    ], FrmChartActivitiesComponent.prototype, "onAfterPrint", null);
    FrmChartActivitiesComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-frm-chart-activities',
            template: __webpack_require__("../../../../../src/app/components/frm-chart-activities/frm-chart-activities.component.html"),
            styles: [__webpack_require__("../../../../../src/app/components/frm-chart-activities/frm-chart-activities.component.css")]
        }),
        __metadata("design:paramtypes", [typeof (_a = (typeof __WEBPACK_IMPORTED_MODULE_2__app_udfs__ !== "undefined" && __WEBPACK_IMPORTED_MODULE_2__app_udfs__).AppUDFService) === "function" && _a || Object, typeof (_b = typeof __WEBPACK_IMPORTED_MODULE_4__app_oplandata__["a" /* AppOplpanDataService */] !== "undefined" && __WEBPACK_IMPORTED_MODULE_4__app_oplandata__["a" /* AppOplpanDataService */]) === "function" && _b || Object])
    ], FrmChartActivitiesComponent);
    return FrmChartActivitiesComponent;
    var _a, _b;
}());

//# sourceMappingURL=frm-chart-activities.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/frm-chart/frm-chart.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "#dragHandle, #dragMask{\r\n    -ms-user-select: none;\r\n        user-select: none;\r\n    -moz-user-select: none;\r\n    -khtml-user-select: none;\r\n    -webkit-user-select: none;\r\n    -o-user-select: none;\r\n}\r\n\r\n.blink{\r\n    font-size: 1em;\r\n    color: rgb(217,83,66);\r\n    \r\n    -webkit-animation: blink 1s infinite;\r\n    \r\n            animation: blink 1s infinite;\r\n }\r\n \r\n @-webkit-keyframes blink{\r\n    0%{opacity: 1;}\r\n    75%{opacity: 1;}\r\n    76%{ opacity: 0;}\r\n    100%{opacity: 0;}\r\n }\r\n \r\n @keyframes blink{\r\n    0%{opacity: 1;}\r\n    75%{opacity: 1;}\r\n    76%{ opacity: 0;}\r\n    100%{opacity: 0;}\r\n }", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/frm-chart/frm-chart.component.html":
/***/ (function(module, exports) {

module.exports = "\n<div [innerHTML]=\"debugText\" id=\"debugDiv\" \n  *ngIf=\"debugOn\"\n  style=\"background:beige;position: absolute;z-index:100;top:0px;left:0px;height:65px;width:100%;overflow:auto\"\n  >\n</div>\n<button style=\"position:absolute;top:0px;left:0px;z-index:50;left:450px;cursor:pointer;\" \nclass=\"btn btn-link p-1\"\ndata-toggle=\"tooltip\" \n[style.top]=\"(udf.bannerHeight-udf.bannerMenuHeight+2)+'px'\"\n(click)=\"reloadCalendar()\"\ntitle=\"Reload calendar with the new specified scope dates ...\"\n[disabled]=\"disableReload()\"\n><i class=\"fa fa-refresh\" aria-hidden=\"true\"\nstyle=\"font-size:20px;\"\n></i></button>\n\n<button style=\"position:absolute;top:0px;z-index:50;left:480px;cursor:pointer;\" \nclass=\"btn btn-link p-1\"\ndata-toggle=\"tooltip\" \n[style.top]=\"(udf.bannerHeight-udf.bannerMenuHeight+3)+'px'\"\n(click)=\"downloadCalendar()\"\ntitle=\"Download activity calendar data ...\"\n><i class=\"fa fa-download\"  aria-hidden=\"true\"\nstyle=\"font-size:20px;\"\n></i></button>\n\n<div style=\"position:absolute;z-index:50;left:520px;\"\n  [style.top]=\"(udf.bannerHeight-udf.bannerMenuHeight+3)+'px'\"\n>\n  <button \n    style=\"cursor:pointer;text-decoration: none;font-size:1.1em;\" \n    class=\"btn btn-link p-1\"\n    onfocus=\"this.blur()\"\n    data-toggle=\"tooltip\" \n    (click)=\"toggleEditMode(oplandb.isEditing ? 'save' : '')\"\n    [title]=\"(oplandb.isEditing ? 'Save changes and exit drag & drop editing mode...' : 'Turn on drag & drop editing mode...')\"\n    ><i class=\"fa\" [ngClass]=\"{'fa-save':oplandb.isEditing,'text-success':oplandb.isEditing,'fa-pencil':!oplandb.isEditing}\"\n      aria-hidden=\"true\"></i>\n    <span *ngIf=\"!oplandb.isEditing\">...</span>\n  </button>\n  <button class=\"btn btn-link p-1 text-danger\" *ngIf=\"oplandb.isEditing\" \n    (click)=\"toggleEditMode('cancel',true)\"\n    style=\"cursor: pointer;\"\n    title=\"Cancel changes and exit drag & drop editing mode...\"\n    >\n    <i class=\"fa fa-ban\"  aria-hidden=\"true\"></i>\n  </button>\n  <span class=\"blink\" *ngIf=\"oplandb.isEditing\" style=\"cursor: default;\"> Drag and Drop editing mode is on !</span>\n\n\n  <div *ngIf=\"alertPrompt\" class=\"alert alert-danger px-4 py-1 m-0 text-center\"\n    role=\"alert\"\n    style=\"position:absolute;top:-2px;left:0px;width:785px;z-index:10;\">\n    <strong style=\"font-size:0.8em;\">{{alertPrompt}}</strong>\n    <button style=\"cursor: pointer;\" class=\"btn btn-sm btn-danger p-0 px-2 ml-2\" (click)=\"btnYes()\">Yes</button>\n    <button style=\"cursor: pointer;\" class=\"btn btn-sm btn-secondary p-0 px-2 ml-2\" (click)=\"btnNo()\">No</button>\n  </div> \n\n\n</div>\n\n<!--button class=\"btn btn-link p-1\">\n  <span><i class=\"fa fa-edit\"></i></span>\n</button-->\n\n\n<label *ngIf=\"statusText.length!=0\" class=\"form-label text-danger\"\nstyle=\"position:absolute;z-index:50;left:530px;font-size:0.9em;\"\n[style.top]=\"(udf.bannerHeight-udf.bannerMenuHeight + 7)+'px'\"\n>\n{{statusText}}\n</label>\n\n<!-- main scrolling wrapper -->\n<div id=\"mainContainer\" style=\"position:absolute;overflow:auto;\"\n  [style.top]=\"contTop('LH')\"\n  [style.width]=\"wrapWidth()\"\n  [style.height]=\"wrapHeight()\"\n  (scroll)=\"onElementScroll($event)\"\n  >\n\n  <div id=\"leftHeader\" name=\"ul_pane\"\n    style=\"z-index:3;position:absolute;overflow:hidden;\"\n    [style.width]=\"contWidth('LH')\"\n    [style.height]=\"contHeight('LH')\">\n\n    <svg \n      [attr.width]=\"svgWidth('LH')\" \n      [attr.height]=\"svgHeight('LH')\" \n      [attr.viewBox]=\"svgViewBox('LH')\"\n      style=\"background-color:white;margin:none;left:0px;\" \n      xmlns=\"http://www.w3.org/2000/svg\">\n\n      <!--rect x=\"0\" y=\"0\" width=\"100\" height=\"100\" fill=\"#f2f2f2\"/-->\n      <!-- Left Column header section -->\n      <rect x=\"0\" [attr.y]=\"gridYPos(7)\" \n        [attr.width]=\"svgWidth('LH')\" \n        [attr.height]=\"gridVSpace() * 3\" fill=\"#f2f2f2\"/>      \n      \n        \n      <!--g *ngIf=\"!noSummaryHeader\">\n        <text \n          [attr.x]=\"txt[2] ? txt[2] : 5\" \n          font-stretch=\"narrower\" \n          [attr.text-anchor]=\"txt[3] ? txt[3] : 'start'\" \n          [attr.y]=\"gridYPos(txt[0],5)\"  \n          font-family=\"Arial\" \n          [attr.font-size]=\"globalFontSize\" \n          [attr.fill]=\"globalTextFill\" \n          *ngFor=\"let txt of [\n          [1,'# OF MOB PASSENGERS',this.svgWidth('LH')-22,'end'],\n          [2,'# OF DEMOB PASSENGERS',this.svgWidth('LH')-22,'end']]\"\n          >\n          {{txt[1]}}\n        </text>\n      </g-->\n\n      <text \n        [attr.x]=\"txt[2] ? txt[2] : 5\" \n        font-stretch=\"narrower\" \n        [attr.text-anchor]=\"txt[3] ? txt[3] : 'start'\" \n        [attr.y]=\"gridYPos(txt[0],5)\"  \n        font-family=\"Arial\" \n        [attr.font-size]=\"globalFontSize\" \n        [attr.fill]=\"globalTextFill\" \n        *ngFor=\"let txt of [\n        [1-noHdr(),'# OF MOB PASSENGERS'],\n        [2-noHdr(),'# OF DEMOB PASSENGERS'],\n        [3-noHdr(),'NIGHT SHIFT REQUIREMENT'],\n        [4-noHdr(),'FEMALE CREW'],\n\n        [5-noHdr(),'TOTAL OFFSHORE POB'],\n        [6-noHdr(),'SWP POB LIMIT'],\n        [7-noHdr(),'AVAILABLE BEDS'],\n\n        [9-noHdr(),'NAME',this.leftHeaderWidths[1] + (this.leftHeaderWidths[0]-this.leftHeaderWidths[1])/2+(isFiltered() ? 40 : 0),'middle'],\n        [9-noHdr(),'POSITION',this.leftHeaderWidths[3] + (this.leftHeaderWidths[2]-this.leftHeaderWidths[3])/2,'middle'],\n        [8.2-noHdr(),'Functional',this.leftHeaderWidths[2] + (this.leftHeaderWidths[1]-this.leftHeaderWidths[2])/2,'middle'],\n        [9-noHdr(),'Sponsor &',this.leftHeaderWidths[2] + (this.leftHeaderWidths[1]-this.leftHeaderWidths[2])/2,'middle'],\n        [9.8-noHdr(),'Company',this.leftHeaderWidths[2] + (this.leftHeaderWidths[1]-this.leftHeaderWidths[2])/2,'middle']\n    ]\"\n      >\n        {{txt[1] }}\n      </text>\n\n      <svg *ngIf=\"isFiltered()\"  \n        (click)=\"applyRemoveFilter({date:dateToLong(udf.calFilterDate)})\"\n        width=\"18\" height=\"18\" viewBox=\"0 0 1500 1500\"\n        [attr.x]=\"90\"  \n        [attr.y]=\"gridYPos(8,-3)\" \n        style=\"cursor:pointer;box-shadow: 2px 2px 8px rgba(0,0,0,0.4);\"\n        >\n        <use  xlink:href=\"#filter\">\n          <title>{{'Personnel activity filter applied on '+ dateToLong(udf.calFilterDate) + '\\n(Click to remove filter)' }}</title>\n        </use>\n      </svg>\n\n      <!-- filter options -->\n      <svg *ngIf=\"isFiltered()\" width=\"150\" height=\"80\" viewBox=\"0 0 150 70\" x=\"20\"  \n        [attr.y]=\"gridYPos(6)\" >\n        <!--rect x=\"0\" y=\"0\" width=\"120\" height=\"40\" fill=\"yellow\" /-->\n\n        <circle\n          (click)=\"toggleFilter('pob')\"\n          [style.cursor]=\"'pointer'\"\n          cx=\"8\"\n            [attr.cy]=\"filterOffset-filterCircleOffset\"\n            [attr.fill]=\"udf.calFilterPOB ? 'lime' : '#c0c0c0'\"\n            fill=\"lime\"\n            stroke-width=\"2\"\n            stroke=\"#c0c0c0\"\n            r=\"5\"\n          >\n          <title>Toggle POB filter</title>\n        </circle>\n\n        <circle\n        (click)=\"toggleFilter('mob')\"\n        [style.cursor]=\"'pointer'\"\n          cx=\"8\"\n          [attr.cy]=\"filterOffset-filterCircleOffset+filterGap\"\n            [attr.fill]=\"udf.calFilterMOB ? 'lime' : '#c0c0c0'\"\n            stroke-width=\"2\"\n            stroke=\"#c0c0c0\"\n            r=\"5\"\n          >\n          <title>Toggle MOB filter</title>\n        </circle> \n\n        <circle\n          (click)=\"toggleFilter('demob')\"\n          [style.cursor]=\"'pointer'\"\n            cx=\"8\"\n            [attr.cy]=\"filterOffset-filterCircleOffset+filterGap*2\"\n            [attr.fill]=\"udf.calFilterDEMOB ? 'lime' : '#c0c0c0'\"\n            stroke-width=\"2\"\n            stroke=\"#c0c0c0\"\n            r=\"5\"\n          >\n          <title>Toggle DEMOB filter</title>\n        </circle> \n\n        <circle\n          (click)=\"toggleFilter('night')\"\n          [style.cursor]=\"'pointer'\"\n            cx=\"8\"\n            [attr.cy]=\"filterOffset-filterCircleOffset+filterGap*3\"\n            [attr.fill]=\"udf.calFilterNIGHT ? 'lime' : '#c0c0c0'\"\n            stroke-width=\"2\"\n            stroke=\"#c0c0c0\"\n            r=\"5\"\n          >\n          <title>Toggle night shift filter</title>\n        </circle>         \n    \n        <text style=\"cursor:pointer;\"\n          (click)=\"toggleFilter('pob')\"\n          x=\"15\" \n          [attr.y]=\"filterOffset\"\n          font-stretch=\"narrower\" \n          text-anchor=\"center\" \n          font-family=\"Arial\" \n          [attr.font-size]=\"10\" \n          [attr.fill]=\"udf.calFilterPOB ? '#5bc0de' : globalTextFill\"\n          >POB\n          <title>Toggle POB filter</title>\n        </text>\n      \n\n        <text style=\"cursor:pointer;\"\n        (click)=\"toggleFilter('mob')\"\n          x=\"15\" \n          [attr.y]=\"filterOffset+filterGap\"\n          font-stretch=\"narrower\" \n          text-anchor=\"center\" \n          font-family=\"Arial\" \n          [attr.font-size]=\"10\" \n          [attr.fill]=\"udf.calFilterMOB ? '#5bc0de' : globalTextFill\"\n          >MOB\n          <title>Toggle MOB filter</title>\n        </text>      \n\n        <text style=\"cursor:pointer;\"\n        (click)=\"toggleFilter('demob')\"\n          x=\"15\" \n          [attr.y]=\"filterOffset+filterGap*2\"\n          font-stretch=\"narrower\" \n          text-anchor=\"center\" \n          font-family=\"Arial\" \n          [attr.font-size]=\"10\" \n          [attr.fill]=\"udf.calFilterDEMOB ? '#5bc0de' : globalTextFill\"\n        >DEMOB\n        <title>Toggle DEMOB filter</title>\n      </text>      \n      \n        <text style=\"cursor:pointer;\"\n          (click)=\"toggleFilter('night')\"\n          x=\"15\" \n          [attr.y]=\"filterOffset+filterGap*3\"\n          font-stretch=\"narrower\" \n          text-anchor=\"center\" \n          font-family=\"Arial\" \n          [attr.font-size]=\"10\" \n          [attr.fill]=\"udf.calFilterNIGHT ? '#5bc0de' : globalTextFill\"\n        >NIGHT SHIFT\n        <title>Toggle Night Shift filter</title>\n      </text> \n                \n      </svg>\n      <!-- filter options -->\n      <symbol id=\"flag\">\n        <path d=\"M2.5 2q0 0.563-0.5 0.859v9.891q0 0.102-0.074 0.176t-0.176 0.074h-0.5q-0.102 0-0.176-0.074t-0.074-0.176v-9.891q-0.5-0.297-0.5-0.859 0-0.414 0.293-0.707t0.707-0.293 0.707 0.293 0.293 0.707zM14 2.5v5.961q0 0.195-0.098 0.301t-0.309 0.215q-1.68 0.906-2.883 0.906-0.477 0-0.965-0.172t-0.848-0.375-0.902-0.375-1.113-0.172q-1.5 0-3.625 1.141-0.133 0.070-0.258 0.070-0.203 0-0.352-0.148t-0.148-0.352v-5.797q0-0.25 0.242-0.43 0.164-0.109 0.617-0.336 1.844-0.937 3.289-0.937 0.836 0 1.563 0.227t1.711 0.688q0.297 0.148 0.688 0.148 0.422 0 0.918-0.164t0.859-0.367 0.688-0.367 0.426-0.164q0.203 0 0.352 0.148t0.148 0.352z\" stroke=\"red\" fill=\"#d9534f\"></path>\n      </symbol>\n\n      <symbol id=\"filter\">\n          <!--path d=\"M505.918,236.117c-26.651-43.587-62.485-78.609-107.497-105.065c-45.015-26.457-92.549-39.687-142.608-39.687\n          c-50.059,0-97.595,13.225-142.61,39.687C68.187,157.508,32.355,192.53,5.708,236.117C1.903,242.778,0,249.345,0,255.818\n          c0,6.473,1.903,13.04,5.708,19.699c26.647,43.589,62.479,78.614,107.495,105.064c45.015,26.46,92.551,39.68,142.61,39.68\n          c50.06,0,97.594-13.176,142.608-39.536c45.012-26.361,80.852-61.432,107.497-105.208c3.806-6.659,5.708-13.223,5.708-19.699\n          C511.626,249.345,509.724,242.778,505.918,236.117z M194.568,158.03c17.034-17.034,37.447-25.554,61.242-25.554\n          c3.805,0,7.043,1.336,9.709,3.999c2.662,2.664,4,5.901,4,9.707c0,3.809-1.338,7.044-3.994,9.704\n          c-2.662,2.667-5.902,3.999-9.708,3.999c-16.368,0-30.362,5.808-41.971,17.416c-11.613,11.615-17.416,25.603-17.416,41.971\n          c0,3.811-1.336,7.044-3.999,9.71c-2.667,2.668-5.901,3.999-9.707,3.999c-3.809,0-7.044-1.334-9.71-3.999\n          c-2.667-2.666-3.999-5.903-3.999-9.71C169.015,195.482,177.535,175.065,194.568,158.03z M379.867,349.04\n          c-38.164,23.12-79.514,34.687-124.054,34.687c-44.539,0-85.889-11.56-124.051-34.687s-69.901-54.2-95.215-93.222\n          c28.931-44.921,65.19-78.518,108.777-100.783c-11.61,19.792-17.417,41.207-17.417,64.236c0,35.216,12.517,65.329,37.544,90.362\n          s55.151,37.544,90.362,37.544c35.214,0,65.329-12.518,90.362-37.544s37.545-55.146,37.545-90.362\n          c0-23.029-5.808-44.447-17.419-64.236c43.585,22.265,79.846,55.865,108.776,100.783C449.767,294.84,418.031,325.913,379.867,349.04\n          z\"  stroke=\"none\" fill=\"#5bc0de\"/-->\n\n          <g\n          transform=\"matrix(1,0,0,-1,182.23729,1201.8983)\"\n          id=\"g3015\">\n        <path\n            d=\"m 1403,1241 q 17,-41 -14,-70 L 896,678 V -64 q 0,-42 -39,-59 -13,-5 -25,-5 -27,0 -45,19 L 531,147 q -19,19 -19,45 V 678 L 19,1171 q -31,29 -14,70 17,39 59,39 h 1280 q 42,0 59,-39 z\"\n            id=\"path3017\"\n            inkscape:connector-curvature=\"0\"\n            style=\"fill:#5bc0de;\" />\n          </g>\n      </symbol>\n      <symbol id=\"filter_gray\">\n          <!--path d=\"M505.918,236.117c-26.651-43.587-62.485-78.609-107.497-105.065c-45.015-26.457-92.549-39.687-142.608-39.687\n          c-50.059,0-97.595,13.225-142.61,39.687C68.187,157.508,32.355,192.53,5.708,236.117C1.903,242.778,0,249.345,0,255.818\n          c0,6.473,1.903,13.04,5.708,19.699c26.647,43.589,62.479,78.614,107.495,105.064c45.015,26.46,92.551,39.68,142.61,39.68\n          c50.06,0,97.594-13.176,142.608-39.536c45.012-26.361,80.852-61.432,107.497-105.208c3.806-6.659,5.708-13.223,5.708-19.699\n          C511.626,249.345,509.724,242.778,505.918,236.117z M194.568,158.03c17.034-17.034,37.447-25.554,61.242-25.554\n          c3.805,0,7.043,1.336,9.709,3.999c2.662,2.664,4,5.901,4,9.707c0,3.809-1.338,7.044-3.994,9.704\n          c-2.662,2.667-5.902,3.999-9.708,3.999c-16.368,0-30.362,5.808-41.971,17.416c-11.613,11.615-17.416,25.603-17.416,41.971\n          c0,3.811-1.336,7.044-3.999,9.71c-2.667,2.668-5.901,3.999-9.707,3.999c-3.809,0-7.044-1.334-9.71-3.999\n          c-2.667-2.666-3.999-5.903-3.999-9.71C169.015,195.482,177.535,175.065,194.568,158.03z M379.867,349.04\n          c-38.164,23.12-79.514,34.687-124.054,34.687c-44.539,0-85.889-11.56-124.051-34.687s-69.901-54.2-95.215-93.222\n          c28.931-44.921,65.19-78.518,108.777-100.783c-11.61,19.792-17.417,41.207-17.417,64.236c0,35.216,12.517,65.329,37.544,90.362\n          s55.151,37.544,90.362,37.544c35.214,0,65.329-12.518,90.362-37.544s37.545-55.146,37.545-90.362\n          c0-23.029-5.808-44.447-17.419-64.236c43.585,22.265,79.846,55.865,108.776,100.783C449.767,294.84,418.031,325.913,379.867,349.04\n          z\"  stroke=\"#c0c0c0\" fill=\"#c0c0c0\"/-->\n      \n          <g\n          transform=\"matrix(1,0,0,-1,182.23729,1201.8983)\"\n          id=\"g3015\">\n        <path\n            d=\"m 1403,1241 q 17,-41 -14,-70 L 896,678 V -64 q 0,-42 -39,-59 -13,-5 -25,-5 -27,0 -45,19 L 531,147 q -19,19 -19,45 V 678 L 19,1171 q -31,29 -14,70 17,39 59,39 h 1280 q 42,0 59,-39 z\"\n            id=\"path3017\"\n            inkscape:connector-curvature=\"0\"\n            style=\"fill:#c0c0c0;\" />\n          </g>\n      </symbol>\n\n      \n      <use *ngIf=\"flagTitle('ISOVERMOB',totalMOB).length!=0\" xlink:href=\"#flag\" [attr.x]=\"svgWidth('LD')-20\" [attr.y]=\"gridYPos(0,-3)\">\n        <title>{{flagTitle('ISOVERMOB',totalMOB)}}</title>\n      </use>\n      <use *ngIf=\"flagTitle('ISOVERDEMOB',totalDEMOB).length!=0\" xlink:href=\"#flag\" [attr.x]=\"svgWidth('LD')-20\" [attr.y]=\"gridYPos(1,-3)\">\n        <title>{{flagTitle('ISOVERDEMOB',totalDEMOB)}}</title>\n      </use>\n      <use *ngIf=\"flagTitle('ISOVERNIGHT',totalNightShift).length!=0\" xlink:href=\"#flag\" [attr.x]=\"svgWidth('LD')-20\" [attr.y]=\"gridYPos(2,-3)\">\n        <title>{{flagTitle('ISOVERNIGHT',totalNightShift)}}</title>\n      </use>\n      <use *ngIf=\"flagTitle('ISOVERPOB',totalOffshorePOB).length!=0\" xlink:href=\"#flag\" [attr.x]=\"svgWidth('LD')-20\" [attr.y]=\"gridYPos(3,-3)\">\n        <title>{{flagTitle('ISOVERPOB',totalOffshorePOB)}}</title>\n      </use>\n\n      <use *ngIf=\"getCount('FEMALE')\" xlink:href=\"#flag\" [attr.x]=\"svgWidth('LD')-20\" [attr.y]=\"gridYPos(3,-3)\">\n          <title>{{flagTitle('FEMALE',totalFemale) }}</title>\n      </use>\n        \n        <line\n          x1=\"0\" \n          [style.display]=\"idx<headerRowsIndex.length-1?'inline':'none'\"\n          [attr.y1]=\"gridYPos(idx+(idx<=6 ? 1 : 3))\" \n          [attr.x2]=\"svgWidth('LD')\" \n          [attr.y2]=\"gridYPos(idx+(idx<=6 ? 1 : 3))\" \n          stroke-width=\".5\" \n          stroke=\"#c0c0c0\"\n        *ngFor=\"let idx of headerRowsIndex;trackBy:oplandb.trackByItem;\"/>      \n\n        <line \n          [attr.x1]=\"width\" \n          [attr.y1]=\"gridYPos(7)\" \n          [attr.x2]=\"width\" \n          [attr.y2]=\"gridYPos(10)\" \n          stroke-width=\".5\" stroke=\"#c0c0c0\"\n          *ngFor=\"let width of leftHeaderWidths;let idx=index;trackBy:oplandb.trackByItem;\"\n        />\n\n        <text \n          [attr.x]=\"-135\" \n          font-stretch=\"narrower\" \n          text-anchor=\"middle\" \n          [attr.y]=\"gridYPos(2.5)\"  \n          font-family=\"Arial\" \n          [attr.font-size]=\"globalFontSize\" \n          [attr.fill]=\"globalTextFill\" \n          transform=\"rotate(-90,0,34)\"\n        >\n        ITEM #\n        </text>      \n\n    \n      \n\n\n        <!--line \n          x1=\"0\" \n          [attr.y1]=\"gridYPos(2)\" \n          [attr.x2]=\"svgWidth('LD')\" \n          [attr.y2]=\"gridYPos(2)\" \n          stroke-width=\".5\" \n          stroke=\"#c0c0c0\"\n        /-->      \n\n      <!--text [attr.x]=\"0\" \n        font-stretch=\"narrower\" \n        text-anchor=\"middle\" \n        [attr.y]=\"globalFontSize\"  \n        [attr.x]=\"contMiddle('LH')\"  \n        font-family=\"Arial\" \n        [attr.font-size]=\"globalFontSize\" \n        [attr.fill]=\"globalTextFill\" \n      >\n      BASE PLAN CHOPPER FLIGHTS\n      </text-->      \n\n    </svg>\n\n    <!--svg \n      [attr.width]=\"svgWidth('LH')\" \n      [attr.height]=\"svgHeight('LH')\" \n      [attr.viewBox]=\"svgViewBox('LH')\"\n      style=\"background-color:red;margin:none;left:0px;\" \n      xmlns=\"http://www.w3.org/2000/svg\">\n\n    </svg-->\n\n  </div> <!--left header -->\n\n  <div id=\"leftDetails\" name=\"ll_pane\"\n    style=\"z-index:2;position:absolute;left:0px;background:rgba(255, 255, 255, 0.93);\"\n    [style.margin-top]=\"contHeight('LH')\"\n    [style.width]=\"contWidth('LD')\">\n    <svg \n      [attr.width]=\"svgWidth('LD')\" \n      [attr.height]=\"svgHeight('LD')\" \n      [attr.viewBox]=\"svgViewBox('LD')\"\n      style=\"background-color:white;margin:none;left:0px;\" \n      xmlns=\"http://www.w3.org/2000/svg\">\n\n      <!-- Vertical grid lines on left details section -->\n      <line \n        [attr.x1]=\"width\" \n        [attr.y1]=\"0\" \n        [attr.x2]=\"width\" \n        [attr.y2]=\"svgHeight('LD')\" \n        stroke-width=\"0.5\" stroke=\"#c0c0c0\"\n        *ngFor=\"let width of leftHeaderWidths;let idx=index;\"\n      />\n\n      <g *ngFor=\"let core of datPersonnelsCoreNonCore;let idxCore=index;trackBy:oplandb.trackByItem;\">\n        <g *ngIf=\"core.isIncluded\">\n\n        <rect x=\"0\" [attr.y]=\"gridYPos(core.rowIndex)\" \n          [attr.width]=\"svgWidth('LD')\" \n          [attr.height]=\"gridVSpace()\" \n          stroke-width=\"0.5\"\n          stroke=\"gray\"\n          fill=\"#f2dcdb\"/>  \n          \n        <text \n          [attr.x]=\"5\" \n          font-stretch=\"narrower\" \n          text-anchor=\"start\" \n          [attr.y]=\"gridYPos(core.rowIndex+1,6)\"  \n          font-weight=\"bold\" \n          font-family=\"Arial\" \n          font-size=\"11\" \n          [attr.fill]=\"globalTextFill\">\n          TOTAL {{core['type']=='Core' ? '' : 'NON-'}}CORE CREW\n        </text>         \n\n        <!-- personnel details grouped by team -->\n        <g *ngFor=\"let team of getCoreNoneCoreTeams(core);let idxTeam=index;\">\n\n          <!--g *ngFor=\"let person of team.data;let idxPerson=index;\">\n              <rect x=\"0\" [attr.y]=\"gridYPos(idxPerson+2)\" \n              [attr.width]=\"svgWidth('LD')\" \n              [attr.height]=\"gridVSpace()\" fill=\"yellow\"/>\n          </g-->        \n\n          <rect x=\"0\" [attr.y]=\"gridYPos(team.rowIndex)\" \n            [attr.width]=\"svgWidth('LD')\" \n            [attr.height]=\"gridVSpace()\" \n            stroke-width=\"0.5\"\n            stroke=\"gray\"\n            fill=\"#FFFFCC\"/>\n\n          \n          <!-- Team name  -->\n          <svg\n            [style.cursor]=\"'pointer'\"\n            fill=\"#505050\" x=\"-2\" \n            [attr.y]=\"gridYPos(team.rowIndex,2)\" \n            height=\"20\"  width=\"24\" \n            viewBox=\"0 0 20 20\"\n            xmlns=\"http://www.w3.org/2000/svg\"\n            (click)=\"clkExpColTeam(team.type)\"\n            >\n            <title>{{(!isTeamExp(team.type) ? 'Expand ' : 'Collapse ') + team.name}}</title>\n            \n            <rect x=\"3\" y=\"3\" \n            width=\"18\"\n            height=\"19\"\n            fill=\"#FFFFCC\" title=\"hello\">\n            </rect>\n\n            <path *ngIf=\"!isTeamExp(team.type)\" \n              d=\"M16.59 8.59L12 13.17 7.41 8.59 6 10l6 6 6-6z\"/>\n            <path *ngIf=\"isTeamExp(team.type)\" \n              d=\"M12 8l-6 6 1.41 1.41L12 10.83l4.59 4.58L18 14z\"/>\n            <path d=\"M0 0h24v24H0z\" fill=\"none\"/>\n            \n          </svg>\n\n          <text \n            [attr.x]=\"25\" \n            font-stretch=\"narrower\" \n            text-anchor=\"start\" \n            [attr.y]=\"gridYPos(team.rowIndex+1,5)\"  \n            font-weight=\"bold\" \n            font-family=\"Arial\" \n            font-size=\"11\" \n            [attr.fill]=\"globalTextFill\">\n            {{team.name}}\n          </text> \n\n          <g *ngIf=\"isTeamExp(team.type)\">\n            <g *ngFor=\"let person of getTeamActivityMembers(team);trackBy:oplandb.trackByItem;let idxPerson=index;\"   \n            (click)=\"rowClicked(idxTeam , idxPerson)\" \n            (mouseenter)=\"rowEntered(idxTeam , idxPerson)\" (mouseleave)=\"hoverOff()\">\n\n              <g *ngIf=\"isNewGroup(team, person,idxPerson)\" >\n\n              <rect \n                x=\"0\" \n                [attr.y]=\"gridYPos(groupRowIndex(team,person.grp)+ idxPerson)+1\" \n                [attr.width]=\"svgWidth('LD')\" \n                [attr.height]=\"gridVSpace()-2\" \n                stroke-width=\"0.5\"\n                stroke=\"gray\"\n                fill=\"#f2f2f2\"/>\n                \n              <text\n                font-stretch=\"narrower\" \n                text-anchor=\"start\" \n                [attr.y]=\"gridYPos(groupRowIndex(team,person.grp) + 1 + idxPerson, 5)\"  \n                [attr.x]=\"5\"\n                font-family=\"Arial\" \n                [attr.font-size]=\"11\"\n                [attr.fill]=\"globalTextFill\"\n              >{{groupText(person.grp)}}\n              </text>\n                \n              </g>\n\n              <!-- Personnel Name groupRowIndex, isNewGroup(team, person,idxPerson), [attr.y]=\"gridYPos(team.rowIndex+2 + idxPerson ,5)\"  -->\n\n            <rect\n              [attr.x]=\"0\" \n              [style.opacity]=\"1\"\n              [attr.y]=\"gridYPos(groupRowIndex(team,person.grp) + 1 + idxPerson)\" \n              [attr.width]=\"svgWidth('RD')\" \n              [attr.height]=\"gridVSpace()+3\"\n              [attr.fill]=\"(idxTeam==currentTeam && idxPerson==currentRow ? '#b4b4d4' : (idxTeam==hoveredTeam && idxPerson==hoveredRow  ?'#c2e7ed':'white'))\"\n            />\n\n              <text\n                [attr.x]=\"item.x\" \n                font-stretch=\"narrower\" \n                [attr.text-anchor]=\"item.a != undefined ? item.a : 'start'\"\n                [attr.y]=\"gridYPos(groupRowIndex(team,person.grp) + 2 + idxPerson ,5)\"\n                font-family=\"Arial\" \n                font-size=\"11\" \n                style=\"user-select:normal;-webkit-user-select: normal;-moz-user-select: normal;\"\n                [attr.fill]=\"(idxTeam==currentTeam && idxPerson==currentRow ? 'white' : person.isIncluded ? globalTextFill : 'red')\"\n                *ngFor=\"let item of [{name:person.name,x:leftHeaderWidths[0]+5},\n                  {name:idxPerson+1,x:getMidPos(0, leftHeaderWidths[0]),a:'middle'}];trackBy:oplandb.trackByItem;\"\n                >\n                {{item.name}}\n              </text> \n              <!-- Company -->\n              <text\n                font-stretch=\"narrower\" \n                text-anchor=\"start\" \n                [attr.y]=\"gridYPos(groupRowIndex(team,person.grp) + 2 + idxPerson, 5)\"  \n                [attr.x]=\"leftHeaderWidths[1]+5\"\n                font-family=\"Arial\" \n                [attr.font-size]=\"11\"\n                [attr.fill]=\"(idxTeam==currentTeam && idxPerson==currentRow ? 'white' : globalTextFill)\"\n                >{{lkpCompanies['COY_'+person.coy].name}}\n              </text>\n\n              <!-- Position -->\n              <text\n                font-stretch=\"narrower\" \n                text-anchor=\"start\" \n                [attr.y]=\"gridYPos(groupRowIndex(team,person.grp) + 2 + idxPerson, 5)\"  \n                [attr.x]=\"leftHeaderWidths[2]+5\"\n                font-family=\"Arial\" \n                [attr.font-size]=\"globalFontSize\"\n                [attr.fill]=\"(idxTeam==currentTeam && idxPerson==currentRow ? 'white' : globalTextFill)\"\n                width=\"50\"\n              >{{lkpPositions['POS_'+person.pos].name}}\n              </text>\n\n            </g>\n          </g>\n\n        </g> <!-- personnel details grouped by team -->    \n\n        </g>\n      </g>\n\n      <line \n        [attr.x1]=\"0\" \n        [attr.y1]=\"svgHeight('LD')\" \n        [attr.x2]=\"svgWidth('LD')\" \n        [attr.y2]=\"svgHeight('LD')\" \n        stroke-width=\"1\" stroke=\"gray\" \n      />              \n\n\n    </svg>\n  </div>  <!--left details -->\n\n  <div id=\"rightHeader\" name=\"ur_pane\"\n    style=\"z-index:2;position:absolute;overflow:hidden;\"\n    [style.margin-left]=\"contWidth('LH')\"\n    [style.height]=\"contHeight('RH')\">\n\n      <svg \n        [attr.width]=\"svgWidth('RH')\" \n        [attr.height]=\"svgHeight('RH')\" \n        [attr.viewBox]=\"svgViewBox('RH')\"\n        style=\"background-color:white;margin:none;left:0px;\" \n          xmlns=\"http://www.w3.org/2000/svg\">\n\n        <!-- Column header section (Excess no. of Night Shift) -->\n        <!--rect x=\"0\" [attr.y]=\"gridYPos(0)+1\" \n          [attr.width]=\"svgWidth('RD')\" \n          [attr.height]=\"gridVSpace()-2\" fill=\"#b8cce4\"/-->\n\n        <!-- Column header section (Night Shift) -->\n        <!--rect x=\"0\" [attr.y]=\"gridYPos(2)+1\" \n          [attr.width]=\"svgWidth('RD')\" \n          [attr.height]=\"gridVSpace()-2\" fill=\"#ffc000\"/-->\n\n        <!-- Column header section -->\n        <!--rect x=\"0\" [attr.y]=\"gridYPos(6)\" \n          [attr.width]=\"svgWidth('RD')\" \n          [attr.height]=\"gridVSpace() * 3\" fill=\"#f2f2f2\"/--> <!-- background rectangle behind all dates -->\n\n        <g *ngFor=\"let item of datSummary;trackBy:oplandb.trackByItem;let idx=index;\">\n\n        <rect \n          style=\"cursor:pointer;\"\n          [attr.x]=\"idx * gridHSpace()\"\n          [attr.y]=\"gridYPos(7)\" \n          [attr.width]=\"gridHSpace()\" \n          [attr.height]=\"gridVSpace() * 3\" \n          [attr.fill]=\"oplandb.isNoFlight(item.date) ? '#c2e7ed' : '#f2f2f2'\"\n          (mouseenter)=\"hoverIn(item)\"\n          (mouseout)=\"hoverOut(item)\"\n          >\n          <title>{{oplandb.isNoFlight(item.date) ? 'No flight avaialble.' : ''}}</title>\n        </rect>\n          <!--//isNoFlight-->\n\n          <!--use *ngIf=\"true\" \n            xlink:href=\"#filter\" [attr.x]=\"idx * gridHSpace()\" [attr.y]=[attr.y]=\"gridYPos(6)\">\n          </use-->\n\n          <!--\n\n                  <div class=\"btn-group dropdown\">\n                    <button type=\"button\" \n                      class=\"btn p-0 px-2 text-primary\" data-toggle=\"dropdown\" \n                      title=\"Click to select company...\"\n                      style=\"cursor: pointer;border-radius:2px;border:1px solid #c0c0c0;\"\n                      aria-haspopup=\"true\" aria-expanded=\"false\"\n                      >\n                      <i class=\"fa fa-list-alt\"></i>\n                    </button>\n    \n                    <app-ctl-drop-down *ngIf=\"this.loadLookups\" class=\"dropdown-menu p-0\"\n                      [drpdData]=\"oplandb.datCompany\"\n                      [(drpdValue)]=\"oplandb.newActivityMember.coyID\" \n                      [drpdDispRows]=\"15\"\n                      \n                      style=\"z-index:100;left:-150px;\"\n                    ></app-ctl-drop-down>          \n            \n\n          -->\n          <svg *ngIf=\"isSameDate(udf.calFilterDate,item.date) || isSameDate(hoveredDate,item.date)\"  width=\"15\" height=\"15\" viewBox=\"0 0 1500 1500\"\n            [attr.x]=\"idx * gridHSpace() + gridHSpace()/2 - 10\"  \n            [attr.y]=\"gridYPos(9,-6)\" style=\"cursor:pointer;\"\n            (mousenter)=\"hoverIn(item)\"\n            >\n            <use (click)=\"applyRemoveFilter(item)\" *ngIf=\"true\" \n            [attr.xlink:href]=\"isSameDate(udf.calFilterDate,item.date) ? '#filter' : '#filter_gray'\"  >\n              <title>\n                {{ isSameDate(udf.calFilterDate,item.date) ? 'Click to remove personnel activity filter' : 'Click to apply personnel activity filter on ' + item.date }}  \n              </title>\n            </use>\n\n          </svg>\n              \n          <text [attr.x]=\"idx * gridHSpace() + gridHSpace()/2\" \n            (click)=\"dateClicked($event)\"\n            font-stretch=\"narrower\" \n            text-anchor=\"middle\" \n            [attr.y]=\"gridYPos(9,5)-10\" \n            font-family=\"Arial\" \n            font-size=\"11\" \n            [attr.fill]=\"globalTextFill\"\n            >{{ item.date }}\n          </text>\n\n          <text [attr.x]=\"idx * gridHSpace() + gridHSpace()/2\" \n            font-stretch=\"narrower\" \n            text-anchor=\"middle\" \n            [attr.y]=\"gridYPos(9,9)+10\"  \n            font-family=\"Arial\" \n            font-size=\"11\" \n            [attr.fill]=\"globalTextFill\"\n            >{{item.DOW}}\n          </text>\n\n          <!-- Total mobilization on details column header section -->\n          <g *ngIf=\"totalMOB[idx]\">\n              <!--rect  \n                [attr.x]=\"idx * gridHSpace()\" \n                [attr.y]=\"gridYPos(0)+1\" \n                [attr.width]=\"gridHSpace()\" \n                [attr.height]=\"gridVSpace()-2\" \n                fill=\"#f79646\"/-->  \n                    \n              <text\n                [attr.x]=\"idx * gridHSpace() + gridHSpace()/2\" \n                font-stretch=\"narrower\" \n                text-anchor=\"middle\" \n                [attr.y]=\"gridYPos(1,6)\"  \n                font-family=\"Arial\" \n                font-size=\"11\" \n                font-weight=\"bold\"\n                [attr.fill]=\"getCount('MOBLIMIT',idx) < totalMOB[idx] ? 'red' : 'black'\">\n                {{totalMOB[idx]}}\n              </text>            \n            </g> <!-- Mobilization -->\n\n          <!-- Total demobilization on details column header section -->\n          <g *ngIf=\"totalDEMOB[idx]\">\n\n            <!--rect  \n              [attr.x]=\"idx * gridHSpace()\" \n              [attr.y]=\"gridYPos(1)+1\" \n              [attr.width]=\"gridHSpace()\" \n              [attr.height]=\"gridVSpace()-2\" \n              fill=\"#f79646\"/-->  \n\n            <text \n              [attr.x]=\"idx * gridHSpace() + gridHSpace()/2\" \n              font-stretch=\"narrower\" \n              text-anchor=\"middle\" \n              [attr.y]=\"gridYPos(2,6)\"  \n              font-family=\"Arial\" \n              font-size=\"11\" \n              font-weight=\"bold\"\n              [attr.fill]=\"getCount('DEMOBLIMIT',idx) < totalDEMOB[idx] ? 'red' : 'black'\">\n              {{totalDEMOB[idx]}}\n            </text>     \n          </g>\n\n\n        <!-- [No. of POB] - [Night Shift] \n          <text *ngIf=\"totalOffshorePOB[idx]-totalNightShift[idx]\"\n          [attr.x]=\"idx * gridHSpace() + gridHSpace()/2\" \n          font-stretch=\"narrower\" \n          text-anchor=\"middle\" \n          [attr.y]=\"gridYPos(2,5)\"  \n          font-family=\"Arial\" \n          font-size=\"11\" \n          font-weight=\"bold\"\n          fill=\"red\">\n          {{totalOffshorePOB[idx]-totalNightShift[idx]}}\n        </text>             \n        -->\n\n          <!-- Night Shift -->\n          <text *ngIf=\"totalNightShift[idx]!=0 || (getCount('NIGHTLIMIT',idx) > totalNightShift[idx])\"\n            [attr.x]=\"idx * gridHSpace() + gridHSpace()/2\" \n            font-stretch=\"narrower\" \n            text-anchor=\"middle\" \n            [attr.y]=\"gridYPos(3,5)\"  \n            font-family=\"Arial\" \n            font-size=\"11\" \n            font-weight=\"bold\"\n            [attr.fill]=\"getCount('NIGHTLIMIT',idx) > totalNightShift[idx] ? 'red' : 'black'\">\n            {{getCount('NIGHTLIMIT',idx)==0 ? totalNightShift[idx] : getCount('NIGHTLIMIT',idx)-totalNightShift[idx]}}\n\n            \n          </text>             \n              \n\n          <!-- #of NS crew\n        <text *ngIf=\"totalOffshorePOB[idx]-bedsSummary[idx]\"\n          [attr.x]=\"idx * gridHSpace() + gridHSpace()/2\" \n          font-stretch=\"narrower\" \n          text-anchor=\"middle\" \n          [attr.y]=\"gridYPos(4,5)\"  \n          font-family=\"Arial\" \n          font-size=\"11\" \n          font-weight=\"bold\"\n          fill=\"black\">\n          {{totalOffshorePOB[idx] - bedsSummary[idx]}}       \n        </text>             \n        -->\n\n          <!-- STD POB -->\n          <!--text *ngIf=\"totalOffshorePOB[idx]-upManTeamSummary[idx]\"\n            [attr.x]=\"idx * gridHSpace() + gridHSpace()/2\" \n            font-stretch=\"narrower\" \n            text-anchor=\"middle\" \n            [attr.y]=\"gridYPos(5,5)\"  \n            font-family=\"Arial\" \n            font-size=\"11\" \n            font-weight=\"bold\"\n            fill=\"red\">\n            {{totalOffshorePOB[idx] - upManTeamSummary[idx]}}       \n          </text-->             \n        <!--\n\n          {{(totalOffshorePOB[idx]-bedsSummary[idx])<0 ? '( '+(bedsSummary[idx]-totalOffshorePOB[idx])+' )' : totalOffshorePOB[idx]-bedsSummary[idx] }}\n\n        -->\n\n          <!-- UPMANNING POB -->\n          <!--text *ngIf=\"upManTeamSummary[idx]\"\n            [attr.x]=\"idx * gridHSpace() + gridHSpace()/2\" \n            font-stretch=\"narrower\" \n            text-anchor=\"middle\" \n            [attr.y]=\"gridYPos(6,5)\"  \n            font-family=\"Arial\" \n            font-size=\"11\" \n            font-weight=\"bold\"\n            fill=\"black\">\n            {{upManTeamSummary[idx]}}\n          </text-->             \n                \n\n\n          <!-- Maximum rooms on details column header section -->\n          <text [attr.x]=\"idx * gridHSpace() + gridHSpace()/2\" \n            font-stretch=\"narrower\" \n            text-anchor=\"middle\" \n            [attr.y]=\"gridYPos(6,5)\"  \n            font-family=\"Arial\" \n            font-size=\"11\" \n            font-weight=\"normal\"\n            [attr.fill]=\"globalTextFill\"\n            >\n            {{getCount('POBLIMIT',idx)}}\n          </text>  \n\n          <!--bedsSummary[idx]-->\n          \n          \n\n          <!-- Up Manning Limit -->\n          <!--text *ngIf=\"upManSummary[idx]\"\n          [attr.x]=\"idx * gridHSpace() + gridHSpace()/2\" \n          font-stretch=\"narrower\" \n          text-anchor=\"middle\" \n          [attr.y]=\"gridYPos(13,5)\"  \n          font-family=\"Arial\" \n          font-size=\"11\" \n          font-weight=\"bold\"\n          fill=\"black\">\n          {{upManSummary[idx]}}\n        </text-->             \n      \n        \n\n          <!--bedsSummary[idx] Maximum rooms on details column header section>\n          <text [attr.x]=\"idx * gridHSpace() + gridHSpace()/2\" \n            font-stretch=\"narrower\" \n            text-anchor=\"middle\" \n            [attr.y]=\"gridYPos(13,5)\"  \n            font-family=\"Arial\" \n            font-size=\"11\" \n            [attr.fill]=\"globalTextFill\">\n            {{upManSummary[idx]}}\n          </text-->    \n\n          <!-- TOTAL OFFSHORE POB -->\n          <text \n            [attr.x]=\"idx * gridHSpace() + gridHSpace()/2\" \n            font-stretch=\"narrower\" \n            text-anchor=\"middle\" \n            [attr.y]=\"gridYPos(5,5)\"  \n            font-family=\"Arial\" \n            font-size=\"11\" \n            font-weight=\"bold\"\n            [attr.fill]=\"totalOffshorePOB[idx] > getCount('POBLIMIT',idx)? 'red' : 'black'\">\n            {{totalOffshorePOB[idx]}}\n          </text>   \n\n          <!-- FEMALE CREW -->\n          <text \n            [attr.x]=\"idx * gridHSpace() + gridHSpace()/2\" \n            font-stretch=\"narrower\" \n            text-anchor=\"middle\" \n            [attr.y]=\"gridYPos(4,5)\"  \n            font-family=\"Arial\" \n            font-size=\"11\" \n            [attr.fill]=\"totalFemale[idx] % 2 ? 'red' : globalTextFill\">\n            {{totalFemale[idx]}}\n          </text>             \n\n      <!-- Vertical maximum rooms on details column header section -->\n      <text\n        [attr.x]=\"idx * gridHSpace() + gridHSpace()/2\" \n        font-stretch=\"narrower\" \n        text-anchor=\"middle\" \n        [attr.y]=\"gridYPos(7,5)\"  \n        font-family=\"Arial\" \n        font-size=\"11\" \n        [attr.fill]=\"globalTextFill\">\n        {{getCount('POBLIMIT',idx)-totalOffshorePOB[idx]}}\n      </text>   \n      \n        </g>\n\n          <!--text \n          [attr.x]=\"leftHeaderWidths[3] + (leftHeaderWidths[2]-leftHeaderWidths[3])/2\" \n          font-stretch=\"narrower\" \n          text-anchor=\"middle\" \n          [attr.y]=\"gridYPos(5,5)\"  \n          font-family=\"Arial\" \n          [attr.font-size]=\"globalFontSize\" \n          [attr.fill]=\"globalTextFill\" \n        >\n        {{getDateDisplay(idx)}}\n        </text-->\n\n\n          <!-- right details header date separator line-->\n          <!--line x1=\"0\" [attr.y1]=\"posHGridOffset\" [attr.x2]=\"canvasWidth()\" [attr.y2]=\"posHGridOffset\" stroke-width=\"0.5\" stroke=\"#c0c0c0\" /-->\n\n          <!-- right details header separator line-->\n          <!--line x1=\"0\" [attr.y1]=\"headSeparatorPos-0.5\" [attr.x2]=\"canvasWidth()\" [attr.y2]=\"headSeparatorPos-0.5\" stroke-width=\"1\" stroke=\"#c0c0c0\" /-->\n\n          <!-- Horizontal grid on header section -->\n          <line \n            x1=\"0\" \n            [style.display]=\"hIdx<(headerRowsIndex.length-1)?'inline':'none'\"\n            [attr.y1]=\"gridYPos(idx+(idx<=6 ? 1 : 3))\" \n            [attr.x2]=\"svgWidth('RD')\" \n            [attr.y2]=\"gridYPos(idx+(idx<=6 ? 1 : 3))\" \n            stroke-width=\".5\" \n            stroke=\"#c0c0c0\"\n          *ngFor=\"let idx of headerRowsIndex;let hIdx=index\"/>            \n\n            \n          <!-- Vertical grid on header section -->\n          <line [attr.x1]=\"idx * gridHSpace()\" \n            y1=\"0\" [attr.x2]=\"idx * gridHSpace()\" \n            y2=\"100vh\" stroke-width=\".5\" stroke=\"#c0c0c0\" \n            *ngFor=\"let idx of indexArray\"/>\n\n          <line [attr.x1]=\"indexArray.length * gridHSpace()\" \n            y1=\"0\" [attr.x2]=\"indexArray.length * gridHSpace()\" \n            y2=\"100vh\" stroke-width=\".5\" stroke=\"#c0c0c0\"/>\n      \n        </svg>\n      \n  </div> <!-- right header *ngIf=\"hovObjSrc\" -->\n\n  <div class=\"btn btn-link p-0 m-0\" *ngIf=\"withUndo()\" (click)=\"undoMove()\"\n    style=\"position:absolute;z-index:74;border-radius: 0px;color:white;cursor: pointer;height:18px;width:18px;\"\n    [style.top.px]=\"dragHandlePosition('top',-5 + dragScrollY)\"\n    [style.left.px]=\"dragHandlePosition('left',gridHSpace()-15 + dragScrollX)\"\n    title=\"Undo assignment date changes ...\"\n    >\n    <span><i class=\"fa fa-undo\" style=\"font-size: 0.9em;\"></i></span>\n</div>\n\n<div id=\"hlMember\" *ngIf=\"this.hovObjSrc && oplandb.showHilights\" \n  style=\"z-index:72;background: lime;opacity: 0.2;position: absolute;\"\n  [style.height.px]=\"gridVSpace() * 3\"\n  [style.width.px]=\"hovObjWidth()\"\n  [style.top.px]=\"gridVSpace() * 6 + dragScrollY\"\n  [style.left.px]=\"hovObjLeft()\"\n  >\n  &nbsp;\n</div>\n<div id=\"hlActivity\"  *ngIf=\"this.hovObjSrc && oplandb.showHilights\" \n  style=\"z-index:5;background: blue;opacity: 0.2;position: absolute;\"\n    [style.height.px]=\"gridVSpace() * 3\"\n    [style.width.px]=\"hovObjActWidth()\"\n    [style.top.px]=\"gridVSpace() * 6 + dragScrollY\"\n    [style.left.px]=\"hovObjActLeft()\"\n  >\n  &nbsp;\n</div>\n\n  <div id=\"dragHandle\" *ngIf=\"ie()\"\n    class=\"d-flex justify-content-between p-0\" \n    style=\"z-index:73;background: black;opacity:0.35;position: absolute;\"\n    [style.cursor]=\"curMove()\"\n    [style.visibility]=\"hovObjSrc ? 'visible' : 'hidden'\"\n    \n    [style.height.px]=\"gridVSpace()\"\n    [style.width.px]=\"300\"\n    title=\"Make activity assignment earlier or later...\"\n    (mousedown)=\"mouseEvent($event)\"\n  >\n    <div id=\"dragLeft\" (mousedown)=\"mouseEvent($event)\"\n      title=\"Make activity assgnment earlier and longer...\"\n      [style.cursor]=\"curLeft()\"\n      style=\"background:black;color:white;width:8px;\">&nbsp;</div>\n    \n    <div id=\"dragRight\" (mousedown)=\"mouseEvent($event)\"\n      title=\"Make activity assgnment longer...\"\n      [style.cursor]=\"curLeft()\"\n      style=\"background:black;color:white;width:8px;\">&nbsp;</div>\n      \n  </div>\n\n  <div id=\"dragMask\" *ngIf=\"hovObj\"\n    style=\"z-index:75;position:absolute;background:white;white-space:normal;opacity: 0.02;\"\n    [style.margin-top]=\"contHeight('RH')\"\n    [style.margin-left]=\"contWidth('LD')\"\n    [style.width.px]=\"svgWidth('RD')\"\n    [style.height.px]=\"svgHeight('RD')\"\n    [style.cursor]=\"!dragValid() ? 'no-drop' : (dragMode=='move' ? 'move' : (dragMode=='e-resize' || dragMode=='w-resize' ? 'ew-resize' : 'default'))\"\n\n    (mousemove)=\"mouseEvent($event)\"\n    (mouseup)=\"mouseEvent($event)\"\n  >\n    &nbsp;\n  </div>\n\n  <div id=\"rightDetails\" name=\"lr_pane\"\n    style=\"z-index:1;position:absolute;background:white;white-space:normal;\"\n    [style.margin-top]=\"contHeight('RH')\"\n    [style.margin-left]=\"contWidth('LD')\"\n    >\n\n    <svg \n      [attr.width]=\"svgWidth('RD')\" \n      [attr.height]=\"svgHeight('RD')\" \n      [attr.viewBox]=\"svgViewBox('RD')\"\n      style=\"background-color:white;position:absolute;margin:none;left:0px;\" \n      xmlns=\"http://www.w3.org/2000/svg\">\n      <!-- Row header section (left side details - names/shifts) -->\n\n      <g *ngFor=\"let core of datPersonnelsCoreNonCore;let idxCore=index;trackBy:oplandb.trackByItem;\">\n\n        <rect x=\"0\" [attr.y]=\"gridYPos(core.rowIndex)\" \n          [attr.width]=\"svgWidth('RD')\" \n          [attr.height]=\"gridVSpace()\" \n          stroke-width=\"0.5\"\n          stroke=\"gray\"\n          fill=\"#f2dcdb\"/>  \n          \n        <g *ngFor=\"let idx of indexArray\">\n          <!-- create text for each column with summation value -->\n          <text *ngIf=\"core.coreSummary[idx]\"\n            [attr.x]=\"idx * gridHSpace() + gridHSpace()/2\" \n            font-stretch=\"narrower\" \n            text-anchor=\"middle\" \n            [attr.y]=\"gridYPos(core.rowIndex+1,5)\"  \n            font-family=\"Arial\" \n            font-size=\"11\" \n            [attr.fill]=\"globalTextFill\">\n            {{core.coreSummary[idx]}}\n          </text>        \n        </g>\n\n        <!-- personnel details grouped by team -->\n        <g *ngFor=\"let team of getCoreNoneCoreTeams(core);let idxTeam=index;trackBy:oplandb.trackByItem;\">\n\n          <rect x=\"0\" [attr.y]=\"gridYPos(team.rowIndex)\" \n            [attr.width]=\"svgWidth('RD')\" \n            [attr.height]=\"gridVSpace()\" \n            stroke-width=\"0.5\"\n            stroke=\"gray\"\n            fill=\"#FFFFCC\"/>                \n\n          <g *ngFor=\"let idx of indexArray;trackBy:oplandb.trackByItem;\">\n            <!-- create text for each column with summation value -->\n            <text *ngIf=\"team.teamSummary[idx]\"\n              [attr.x]=\"idx * gridHSpace() + gridHSpace()/2\" \n              font-stretch=\"narrower\" \n              text-anchor=\"middle\" \n              [attr.y]=\"gridYPos(team.rowIndex+1,5)\"  \n              font-family=\"Arial\" \n              font-size=\"11\" \n              [attr.fill]=\"globalTextFill\">\n              {{team.teamSummary[idx]}}\n            </text>        \n          </g>\n\n          <!--text \n            [attr.x]=\"5\" \n            font-stretch=\"narrower\" \n            text-anchor=\"start\" \n            [attr.y]=\"gridYPos(team.rowIndex+1,5)\"  \n            font-family=\"Arial\" \n            font-size=\"11\" \n            [attr.fill]=\"globalTextFill\">\n            {{team.name}}\n          </text--> \n\n          \n\n          <!-- Vertical grid on details section -->\n\n          <g *ngIf=\"isTeamExp(team.type)\" >\n\n            <!--g *ngFor=\"let person of getTeamActivityMembers(team);trackBy:oplandb.trackByItem;let idxPerson=index;\"  (click)=\"rowClicked(idxTeam , idxPerson)\"  (mouseenter)=\"rowEntered(idxTeam , idxPerson)\" (mouseleave)=\"hoverOff()\"-->\n            <g *ngFor=\"let person of getTeamActivityMembers(team);trackBy:oplandb.trackByItem;let idxPerson=index;\"  (click)=\"rowClicked(idxTeam , idxPerson)\"  (mouseenter)=\"rowEntered(idxTeam , idxPerson)\" (mouseleave)=\"hoverOff()\">\n\n            <!-- for hilight on click start -->\n              <!--rect  \n                [attr.x]=\"0\" \n                [style.opacity]=\"1\"\n                [attr.y]=\"gridYPos(groupRowIndex(team,person.grp) + 1 + idxPerson)\" \n                [attr.width]=\"svgWidth('RD')\" \n                [attr.height]=\"gridVSpace()\"\n                [attr.fill]=\"'white'\"\n              /-->              \n              <rect\n              [attr.x]=\"0\" \n              [style.opacity]=\"1\"\n              [attr.y]=\"gridYPos(groupRowIndex(team,person.grp) + 1 + idxPerson)\" \n              [attr.width]=\"svgWidth('RD')\" \n              [attr.height]=\"gridVSpace()+3\"\n              [attr.fill]=\"(idxTeam==currentTeam && idxPerson==currentRow ? '#b4b4d4' : (idxTeam==hoveredTeam && idxPerson==hoveredRow  ?'#c2e7ed':'white'))\"\n            />\n            <!-- for hilight on click end -->            \n          \n\n              <!--text \n                [attr.x]=\"25\" \n                font-stretch=\"narrower\" \n                text-anchor=\"start\" \n                [attr.y]=\"gridYPos(team.rowIndex+2 + idxPerson ,5)\"  \n                font-family=\"Arial\" \n                font-size=\"11\" \n                [attr.fill]=\"globalTextFill\">\n                {{person.name}}\n              </text--> \n\n              <g *ngIf=\"isNewGroup(team, person,idxPerson)\" >\n    \n                <rect \n                  x=\"0\" \n                  [attr.y]=\"gridYPos(groupRowIndex(team,person.grp)+ idxPerson)+1\" \n                  [attr.width]=\"svgWidth('RD')\" \n                  [attr.height]=\"gridVSpace()-2\" \n                  stroke-width=\"0.5\"\n                  stroke=\"gray\"\n                  fill=\"#f2f2f2\"/>\n                  \n                <!--text\n                  font-stretch=\"narrower\" \n                  text-anchor=\"start\" \n                  [attr.y]=\"gridYPos(groupRowIndex(team,person.grp) + 1 + idxPerson, 5)\"  \n                  [attr.x]=\"5\"\n                  font-family=\"Arial\" \n                  [attr.font-size]=\"globalFontSize\"\n                  [attr.fill]=\"globalTextFill\"\n                >{{groupText(person.grp)}}\n                </text-->\n                  \n                </g>            \n\n              <g *ngFor=\"let act of person.activities;let actIdx=index;trackBy:oplandb.trackByItem;\">\n\n                <!-- activity bar (draggable bar)-->\n                <rect id=\"mem_act_{{act.id}}_{{act.acpID}}\"\n                  [style.cursor]=\"ie() ? 'default' : 'pointer'\"\n\n                  [attr.x]=\"dateDiff(udf.calStartDate,act.startDate) * gridHSpace()\" \n                  [attr.y]=\"gridYPos(groupRowIndex(team,person.grp) + 1 + idxPerson) + 2\" \n\n                  [attr.width]=\"(act.rectSpan==1 ? 1 + (act.demob && !act.isDay ? 1 : 0) :(act.rectSpan  + (act.demob ? 1 : 0))) * gridHSpace()\" \n                  [attr.height]=\"gridVSpace()-4\"\n                  [attr.fill]=\"act.noPOB ? (!act.sdBackup ? '#d042ea' : modNonPOBBack) : (!act.sdBackup ? '#30e030' : modPOBBack)\"\n\n                  (click)=\"editActivityMember(act,person)\"\n                  (mousedown)=\"mouseEvent($event)\"\n                  (mouseup)=\"mouseEvent($event)\"\n                  [attr.data-toggle]=\"oplandb.logInfo.activityMemAdmin && !oplandb.isEditing? 'modal' : ''\"\n                  [attr.data-target]=\"oplandb.logInfo.activityMemAdmin && !oplandb.isEditing ? '#exampleModal' : ''\"\n\n                  (mouseenter)=\"mouseEvent($event)\"\n                  (mouseleave)=\"mouseEvent($event)\"\n                  \n                >\n                <title>\n                  {{ getActivityBarTitle(act,person) }} \n                </title>\n              </rect>\n                \n                <g *ngFor=\"let cIdx of fillArrNum(act.rectSpan)\">\n                  <!--\n                    display circle for all days of activity except for cells with\n                    mob/demob text,\n                    display only on a single-day activity where mob and demob is false\n\n                    modified 1/29/2018 to include  single day assignment !mob && demob\n                  -->\n\n                  <circle *ngIf=\"!act.noPOB && ((cIdx > 0 || (cIdx == 0 && !act.mob)) && (!(act.demob && act.rectSpan==1) || (act.demob && !act.mob && act.rectSpan==1)) )\"\n                    [attr.cx]=\"(dateDiff(udf.calStartDate,act.startDate) * gridHSpace()) + gridHSpace() * cIdx + gridHSpace()/2\" \n                    [attr.cy]=\"gridYPos(groupRowIndex(team,person.grp) + 1 + idxPerson) + gridVSpace()/2\" \n                    [attr.fill]=\"act.noPOB ? 'teal' : 'blue'\"\n                    stroke-width=\"2\"\n                    stroke=\"#c0c0c0\"\n                    r=\"4.5\"\n                  />\n                </g>\n      \n                <text *ngIf=\"act.mob\" \n                  [attr.x]=\"dateDiff(udf.calStartDate,act.startDate) * gridHSpace() + gridHSpace()/2\" \n                  font-stretch=\"narrower\" \n                  text-anchor=\"middle\" \n                  [attr.y]=\"gridYPos(groupRowIndex(team,person.grp) + 2 + idxPerson, 6)\"  \n                  font-family=\"Arial\" \n                  font-size=\"11\" \n                  font-weight=\"bold\" \n                  fill=\"white\"\n                  >\n                  {{isDayTrip(act) ? 'DAY' :'MOB'}}\n                </text>                  \n      \n                <text *ngIf=\"act.demob && (!(act.mob && act.rectSpan==1) || (act.mob && act.rectSpan==1 && !act.isDay))\" \n                  [attr.x]=\"(dateDiff(udf.calStartDate,act.endDate)+(act.rectSpan!=1 || (act.rectSpan==1 && !act.isDay) ? 1 : 0)) * gridHSpace() + gridHSpace()/2\" \n                  font-stretch=\"narrower\" \n                  text-anchor=\"middle\" \n                  [attr.y]=\"gridYPos(groupRowIndex(team,person.grp) + 2 + idxPerson, 6)\"  \n                  font-family=\"Arial\" \n                  font-size=\"11\" \n                  fill=\"white\">\n                  DEMOB\n                </text> \n\n              </g>          \n\n            </g>\n\n          </g>\n\n        </g> <!-- personnel details grouped by team -->    \n\n      </g>\n      \n      <line \n        [attr.x1]=\"idx * gridHSpace()\" \n        [attr.y1]=\"0\" \n        [attr.x2]=\"idx * gridHSpace()\" \n        [attr.y2]=\"svgHeight('RD')\" \n        stroke-width=\".5\" stroke=\"gray\" \n        *ngFor=\"let idx of indexArray\"\n      />        \n      \n      <line \n        [attr.x1]=\"0\" \n        [attr.y1]=\"svgHeight('RD')\" \n        [attr.x2]=\"svgWidth('RD')\" \n        [attr.y2]=\"svgHeight('RD')\" \n        stroke-width=\"1\" stroke=\"gray\" \n      />              \n    \n\n\n      <!-- details border -->\n      <!--rect x=\"0\" y=\"0\" \n        [attr.width]=\"svgWidth('RD')\" \n        [attr.height]=\"svgHeight('RD')\"\n        stroke-width=\"1\"\n        stroke=\"#f2f2f2\"\n        fill=\"transparent\"\n      /-->\n      \n\n\n    </svg>    \n\n  </div>\n\n</div> <!-- main scrollable wrapper -->\n\n\n\n<div *ngIf=\"false\">\n<br/>\n<h5>{{gridYPos(1) +', ' + dateSpan}}</h5>\n<br/>\n<input type=\"text\" [(ngModel)]=\"posHGridOffset\" name=\"posHGridOffset\">\n</div>\n\n\n\n\n  <!-- Modal -->\n  <div class=\"modal fade\" id=\"exampleModal\" tabindex=\"-1\" \n    role=\"dialog\" aria-labelledby=\"exampleModalLabel\" aria-hidden=\"true\">\n    <div class=\"modal-dialog\" role=\"document\">\n      <div class=\"modal-content\">\n        <div class=\"modal-header\">\n          <h5 class=\"modal-title\" id=\"exampleModalLabel\">Activity Member</h5>\n          <button type=\"button\" \n            class=\"close\" data-dismiss=\"modal\" \n             aria-label=\"Close\">\n            <span aria-hidden=\"true\">&times;</span>\n          </button>\n        </div>\n        <div class=\"modal-body\" style=\"position:relative;font-size:0.9em\">\n\n          <div *ngIf=\"oplandb.msgTempPromptModal.length>0\"class=\"alert alert-danger row\"\n            role=\"alert\"\n            style=\"position: absolute;z-index:20;top:0px;left:15px;width:100%;\">\n            <strong>Reminder:</strong>  {{oplandb.msgTempPromptModal}}\n          </div>\n\n\n            <div class=\"row px-3\">\n                <div class=\"form-group\" >\n                  <label class=\"col-form-label my-0 py-0\">\n                    {{displayLabel()}}\n                  </label><br>\n                  <app-ctl-user-select *ngIf=\"oplandb.datPersonnel.length!=0\" \n                  [width]=\"468\"\n                  [height]=\"250\" \n                  [(counter)]=\"oplandb.newActivityMember.perID\" \n      \n                  [lookup]=\"oplandb.idxPersonnel\"\n                  [displayField]=\"'fullName'\"\n                  [valueField]=\"'id'\"\n\n                  [disableSelect]=\"false\"\n      \n                  [source]=\"oplandb.datTeams\"\n                  [groupName]=\"'name'\"\n                  [dataName]=\"'members'\"\n                  [valueColumn]=\"'perID'\"\n                  [displayColumn]=\"''\"\n                  (counterChange)=\"onUserChange($event)\"\n                  (itemSelected)=\"onUserItemSelected($event)\"\n                  >\n              </app-ctl-user-select>        \n\n                </div>  \n            </div>\n  <!--[disableSelect]=\"disableControls() || udf.isActivityMemberEdit()\"-->\n        \n            <div class=\"row px-3\">\n\n              <label class=\"col-form-label\" style=\"margin-top:0px;padding-top:0px;\">Functional Sponsor / Company</label>\n\n              <div class=\"input-group input-group\" >\n                  <input type=\"text\" class=\"form-control form-control-sm\" id=\"firstName\"  \n                    [value]=\"oplandb.getCompanyName(oplandb.newActivityMember.coyID)\" \n                    onfocus=\"this.blur()\"\n                  >           \n      \n                <!-- custom dropdown use \"btn-group dropdown\" class [start] -->\n                <div class=\"btn-group dropdown\">\n                  <button type=\"button\" \n                    class=\"btn p-0 px-2 text-primary\" data-toggle=\"dropdown\" \n                    title=\"Click to select company...\"\n                    style=\"cursor: pointer;border-radius:2px;border:1px solid #c0c0c0;\"\n                    aria-haspopup=\"true\" aria-expanded=\"false\"\n                    >\n                    <i class=\"fa fa-list-alt\"></i>\n                  </button>\n  \n                  <app-ctl-drop-down *ngIf=\"this.loadLookups\" class=\"dropdown-menu p-0\"\n                    [drpdData]=\"oplandb.datCompany\"\n                    [(drpdValue)]=\"oplandb.newActivityMember.coyID\" \n                    [drpdDispRows]=\"15\"\n                    \n                    style=\"z-index:100;left:-150px;\"\n                  ></app-ctl-drop-down>          \n                  \n                </div> \n                <!-- custom dropdown use \"btn-group dropdown\" class [end] -->\n      \n              </div>\n\n              <!--select *ngIf=\"loadLookups\" class=\"form-control form-control-sm\"  (onload)=\"selectOnload()\"\n                id=\"exampleSelect1\" \n                [disabled]=\"isCoreTeam(oplandb.newActivityMember.teamID)\"\n                [(ngModel)]=\"oplandb.newActivityMember.coyID\" \n                name=\"coyID\" >\n                <option *ngFor=\"let coy of oplandb.datCompany\" [ngValue]=\"coy.id\">{{ coy.name }}</option>\n              </select-->      \n\n            </div>\n\n            <div class=\"row px-3\">\n              \n              <label class=\"col-form-label\">Position</label>\n\n              <div class=\"input-group input-group\" >\n                  <input type=\"text\" class=\"form-control form-control-sm\" id=\"firstName\"  \n                    onfocus=\"this.blur()\"\n                    [value]=\"oplandb.getPositionName(oplandb.newActivityMember.posID)\" \n                  >           \n      \n                <!-- custom dropdown use \"btn-group dropdown\" class [start] -->\n                <div class=\"btn-group dropdown\">\n                  <button type=\"button\" \n                    class=\"btn p-0 px-2 text-primary\" data-toggle=\"dropdown\" \n                    title=\"Click to select position...\"\n                    style=\"cursor: pointer;border-radius:2px;border:1px solid #c0c0c0;\"\n                    aria-haspopup=\"true\" aria-expanded=\"false\"\n                    >\n                    <i class=\"fa fa-list-alt\"></i>\n                  </button>\n\n                  <app-ctl-drop-down *ngIf=\"this.loadLookups\" class=\"dropdown-menu p-0\"\n                    [drpdData]=\"oplandb.datPosition\"\n                    [(drpdValue)]=\"oplandb.newActivityMember.posID\" \n                    [drpdDispRows]=\"15\"\n                    \n                    style=\"z-index:100;left:-150px;\"\n                  ></app-ctl-drop-down>          \n\n                  \n                </div> \n                <!-- custom dropdown use \"btn-group dropdown\" class [end] -->\n      \n              </div>\n              \n\n              <!--select *ngIf=\"loadLookups\" class=\"form-control form-control-sm\" \n                id=\"exampleSelect1\" \n                [disabled]=\"isCoreTeam(oplandb.newActivityMember.teamID)\"\n                [(ngModel)]=\"oplandb.newActivityMember.posID\" \n                name=\"posID\" >\n                <option *ngFor=\"let pos of oplandb.datPosition\" [ngValue]=\"pos.id\">{{ pos.name }}</option>\n              </select-->      \n\n            </div>\n             <!-- start of activity member fields 8888 -->\n              <div class=\"row px-3\" style=\"display:flex;justify-content:space-between;padding-top:5px;\">\n                <div class=\"form-group\">\n                  <label class=\"col-form-label my-0 py-0\">Start</label>\n                  <i *ngIf=\"oplandb.isActMemNoFlight(oplandb.newActivityMember,false).noFlight\" \n                  class=\"fa fa-plane text-danger p-0 m-0\" \n                  [title]=\"'No flight available on mob date ' + oplandb.isActMemNoFlight(oplandb.newActivityMember,false).date +'.'\" \n                  style=\"font-size:1.2em;cursor:pointer;margin-left:10px;\"></i>\n                  \n                  <!--td [style.width]=\"actsColsPX[1]+'px'\" style=\"text-align:center;\">\n                      {{ mem.startDate }}\n                      &nbsp;<i *ngIf=\"oplandb.isActMemNoFlight(mem,false).noFlight\" \n                      class=\"fa fa-plane text-danger p-0 m-0\" \n                      [title]=\"'No flight available on mob date ' + oplandb.isActMemNoFlight(mem,false).date +'.'\" \n                      style=\"font-size:1.2em;cursor:pointer;\"></i-->\n\n                    <div class=\"input-group input-group-sm\" >\n                      <input class=\"form-control text-center\" \n                        [dpDayPicker]=\"udf.datePickerConfig\"\n                        [(ngModel)]=\"oplandb.newActivityMember.startDate\"\n                        id=\"memStart\"\n                        onfocus=\"this.blur()\"\n                        name=\"assigmentStartDate\"\n                        type=\"text\"\n                        style=\"width:170px;\"\n                        >\n                        <div class=\"input-group-addon mx-0\">\n                            <a href=\"#\" (click)=\"openCalendar('memStart')\"><i class=\"fa fa-calendar text-primary\"></i></a>\n                        </div>\n              </div>                    \n                </div> \n                \n                <div class=\"form-group\">\n                  <label class=\"col-form-label py-0 my-0\">End</label>\n                  <i *ngIf=\"oplandb.isActMemNoFlight(oplandb.newActivityMember).noFlight\" \n                  class=\"fa fa-plane text-danger p-0 m-0\" \n                  [title]=\"'No flight available on demob date ' + oplandb.isActMemNoFlight(oplandb.newActivityMember).date +'.'\" \n                  style=\"font-size:1.2em;cursor:pointer;margin-left:10px;\"></i>\n                <div class=\"input-group input-group-sm\" >\n                      <input class=\"form-control text-center\"\n                        [dpDayPicker]=\"udf.datePickerConfig\"\n                        [(ngModel)]=\"oplandb.newActivityMember.endDate\"\n                        name=\"assigmentEndDate\"\n                        onfocus=\"this.blur()\"\n                        id=\"memEnd\"\n                        type=\"text\"\n                        style=\"width:170px;\"\n                        >\n                      <div class=\"input-group-addon mx-0\">\n                          <a href=\"#\" (click)=\"openCalendar('memEnd')\"><i class=\"fa fa-calendar text-primary\"></i></a>\n                      </div>\n                    \n                    </div>                    \n                </div>\n              </div>\n\n              <div class=\"row px-3\">\n                <!--*ngIf=\"isCoreTeam(oplandb.newActivityMember.teamID)\"-->\n                <div class=\"form-group\" style=\"width:200px;\">\n\n                    <label class=\"col-form-label my-0 py-0\">Group</label>\n                    <select class=\"form-control form-control-sm\" \n                      id=\"selGrpID\" \n                      [(ngModel)]=\"oplandb.newActivityMember.grpID\" \n                      name=\"grpID\" >\n                      <option *ngFor=\"let grp of oplandb.config.teamGroupsArr\" [ngValue]=\"grp.id\">{{ grp.name }}</option>\n                    </select>    \n\n                </div>\n                <div class=\"col\">\n                  <div class=\"row\">\n                      <label class=\"col-3 text-center my-0\">Mob?</label>\n                      <label class=\"col-3 text-center my-0\">Demob?</label>\n                      <label class=\"col-3 text-center my-0\">Night Shift?</label>\n                      <label class=\"col-3 text-center my-0\"  *ngIf=\"isDayVisible()\">Day Trip?</label>\n                    </div>\n                  <div class=\"row\">\n                      <label class=\"col-3 text-center\">\n                        <i (click)=\"toggleYesNo(oplandb.newActivityMember,'mob')\" \n                        class=\"fa my-1\" style=\"font-size:1.3em;cursor:pointer;\"\n                        [ngClass]=\"clsYesNo(oplandb.newActivityMember,'mob')\" aria-hidden=\"true\"></i>\n                      </label>\n                      <label class=\"col-3 text-center\">\n                          <i (click)=\"toggleYesNo(oplandb.newActivityMember,'demob')\" \n                          class=\"fa my-1\" style=\"font-size:1.3em;cursor:pointer;\"\n                          [ngClass]=\"clsYesNo(oplandb.newActivityMember,'demob')\" aria-hidden=\"true\"></i>\n                      </label>\n                      <label class=\"col-3 text-center\">\n                          <i (click)=\"toggleYesNo(oplandb.newActivityMember,'night')\" \n                          class=\"fa my-1\" style=\"font-size:1.3em;cursor:pointer;\"\n                          [ngClass]=\"clsYesNo(oplandb.newActivityMember,'night')\" aria-hidden=\"true\"></i>\n                      </label>\n                      <label class=\"col-3 text-center\" *ngIf=\"isDayVisible()\">\n                        <i (click)=\"toggleYesNo(oplandb.newActivityMember,'isDay')\" \n                        class=\"fa my-1\" style=\"font-size:1.3em;cursor:pointer;\"\n                        [ngClass]=\"clsYesNo(oplandb.newActivityMember,'isDay')\" aria-hidden=\"true\"></i>\n                      </label>\n                </div>                    \n                    \n                </div>\n\n            </div>\n\n        </div>\n        <div class=\"modal-footer\">\n          <button \n            id=\"btnClose\"\n            type=\"button\" \n            class=\"btn btn-secondary\" \n            data-dismiss=\"modal\" \n            >Close</button>\n          <button type=\"button\" class=\"btn btn-primary\" \n            (click)=\"saveActivityMember()\">Save changes</button>\n        </div>\n      </div>\n    </div>\n  </div>\n  <!-- modal end -->"

/***/ }),

/***/ "../../../../../src/app/components/frm-chart/frm-chart.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return FrmChartComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__app_globals__ = __webpack_require__("../../../../../src/app/app.globals.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__app_udfs__ = __webpack_require__("../../../../../src/app/app.udfs.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3_jquery__ = __webpack_require__("../../../../jquery/dist/jquery.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3_jquery___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_3_jquery__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_4__angular_common_http__ = __webpack_require__("../../../common/@angular/common/http.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_5__app_oplandata__ = __webpack_require__("../../../../../src/app/app.oplandata.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_6__models_data_models__ = __webpack_require__("../../../../../src/app/models/data-models.ts");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};







var FrmChartComponent = (function () {
    function FrmChartComponent(appDataService, udf, http, oplandb) {
        this.appDataService = appDataService;
        this.udf = udf;
        this.http = http;
        this.oplandb = oplandb;
        this.currentTeam = -1;
        this.currentRow = -1;
        this.hoveredTeam = -1;
        this.hoveredRow = -1;
        this.hoveredDate = null;
        this.loadLookups = false;
        this.dragMode = "move";
        this.modPOBBack = "#2b8c2b";
        this.modNonPOBBack = "#732482";
        this.statusText = '';
        this.hovObj = null;
        this.hovObjSrc = null;
        this.scaleFactor = 1;
        this.noSummaryHeader = false;
        this.totalHeaders = 8; // dictates header row height
        this.summaryHeaders = 8;
        this.posDateHeader = 18;
        this.posHGridOffset = this.posDateHeader + this.gridVSpace();
        this.dataRows = 0; // total number of rows shown on the details section of the svg;
        this.vertSeparatorPos = 400;
        this.splitSeparatorSize = 1;
        this.lkpCompanies = {};
        this.lkpPositions = {};
        this.lkpTeams = {};
        this.lkpActivities = {};
        this.filterOffset = 30;
        this.filterGap = 13;
        this.filterCircleOffset = 4;
        // defines the number of rows the header contains
        this.headerRowsIndex = this.fillArrNum(this.totalHeaders - (this.noSummaryHeader ? this.summaryHeaders : 0));
        // defines the horizontal split position
        this.headSeparatorPos = this.posHGridOffset +
            this.gridVSpace() * (this.headerRowsIndex.length + 0.2);
        this.leftHeaderWidths = [
            20, 130, "40%", "60%"
        ];
        this.hideHeaderA = false;
        this.globalTextFill = "gray";
        this.globalFontSize = "12";
        this.datTeams = [
            { id: 1, name: "Operations Core Crew", isCore: true, order: 1 },
            { id: 2, name: "Maintenance Core Crew", isCore: true, order: 2 }
        ];
        this.datCompanies = [
            { id: 1, name: "OPS SPEX" }
        ];
        this.datPositions = [
            { id: 1, name: "OE A" }
        ];
        this.datActivities = [
            { id: 1, name: "Activity 1" },
        ];
        this.datPersonnels = [];
        this.upManningDates = [
            { date: null, limit: 78 },
            { date: "20170405", limit: 70 },
            { date: "20170511", limit: 78 },
        ];
        this.maxBedsDates = [
            { date: null, limit: 59 },
            { date: "20170405", limit: 50 },
            { date: "20170411", limit: 59 },
        ];
        this.dateSpan = 1;
        this.indexArray = [];
        this.bedsSummary = [];
        this.upManTeamSummary = [];
        this.totalOffshorePOB = [];
        this.totalFemale = [];
        this.totalNightShift = [];
        this.totalMOB = [];
        this.totalDEMOB = [];
        this.datSummary = [];
        this.vertLines = [10, 20, 30, 40, 50, 60, 70];
        this.rects = [
            { x: 10, y: 10, c: "red" },
            { x: 60, y: 10, c: "blue" },
            { x: 110, y: 10, c: "red" },
            { x: 160, y: 10, c: "cyan" },
            { x: 210, y: 10, c: "olive" },
            { x: 260, y: 10, c: "teal" },
            { x: 310, y: 10, c: "green" },
            { x: 360, y: 10, c: "lime" }
        ];
        this.alertPrompt = "";
        this.filterTimout = 0;
        this.debugOn = false;
        this.debugText = "test";
        this.dragOffX = 0;
        this.dragOffY = 0;
        this.dragScrollX = 0;
        this.dragScrollY = 0;
        this.noMove = false;
        this.noLeft = false;
        this.noRight = false;
        // process window and table dimensions
        //this.udf.setWindowSizeVars();
        //this.setWindowSizeVars();
        this.getWidthValues(this.leftHeaderWidths, this.svgWidth('LD'));
    }
    FrmChartComponent.prototype.x1 = function () {
        return 10;
    };
    FrmChartComponent.prototype.noHdr = function () {
        return (this.noSummaryHeader ? this.summaryHeaders : 0);
    };
    FrmChartComponent.prototype.canvasWidth = function (rightPadding) {
        if (rightPadding == undefined)
            rightPadding = 0;
        return String(this.dateSpan * this.gridHSpace() + rightPadding);
    };
    FrmChartComponent.prototype.canvasWidthLeft = function (rightPadding) {
        if (rightPadding == undefined)
            rightPadding = 0;
        return String(500 + rightPadding);
    };
    FrmChartComponent.prototype.viewBoxLeft = function (rightPadding) {
        if (rightPadding == undefined)
            rightPadding = 0;
        return "0 0 " + (parseFloat(this.canvasWidthLeft(rightPadding))) / this.scaleFactor + " " + (parseFloat(this.contHeight('LH')) / this.scaleFactor);
    };
    FrmChartComponent.prototype.canvasHeight = function () {
        return "1640";
    };
    FrmChartComponent.prototype.viewBox = function (rightPadding) {
        if (rightPadding == undefined)
            rightPadding = 0;
        return "0 0 " + (parseFloat(this.canvasWidth(rightPadding))) / this.scaleFactor + " " + (parseFloat(this.canvasHeight()) / this.scaleFactor);
    };
    FrmChartComponent.prototype.gridVSpace = function () {
        return 20;
    };
    FrmChartComponent.prototype.gridHSpace = function () {
        //return parseFloat(this.canvasWidth())/(this.dateSpan);
        return 60;
    };
    FrmChartComponent.prototype.gridVCount = function () {
        return this.fillArrNum(20);
    };
    //gridYPos(idxRow:number,bottomOffset?:number):number{
    FrmChartComponent.prototype.gridYPos = function (idxRow, bottomOffset) {
        var spacing = this.gridVSpace();
        if (bottomOffset == undefined)
            bottomOffset = 0;
        return (idxRow * spacing) - bottomOffset; // + parseFloat(this.posHGridOffset.toString());
        //return (idxRow * spacing);// + parseFloat(this.posHGridOffset.toString());
    };
    FrmChartComponent.prototype.dateDiff = function (d1, d2) {
        return Math.ceil((d2.getTime() - d1.getTime()) / (1000 * 3600 * 24));
    };
    FrmChartComponent.prototype.fillArrNum = function (arrLen, iVal) {
        var retVal = Array.from({ length: arrLen }, function (v, i) { return i; });
        if (iVal != undefined)
            retVal.forEach(function (elem) { return elem = iVal; });
        return retVal;
    };
    FrmChartComponent.prototype.initArrVal = function (arrLen, iVal) {
        var retVal = [];
        var i;
        for (i = 0; i < arrLen; i++) {
            retVal.push(iVal);
        }
        return retVal;
    };
    FrmChartComponent.prototype.getDateDisplay = function (idx, getDOW) {
        var dt = new Date(this.udf.calStartDate.getFullYear(), this.udf.calStartDate.getMonth(), this.udf.calStartDate.getDate() + idx);
        if (getDOW == undefined)
            getDOW = false;
        return getDOW ? __WEBPACK_IMPORTED_MODULE_1__app_globals__["getDayOfWeek"](dt.getDay(), true) : dt.getDate() + "-" +
            __WEBPACK_IMPORTED_MODULE_1__app_globals__["getMonthName"](dt.getMonth(), true) + "-" +
            dt.getFullYear().toString().substr(2, 2);
        //+ "<br/>" + G.getDayOfWeek(dt.getDay(),true);
    };
    FrmChartComponent.prototype.getMaxRooms = function (dateIndex) {
        return 59;
    };
    FrmChartComponent.prototype.dateClicked = function (e) {
    };
    FrmChartComponent.prototype.rowClicked = function (teamNumber, rowNumber) {
        this.currentTeam = teamNumber;
        this.currentRow = rowNumber;
    };
    FrmChartComponent.prototype.rowEntered = function (teamNumber, rowNumber) {
        this.hoveredTeam = teamNumber;
        this.hoveredRow = rowNumber;
    };
    FrmChartComponent.prototype.hoverOff = function () {
        this.hoveredTeam = -1;
        this.hoveredRow = -1;
    };
    FrmChartComponent.prototype.ngOnDestroy = function () {
        this.udf.isActivityNormal(true);
        this.oplandb.clearSubscription(this.dataSubscription);
        this.oplandb.clearSubscription(this.subsPersonnel);
        this.oplandb.clearSubscription(this.subsTeam);
        this.oplandb.clearSubscription(this.subsCompany);
        this.oplandb.clearSubscription(this.subsPosition);
    };
    FrmChartComponent.prototype.ngOnInit = function () {
        var _this = this;
        this.oplandb.isEditing = false; // force editing mode
        //this.oplandb.arrangeTeamData()
        //
        setTimeout(function () {
            _this.oplandb.getPersonnel(true, function () {
                _this.oplandb.getTeams(true, function () {
                    _this.oplandb.getCompanies(true, function () {
                        _this.oplandb.getPositions(true, function () {
                            if (_this.oplandb.datActivities.length == 0) {
                                _this.oplandb.getActivities(true, function (data) {
                                    //
                                });
                            }
                        });
                    });
                });
            });
            /*
                this.subsPersonnel = this.oplandb.getPersonnel(true,(data)=>{
                  this.subsTeam = this.oplandb.getTeams(true,(data)=>{
                    this.oplandb.getCompanies(true,(data)=>{
                      this.oplandb.getPositions(true,(data)=>{
                      })
                    })
          
                  })
                })*/
        }, 100);
        //this.oplandb.newActivityMember=new activityMember()
        if (!this.oplandb.newActivityMember) {
            this.oplandb.newActivityMember = new __WEBPACK_IMPORTED_MODULE_6__models_data_models__["b" /* activityMember */]();
        }
        this.oplandb.clearAllSubscription();
        this.currentRow = -1;
        this.currentTeam = -1;
        this.hoveredTeam = -1;
        this.hoveredRow = -1;
        try {
            var calDispStart = localStorage.getItem("calDispStartDate");
            var calDispEnd = localStorage.getItem("calDispEndDate");
            var calFilterDateShort = localStorage.getItem("calDispFilterDate");
            var calFilterPOBvar = localStorage.getItem("calFilterPOB");
            if (calFilterPOBvar != undefined) {
                this.udf.calFilterPOB = calFilterPOBvar == "1";
            }
            var calFilterMOBvar = localStorage.getItem("calFilterMOB");
            if (calFilterMOBvar != undefined) {
                this.udf.calFilterMOB = calFilterMOBvar == "1";
            }
            var calFilterDEMOBvar = localStorage.getItem("calFilterDEMOB");
            if (calFilterDEMOBvar != undefined) {
                this.udf.calFilterDEMOB = calFilterDEMOBvar == "1";
            }
            var calFilterNIGHTvar = localStorage.getItem("calFilterNIGHT");
            if (calFilterNIGHTvar != undefined) {
                this.udf.calFilterNIGHT = calFilterNIGHTvar == "1";
            }
            if (calFilterDateShort != undefined) {
                if (calFilterDateShort.length == 0) {
                    this.udf.calFilterDate = null;
                }
                else {
                    this.udf.calFilterDate = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](calFilterDateShort);
                }
            }
            if (calDispStart != undefined && calDispEnd != undefined) {
                this.udf.calStartDate = __WEBPACK_IMPORTED_MODULE_1__app_globals__["longToDate"](calDispStart);
                this.udf.calEndDate = __WEBPACK_IMPORTED_MODULE_1__app_globals__["longToDate"](calDispEnd);
            }
            else {
                this.udf.calStartDate = new Date();
                this.udf.calEndDate = new Date(this.udf.calStartDate.getFullYear(), this.udf.calStartDate.getMonth(), this.udf.calStartDate.getDate() + this.udf.calcDefaultOffset);
            }
            this.udf.calDispStartDate = __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](this.udf.calStartDate, true);
            this.udf.calDispEndDate = __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](this.udf.calEndDate, true);
            /*
              calStartDate:Date=new Date(2017,3,1);
              calEndDate:Date=new Date(2017,5,30);
            */
            this.reloadCalendar(true);
            //setTimeout(()=>this.reloadCalendar(true),2000);
        }
        catch (e) {
            alert(e.message);
        }
    };
    FrmChartComponent.prototype.getCount = function (type, idx, compWith) {
        var cfg = this.oplandb.config;
        if (!cfg.arranged)
            return "";
        var ctr = 0, i;
        switch (type) {
            case "FEMALE":
                var fArr = this.totalFemale.filter(function (e) { return (e % 2) != 0; });
                return fArr.length;
            case "POBLIMIT":
                return cfg['actualPOBValues'][idx];
            case "ISOVERPOB":
                ctr = 0;
                for (i = 0; i < cfg['actualPOBValues'].length; i++)
                    if (cfg['actualPOBValues'][i] < compWith[i])
                        ctr++;
                return ctr;
            case "NIGHTLIMIT":
                // initial limit value is the actual POB Limit less the standard limit
                // this will only have a value greater than 0 if the upmanning is enforced
                /*
                  2018-02-09 - modified overnight limits to use total
                    offshore pob instead of the upmanning limit as the higher reference value
                */
                //let ret:number = cfg['actualPOBValues'][idx]-cfg['stdPOBValues'][idx];
                var ret = cfg['actualPOBValues'][idx] > cfg['stdPOBValues'][idx] ? this.totalOffshorePOB[idx] - cfg['stdPOBValues'][idx] : 0;
                //let ret:number =  cfg['actualPOBValues'][idx]-cfg['stdPOBValues'][idx];
                /*if(ret==0){
                  let availBeds:number= this.getCount('POBLIMIT',idx) - this.totalOffshorePOB[idx];
                  if(availBeds>=0)ret=999;
                }*/
                return ret;
            case "MOBLIMIT":
                return cfg['mobValues'][idx];
            case "ISOVERMOB":
                ctr = 0;
                for (i = 0; i < cfg['mobValues'].length; i++)
                    if (cfg['mobValues'][i] < compWith[i])
                        ctr++;
                return ctr;
            case "DEMOBLIMIT":
                return cfg['demobValues'][idx];
            case "ISOVERDEMOB":
                ctr = 0;
                for (i = 0; i < cfg['demobValues'].length; i++)
                    if (cfg['demobValues'][i] < compWith[i])
                        ctr++;
                return ctr;
            case "ISOVERNIGHT":
                ctr = 0;
                for (i = 0; i < cfg['actualPOBValues'].length; i++) {
                    /*
                      2018-02-09 - modified overnight limits to use total
                        offshore pob instead of the upmanning limit as the higher reference value
                    */
                    // let limit:number = cfg['actualPOBValues'][i]-cfg['stdPOBValues'][i];
                    var limit = this.totalOffshorePOB[i] - cfg['stdPOBValues'][i];
                    if (limit == 0) {
                        var availBeds = cfg['actualPOBValues'][i] - this.totalOffshorePOB[i];
                        if (availBeds >= 0) {
                            // arbitrary number that will make sure that the actual value to be compared is lower
                            limit = 999;
                        }
                    }
                    else {
                        if (limit > compWith[i])
                            ctr++; // if overnight crew is not filling up the slots for night shift
                    }
                }
                return ctr;
            default:
        }
        return 0;
    };
    FrmChartComponent.prototype.pInfo = function (person) {
        //return this.datPersonnels[index];
        var _this = this;
        if (person['chartInfo'] == undefined) {
            // define and initialize all elements to null chart information
            person['chartInfo'] = [];
            this.indexArray.forEach(function (element) {
                person['chartInfo'].push(1);
            });
            person['activities'].forEach(function (element) {
                // convert activity dates to date type  
                var sd = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](element['start']);
                var ed = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](element['end']);
                // store start date, end date and days to activity object
                element['startDate'] = sd;
                element['endDate'] = ed;
                element['rectSpan'] = _this.dateDiff(sd, ed) + 1;
            });
        }
        return person;
    };
    FrmChartComponent.prototype.dispTestDate = function () {
        return new Date();
    };
    FrmChartComponent.prototype.getPersonDetails = function (person) {
        if (person['chartInfo'] == undefined) {
            var info = {};
            var acts = person['activities'];
            acts.forEach(function (element) {
                var sd;
                //element['dateStart'] = new Date();
                //start:"20170401", end:"20170406"
            });
        }
        return person;
    };
    FrmChartComponent.prototype.setLookupObjects = function () {
        var _this = this;
        this.lkpCompanies = {};
        this.datCompanies.forEach(function (element) {
            _this.lkpCompanies["COY_" + String(element['id'])] = element;
        });
        this.lkpPositions = {};
        this.datPositions.forEach(function (element) {
            _this.lkpPositions["POS_" + String(element['id'])] = element;
        });
        this.lkpTeams = {};
        this.datTeams.forEach(function (element) {
            _this.lkpTeams["TEAM_" + String(element['id'])] = element;
        });
        this.lkpActivities = {};
        this.datActivities.forEach(function (element) {
            _this.lkpActivities["ACT_" + String(element['id'])] = element;
        });
        /*    lkpCompanies:object={};
            lkpPositions:object={};
            lkpTeams:object={};
            lkpActivities:object={};
         */
    };
    FrmChartComponent.prototype.onResize = function (e) {
        //this.setWindowSizeVars()
    };
    FrmChartComponent.prototype.setWindowSizeVars = function () {
        this.windowWidth = __WEBPACK_IMPORTED_MODULE_3_jquery__(window).width();
        this.windowHeight = __WEBPACK_IMPORTED_MODULE_3_jquery__(window).height();
    };
    FrmChartComponent.prototype.onElementScroll = function (e) {
        var cont = __WEBPACK_IMPORTED_MODULE_3_jquery__(e.srcElement);
        var ul = cont.find('[name="ul_pane"]');
        this.dragScrollX = cont.scrollLeft();
        this.dragScrollY = cont.scrollTop();
        ul.css('top', cont.scrollTop());
        ul.css('left', cont.scrollLeft());
        var ur = cont.find('[name="ur_pane"]');
        ur.css('top', cont.scrollTop());
        var ll = cont.find('[name="ll_pane"]');
        ll.css('left', cont.scrollLeft());
    };
    FrmChartComponent.prototype.onResizeDiv = function (e) {
        alert('resize div');
    };
    FrmChartComponent.prototype.contTop = function (area) {
        if (area == 'LH' || area == 'RH') {
            // left and right headers
            return String(this.udf.bannerHeight) + 'px';
        }
        else if (area == 'BN') {
            return '0px';
        }
        else {
            // left and right details (LD & RD)
            return String(this.udf.bannerHeight + this.headSeparatorPos * this.scaleFactor) + 'px';
        }
    };
    FrmChartComponent.prototype.contLeft = function (area) {
        if (area == 'LH' || area == 'LD' || area == 'BN') {
            // left header and left details
            return '0px';
        }
        else {
            // right header and details(RH & RD)
            return String(this.vertSeparatorPos) + 'px';
        }
    };
    FrmChartComponent.prototype.contHeight = function (area) {
        var offsetHeight = this.headSeparatorPos * this.scaleFactor;
        if (area == 'LH' || area == 'RH') {
            return String(offsetHeight - this.splitSeparatorSize) + 'px';
        }
        else if (area == 'BN') {
            return String(this.udf.bannerHeight) + 'px';
        }
        else {
            return String(this.udf.winHt() - offsetHeight - this.udf.bannerHeight) + 'px';
        }
    };
    FrmChartComponent.prototype.contWidth = function (area) {
        var offsetHeight = this.headSeparatorPos * this.scaleFactor;
        if (area == 'LH' || area == 'LD') {
            // left header and left details
            return String(this.vertSeparatorPos - this.splitSeparatorSize) + 'px';
        }
        else if (area == 'BN') {
            return String(this.udf.winWd()) + 'px';
        }
        else {
            // right header and details(RH & RD)
            return String((this.udf.winWd() - this.vertSeparatorPos)) + 'px';
        }
    };
    FrmChartComponent.prototype.contMiddle = function (area) {
        return String(parseInt(this.contWidth(area)) / 2);
    };
    FrmChartComponent.prototype.svgWidth = function (area) {
        if (area == undefined)
            area = "";
        var scrollWidth = 18;
        var retVal = 0;
        switch (area) {
            case "LH":
                retVal = this.vertSeparatorPos - this.splitSeparatorSize + 18;
                break;
            case "LD":
                retVal = this.vertSeparatorPos - this.splitSeparatorSize;
                break;
            case "RH":
                retVal = this.dateSpan * this.gridHSpace() + scrollWidth;
                break;
            case "RD":
                retVal = this.dateSpan * this.gridHSpace();
                break;
        }
        return retVal;
    };
    FrmChartComponent.prototype.svgHeight = function (area) {
        var retVal = 0;
        if (this.datPersonnelsPerTeam == undefined)
            return retVal;
        if (this.datPersonnels.length == undefined)
            return retVal;
        if (area == undefined)
            area = "";
        switch (area) {
            case "LH":
                retVal = parseInt(this.contHeight(area)); //parseInt(this.contHeight(area)) * this.gridVSpace();
                break;
            case "RH":
                retVal = parseInt(this.contHeight(area)) * this.gridVSpace();
                break;
            case "LD":
            case "RD":
                //retVal = (this.datPersonnelsPerTeam.length + this.datPersonnels.length + 2) * this.gridVSpace();
                retVal = (this.dataRows + 1) * this.gridVSpace();
                break;
        }
        return retVal;
    };
    FrmChartComponent.prototype.wrapHeight = function () {
        return (parseInt(this.contHeight('LH')) + parseInt(this.contHeight('LD'))) + 'px';
    };
    FrmChartComponent.prototype.wrapWidth = function () {
        return (parseInt(this.contWidth('LD')) + parseInt(this.contWidth('RD'))) + 'px';
    };
    FrmChartComponent.prototype.svgViewBox = function (area) {
        return "0 0 " + this.svgWidth(area) + " " + this.svgHeight(area);
    };
    FrmChartComponent.prototype.getWidthValues = function (arr, reference) {
        var i;
        var curVal;
        var iCurrentValue = 0;
        var iTotalAbsolute = 0;
        var iRemainingWidth = 0;
        var iValStr;
        var withPercent = false;
        // get total absolute value
        for (i = 0; i < arr.length; i++) {
            iValStr = String(arr[i]);
            if (iValStr.indexOf('%') == -1) {
                // absolute value
                iTotalAbsolute += parseFloat(iValStr);
            }
            else {
                withPercent = true;
            }
        }
        if (withPercent) {
            iRemainingWidth = reference - iTotalAbsolute;
            for (i = 0; i < arr.length; i++) {
                iValStr = String(arr[i]);
                if (iValStr.indexOf('%') != -1) {
                    arr[i] = iRemainingWidth * parseFloat(iValStr) / 100;
                }
            }
        }
        for (i = 1; i < arr.length; i++) {
            arr[i] = parseFloat(arr[i]) + parseFloat(arr[i - 1]);
        }
    };
    FrmChartComponent.prototype.getHeaderXPos = function (width, index) {
        return 50 * index;
    };
    FrmChartComponent.prototype.getDetailYPos = function (idxTeam, idxPerson) {
        var retVal = 0;
        if (idxPerson == undefined) {
            // team y position
        }
        else {
        }
        return retVal;
    };
    FrmChartComponent.prototype.writeSVGText = function (text, row) {
        var retVal = '<text [attr.x]="leftHeaderWidths[2] + 5" font-stretch="narrower" text-anchor="left" [attr.y]="gridYPos(' + row + ',5)"  font-family="Arial" [attr.font-size]="globalFontSize" [attr.fill]="globalTextFill" >' + text + '</text>';
        //let retVal:string="hello";
        return retVal;
    };
    FrmChartComponent.prototype.arrangeData = function () {
        var _this = this;
        this.initVars();
        if (this.bedsSummary == undefined)
            return;
        if (this.maxBedsDates == undefined)
            return;
        var i;
        var summItem;
        this.datSummary = [];
        // set maximum beds
        for (i = 0; i < this.bedsSummary.length; i++) {
            var dt = new Date(this.udf.calStartDate.getFullYear(), this.udf.calStartDate.getMonth(), this.udf.calStartDate.getDate() + i);
            summItem = {
                date: this.getDateDisplay(i),
                DOW: this.getDateDisplay(i, true),
                offPOB: 0,
            };
            var bIdx = void 0;
            var maxIdx = this.maxBedsDates.length - 1;
            for (bIdx = maxIdx; bIdx >= 0; bIdx--) {
                var bed = this.maxBedsDates[bIdx];
                var setMaxBed = (bIdx == 0);
                if (!setMaxBed) {
                    var bedDate = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](bed['date']);
                    setMaxBed = (dt >= bedDate);
                }
                if (setMaxBed) {
                    this.bedsSummary[i] = bed['limit'];
                    summItem['maxBeds'] = bed['limit'];
                    break;
                }
            }
            maxIdx = this.upManningDates.length - 1;
            for (bIdx = maxIdx; bIdx >= 0; bIdx--) {
                var man = this.upManningDates[bIdx];
                var setMaxUp = (bIdx == 0);
                if (!setMaxUp) {
                    var bedDate = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](man['date']);
                    setMaxUp = (dt >= bedDate);
                }
                if (setMaxUp) {
                    //this.upManTeamSummary[i]=man['limit'];
                    this.upManSummary[i] = man['limit'];
                    summItem['maxUpMan'] = man['limit'];
                    break;
                }
            }
            this.datSummary.push(summItem);
        } // grandparent for loop
        // arrange personnel data according to team assignments and 
        // lookup team information from the given lookup table 
        // (by default attach name field and the other optional fields provided on the array field(s) parameter)
        var inputArray = this.datPersonnels;
        var fltIsOn = this.isFiltered();
        if (fltIsOn) {
            // determine if a person will be included in rendering based on the activities under it
            this.datPersonnels.forEach(function (per) {
                per.isIncluded = false;
                per.activities.forEach(function (act) {
                    //if(per.isIncluded)
                    var fltd = __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](_this.udf.calFilterDate);
                    var isDayTrip = (act.mob && act.demob && act.start == act.end && act.isDay);
                    var demoDate = isDayTrip ? act.end : __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](__WEBPACK_IMPORTED_MODULE_1__app_globals__["addDays"](__WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](act.end), 1));
                    if (!per.isIncluded) {
                        // apply filtering criteria here
                        var incTemp = false;
                        if (_this.udf.calFilterPOB) {
                            incTemp = incTemp ||
                                (act.start < fltd && fltd < demoDate ||
                                    (!act.mob && act.start == fltd) ||
                                    (!act.demob && act.end == fltd)) && !act.noPOB;
                        }
                        if (_this.udf.calFilterNIGHT) {
                            incTemp = incTemp ||
                                act.start <= fltd && fltd <= act.end &&
                                    !act.noPOB && act.isNight && !isDayTrip;
                        }
                        if (_this.udf.calFilterMOB) {
                            incTemp = incTemp || (fltd == act.start && act.mob);
                        }
                        if (_this.udf.calFilterDEMOB) {
                            incTemp = incTemp || (fltd == demoDate && act.demob);
                        }
                        per.isIncluded = incTemp;
                        //per.isIncluded = fltd==act.start || 
                        //  fltd==demoDate || (fltd>act.start && fltd<act.end && !act.noPOB) ;
                    }
                });
            });
            //inputArray = this.datPersonnels.filter((per:any)=>{return per.isIncluded});
        }
        else {
            // reset isIncluded to true
            this.datPersonnels.forEach(function (per) {
                // test if
                per.isIncluded = true;
            });
        }
        this.datPersonnelsPerTeam = __WEBPACK_IMPORTED_MODULE_1__app_globals__["groupBy"](inputArray, 'team', this.datTeams, null, ['order', 'isUpManning', 'expChart'], 'order');
        // determine if team group isIncluded based on the isIncluded propery of personnel under
        if (fltIsOn) {
            this.datPersonnelsPerTeam.forEach(function (team) {
                team.isIncluded = false;
                team.dataIncluded = undefined;
                // iterate through each person under the team and cascade isIncluded property
                team.data.forEach(function (per) {
                    if (per.isIncluded) {
                        team.isIncluded = true;
                        return;
                    }
                });
            });
        }
        else {
            this.datPersonnelsPerTeam.forEach(function (team) {
                team.dataIncluded = team.data;
                team.isIncluded = true;
            });
        }
        // set team table element row index
        this.datPersonnelsPerTeam.forEach(function (element) {
            var team = _this.lkpTeams['TEAM_' + element['type']];
            element['isCore'] = team['isCore'] ? "Core" : "NoneCore";
        });
        // group personnel datPersonnelsPerTeam by core and non-core
        this.datPersonnelsCoreNonCore = __WEBPACK_IMPORTED_MODULE_1__app_globals__["groupBy"](this.datPersonnelsPerTeam, 'isCore');
        this.datPersonnelsCoreNonCore.sort(function (a, b) { return (a['type'] > b['type']) ? 1 : ((b['type'] > a['type']) ? -1 : 0); });
        // determine if core/noncore group isIncluded based on the isIncluded propery ot teams under
        if (fltIsOn) {
            this.datPersonnelsCoreNonCore.forEach(function (core) {
                core.isIncluded = false;
                core.dataIncluded = undefined;
                core.data.forEach(function (team) {
                    if (team.isIncluded) {
                        core.isIncluded = true;
                        return;
                    }
                });
            });
        }
        else {
            this.datPersonnelsCoreNonCore.forEach(function (core) {
                core.dataIncluded = core.data;
                core.isIncluded = true;
            });
        }
        // initialize row index thaw will be used to determine band y position
        var currentRowIndex = 0;
        this.datPersonnelsCoreNonCore.forEach(function (core) {
            // iterate through core group
            core['coreSummary'] = [];
            _this.indexArray.forEach(function (idx) {
                core['coreSummary'].push(0); // initialize all elements to null
            });
            // assign and increment row index only if the item will be displayed on the chart
            if (core.isIncluded) {
                core['rowIndex'] = currentRowIndex; // assign row index for the core/noncore
                currentRowIndex++; // increment index for the initial team
            }
            core['data'].forEach(function (team) {
                // iterate through teams
                // sort personnels by group, position, name
                team['data'].sort(function (a, b) {
                    var exprA = (a['grp'] + '|' + a['posName'] + '|' + a['name']);
                    var exprB = (b['grp'] + '|' + b['posName'] + '|' + b['name']);
                    return (exprA > exprB) ? 1 : ((exprB > exprA) ? -1 : 0);
                });
                team['teamSummary'] = [];
                _this.indexArray.forEach(function (idx) {
                    team['teamSummary'].push(0); // initialize all elements to null
                });
                // assign and increment row index only if the item will be displayed on the chart
                if (team.isIncluded) {
                    // assign row index for the team
                    team['rowIndex'] = currentRowIndex;
                    // add all personnels for the current team, then increment by 1
                    if (team['expChart']) {
                        // is expanded in chart
                        // row index must only be incremented by data(personnel) where isIncluded flag is set
                        // plus team group count where at least on personnel under a group is included
                        currentRowIndex += team['data'].filter(function (per) { return per.isIncluded; }).length +
                            _this.groupCount(team) + 1;
                    }
                    else {
                        currentRowIndex++;
                    }
                }
                _this.dataRows = currentRowIndex;
                // roll up summation of personnels to team
                team['data'].forEach(function (person) {
                    person['actSummary'] = [];
                    _this.indexArray.forEach(function (idx) {
                        person['actSummary'].push(null); // initialize all elements to null
                    });
                    // initialize included flag as true when not filtered and false is otherwise
                    //let fltOn:boolean=this.isFiltered();
                    //person['isIncluded'] = (!fltOn)
                    person['activities'].forEach(function (activity) {
                        // convert activity dates to date type  
                        var sd = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](activity['start']);
                        var ed = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](activity['end']);
                        // store start date, end date and days to activity object
                        activity['startDate'] = sd;
                        activity['endDate'] = ed;
                        activity['rectSpan'] = _this.dateDiff(sd, ed) + 1;
                        /*
                        if(activity['mob'] && activity['demob'] && activity['rectSpan'] == 1 && activity['isDay']){
                          // set demob date to end date
                          activity['demobDate'] = ed;
                        }else{
                          // initial demob date is end date + 1
                          activity['demobDate'] = G.addDays(G.stringToDate(activity['end']),1);
                        }
                        
                        
                        if(fltOn && !person['isIncluded']){
                          let fltd:string=G.dateToString(this.udf.calFilterDate);
                          // included if
                          person['isIncluded'] = (fltd==G.dateToString(activity['demobDate'])) || (activity['start'] <= fltd && activity['end'] >= fltd);
            
                        }
                        */
                        _this.fillArrNum(activity['rectSpan']).forEach(function (actIdx) {
                            var actSt = new Date(activity['startDate'].getFullYear(), activity['startDate'].getMonth(), activity['startDate'].getDate() + actIdx);
                            if (actIdx == 0) {
                                var actEndRef = __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](__WEBPACK_IMPORTED_MODULE_1__app_globals__["addDays"](__WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](activity['end']), 1));
                                if (__WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](_this.udf.calStartDate) == actEndRef) {
                                    if (activity["demob"]) {
                                        var sameDay = (activity["start"] == activity["end"]);
                                        if (!sameDay ||
                                            (sameDay && activity["mob"] && activity["isDay"] == 0) ||
                                            (sameDay && !activity["mob"])) {
                                            _this.totalDEMOB[0]++;
                                        }
                                    }
                                }
                            }
                            if (actSt >= _this.udf.calStartDate && actSt <= _this.udf.calEndDate) {
                                // current activity day is within the selection date scope
                                var markedIndex = _this.dateDiff(_this.udf.calStartDate, actSt);
                                // marker -1:mob, 1:normal, 2:demob
                                if (activity['mob'] && actIdx == 0) {
                                    person['actSummary'][markedIndex] = -1;
                                    _this.totalMOB[markedIndex]++; // increment mobilization 
                                    // day trip, increment demobilization
                                    if (activity['rectSpan'] == 1 && activity['demob']) {
                                        _this.totalDEMOB[markedIndex + (activity['isDay'] ? 0 : 1)]++;
                                    }
                                }
                                else if (activity['demob'] && (actIdx == (activity['rectSpan'] - 1))) {
                                    person['actSummary'][markedIndex] = 1;
                                    if (markedIndex < (_this.dateSpan - 1)) {
                                        person['actSummary'][markedIndex + 1] = 2;
                                        _this.totalDEMOB[markedIndex + 1]++; // increment total demobilization
                                    }
                                    //}else if(){
                                    //}
                                }
                                else {
                                    person['actSummary'][markedIndex] = 1;
                                }
                                if (!activity['noPOB']) {
                                    //  POB counting only applies to members of activity with noPOB is set to false
                                    if (activity['rectSpan'] != 1 || (activity['rectSpan'] == 1 && !activity['isDay'])) {
                                        _this.totalOffshorePOB[markedIndex]++; // overall POB count
                                        if (person.gender == "F")
                                            _this.totalFemale[markedIndex]++; // overall POB count
                                        var perObj = _this.oplandb.datPersonnel.find(function (e) { return e.id == person.id; });
                                        if (activity['isNight'])
                                            _this.totalNightShift[markedIndex]++; // night shift count
                                        core['coreSummary'][markedIndex]++; // core/non=core grouping count
                                        team['teamSummary'][markedIndex]++; // team level count
                                        // if person belongs to the up manning team/group
                                        if (team['isUpManning'] != undefined) {
                                            if (team['isUpManning'])
                                                _this.upManTeamSummary[markedIndex]++;
                                        }
                                    }
                                }
                            }
                        }); // for each item in activity.rectSpan
                    }); // for each person's activity
                }); // for each person in a team
            }); // for each team in a core group
        }); // for each core group
        var fltd = "";
        if (this.udf.calFilterDate != null)
            fltd = __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](this.udf.calFilterDate);
        localStorage.setItem("calDispFilterDate", fltd);
    };
    FrmChartComponent.prototype.fillChart = function (person) {
    };
    FrmChartComponent.prototype.isFilteredDate = function (calDate) {
        if (this.udf.calFilterDate == null) {
            return false;
        }
        else {
            return this.udf.calFilterDate == __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](__WEBPACK_IMPORTED_MODULE_1__app_globals__["longToShortDate"](calDate));
        }
    };
    FrmChartComponent.prototype.isSameDate = function (varDate, cmpDate) {
        if (varDate == null) {
            return false;
        }
        else {
            return __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](varDate) == __WEBPACK_IMPORTED_MODULE_1__app_globals__["longToShortDate"](cmpDate);
        }
    };
    FrmChartComponent.prototype.isFiltered = function () {
        if (this.udf.calFilterDate == null) {
            return false;
        }
        else {
            var flt = __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](this.udf.calFilterDate);
            var sd = __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](this.udf.calStartDate);
            var ed = __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](this.udf.calEndDate);
            return (sd <= flt && ed >= flt);
        }
    };
    FrmChartComponent.prototype.dateToLong = function (date) {
        return __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](date, true);
    };
    FrmChartComponent.prototype.getPersonnelActivitiesRaw = function () {
        var _this = this;
        var params = { "start": '20170325', "end": "20170425" };
        var fltd = "";
        if (this.udf.calFilterDate != null) {
            if (this.isFiltered()) {
                //fltd=G.dateToString(this.udf.calFilterDate);
            }
        }
        var url = __WEBPACK_IMPORTED_MODULE_1__app_globals__["getUrlRequest"]("oplanData", "getChartActivities", "start=" + __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](this.udf.calStartDate) +
            "&end=" + __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](this.udf.calEndDate) +
            "&fltd=" + fltd);
        this.oplandb.setLoading("Loading calendar data. Please wait...");
        this.oplandb.clearSubscription(this.dataSubscription);
        this.dataSubscription = this.http
            .get(url)
            .subscribe(function (data) {
            _this.processData(data); // normalize raw data to have the structure expected by the component
            _this.setLookupObjects(); // convert lookup data to hashtable/dictionary format
            _this.arrangeData(); // set data heirarchy
            _this.oplandb.noLoading();
        }, function (err) {
            _this.oplandb.withError(err);
        });
    };
    FrmChartComponent.prototype.processData = function (data) {
        var _this = this;
        var datAct = [];
        var datCoy = [];
        var datTms = [];
        var datPos = [];
        var datPer = [];
        // process activities data
        data['activities'].forEach(function (elem) {
            datAct[datAct.length] = { "id": elem[0], "name": elem[1], "start": elem[2], "end": elem[3] };
        });
        // process companies data
        data['companies'].forEach(function (elem) {
            datCoy[datCoy.length] = { "id": elem[0], "name": elem[1] };
        });
        // process positions data
        data['positions'].forEach(function (elem) {
            datPos[datPos.length] = { "id": elem[0], "name": elem[1] };
        });
        // process teams data
        //sample data format - {id:1,name:"Operations Core Crew",isCore:true,order:1}
        var isFirst = true;
        data['teams'].forEach(function (elem) {
            var exp = localStorage.getItem("texp" + elem[0]);
            if (exp == undefined)
                exp = "0";
            datTms[datTms.length] = {
                "id": elem[0],
                "name": elem[1],
                "isCore": (elem[2] == 1),
                "order": elem[3],
                "isUpManning": (elem[4] == 1),
                "expChart": (exp == "1")
                //"expChart":isFirst
            };
            isFirst = false;
        });
        // process bed and upmanning limits
        // ":[["",78],["20170419",70],["    
        this.maxBedsDates = [];
        this.upManningDates = [];
        data['maxBeds'].forEach(function (elem) {
            _this.maxBedsDates[_this.maxBedsDates.length] = {
                "date": (elem[0].length ? elem[0] : null),
                "limit": elem[1]
            };
        });
        data['upManning'].forEach(function (elem) {
            _this.upManningDates[_this.upManningDates.length] = {
                "date": (elem[0].length ? elem[0] : null),
                "limit": elem[1]
            };
        });
        // process personnels data
        /*sample data format -
        
        {
          id:1,
          name:"Bonifacio Lapuz",
          team:1,
          order:1,
          coy:1,
          pos:6,
          activities:[
            {id:1, start:"20170323", end:"20170408", mob:true, demob:true,coy:1,pos:6,team:1,isNight:true},
            {id:2, start:"20170426", end:"20170511", mob:true, demob:true,coy:1,pos:6,team:1,isNight:true}
          ]
        },
    
        */
        data['personnels'].forEach(function (elem) {
            if (elem != null) {
                var acts_1 = [];
                var posObj = datPos.find(function (item) { return item.id == elem[5]; });
                var posName = elem[5];
                if (posObj)
                    posName = posObj.name;
                datPer[datPer.length] = { "id": elem[0], "name": elem[1],
                    "team": elem[2], "order": elem[3],
                    "coy": elem[4], "pos": elem[5], "grp": elem[6], "gender": elem[9],
                    "activities": acts_1, "posName": (posName + "                                                                                                       ").substr(0, 50) };
                elem[7].forEach(function (act) {
                    acts_1[acts_1.length] = {
                        "id": act[0],
                        "start": act[1],
                        "end": act[2],
                        "mob": (act[3] == 1),
                        "demob": (act[4] == 1),
                        "coy": act[5],
                        "pos": act[6],
                        "team": act[7],
                        "isNight": (act[8] == 1),
                        "noPOB": (act[9] == 1),
                        "grpID": act[10],
                        "isDay": act[11],
                        "acpID": act[12]
                    };
                });
            }
        });
        this.datActivities = datAct;
        this.datCompanies = datCoy;
        this.datPositions = datPos;
        this.datTeams = datTms;
        this.datPersonnels = datPer;
    };
    FrmChartComponent.prototype.isWithActivity = function (person, idxCol) {
        return false;
    };
    FrmChartComponent.prototype.getMidPos = function (p1, p2) {
        return p1 + (p2 - p1) / 2;
    };
    FrmChartComponent.prototype.initVars = function (init) {
        this.dateSpan = this.dateDiff(this.udf.calStartDate, this.udf.calEndDate) + 1;
        this.indexArray = this.fillArrNum(this.dateSpan);
        this.bedsSummary = this.initArrVal(this.dateSpan, 0);
        this.upManSummary = this.initArrVal(this.dateSpan, 0);
        this.upManTeamSummary = this.initArrVal(this.dateSpan, 0);
        this.totalOffshorePOB = this.initArrVal(this.dateSpan, 0);
        this.totalFemale = this.initArrVal(this.dateSpan, 0);
        this.totalNightShift = this.initArrVal(this.dateSpan, 0);
        this.totalMOB = this.initArrVal(this.dateSpan, 0);
        this.totalDEMOB = this.initArrVal(this.dateSpan, 0);
        this.dataRows = 0;
    };
    FrmChartComponent.prototype.reloadCalendar = function (init) {
        if (init == undefined)
            init = false;
        //alert(this.udf.calDispStartDate)
        // initial call to the setLookupObjects and arrangeData is required to format temporary data and
        // to avoid exception error
        if (!init) {
            localStorage.setItem("calDispStartDate", this.udf.calDispStartDate);
            localStorage.setItem("calDispEndDate", this.udf.calDispEndDate);
            this.udf.calStartDate = __WEBPACK_IMPORTED_MODULE_1__app_globals__["longToDate"](this.udf.calDispStartDate);
            this.udf.calEndDate = __WEBPACK_IMPORTED_MODULE_1__app_globals__["longToDate"](this.udf.calDispEndDate);
        }
        else {
            // on initial reload, get configuration
            /*this.oplandb.getConfig(()=>{
              this.oplandb.arrangeConfig();
              setTimeout(()=>{
      
                this.setLookupObjects();
                this.getPersonnelActivitiesRaw();
              },100);
            });*/
        }
        // set date specific limits
        this.oplandb.arrangeConfig();
        // convert lookup data to hashtable/dictionary format
        this.setLookupObjects();
        this.getPersonnelActivitiesRaw();
    };
    FrmChartComponent.prototype.disableReloadX = function () {
        // check if no change on the current filter scope
        var sd = __WEBPACK_IMPORTED_MODULE_1__app_globals__["longToDate"](this.udf.calDispStartDate);
        //let ed:Date=G.longToDate(this.udf.calDispEndDate)
        return true;
    };
    FrmChartComponent.prototype.disableReload = function () {
        // check if no change on the current filter scope
        var sd = __WEBPACK_IMPORTED_MODULE_1__app_globals__["longToDate"](this.udf.calDispStartDate);
        var ed = __WEBPACK_IMPORTED_MODULE_1__app_globals__["longToDate"](this.udf.calDispEndDate);
        var retVal = ((this.udf.calDispStartDate == __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](this.udf.calStartDate, true)) &&
            (this.udf.calDispEndDate == __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](this.udf.calEndDate, true)));
        this.statusText = '';
        if (!retVal) {
            // check if invalid date
            retVal = (ed < sd);
            if (retVal)
                this.statusText = '( Start date cannot be later than end date )';
        }
        else {
            // date was changed
            this.statusText = '';
        }
        if (!retVal) {
            // check maximum dates
            retVal = (__WEBPACK_IMPORTED_MODULE_1__app_globals__["daysDiff"](sd, ed) > this.udf.calScopeLimit);
            if (retVal)
                this.statusText = '( Scope cannot be more than ' + this.udf.calScopeLimit + ' days )';
        }
        //alert(ed-sd);
        //this.statusText=String(G.daysDiff(sd,ed) + ' , ' + this.udf.calScopeLimit);
        return retVal;
    };
    FrmChartComponent.prototype.downloadCalendar = function () {
        //let win=window.open("oplanData.asmx/getFlightPlanData?startDate="+ G.dateToString(this.udf.calStartDate) +"&endDate="+ G.dateToString(this.udf.calEndDate));
        //setTimeout(()=>{win.close();},100);
        this.oplandb.processData({
            method: "getFlightPlanDataNoDownload",
            data: {
                startDate: __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](this.udf.calStartDate),
                endDate: __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](this.udf.calEndDate)
            },
            isPosting: false,
            processMessage: "Extracting flight data. Please wait...",
            onSuccess: function (data) {
                var url = __WEBPACK_IMPORTED_MODULE_1__app_globals__["getUrlRequest"]("oplanData", "downloadFlightPlans");
                var win = window.open(url + "?startDate=" + data['start'] + "&endDate=" + data['end']);
                console.log(url + "?startDate=" + data['start'] + "&endDate=" + data['end']);
                win.blur();
            }
        });
    };
    FrmChartComponent.prototype.clkExpColTeam = function (teamID) {
        var team = this.lkpTeams["TEAM_" + String(teamID)];
        team.expChart = !team.expChart;
        localStorage.setItem("texp" + teamID, team.expChart ? "1" : "0");
        this.setTeamRowPosition();
    };
    FrmChartComponent.prototype.setTeamRowPosition = function () {
        var _this = this;
        var teamObj;
        var currentRowIndex = 0;
        this.datPersonnelsCoreNonCore.forEach(function (core) {
            // iterate through core group
            if (core.isIncluded) {
                core['rowIndex'] = currentRowIndex; // assign row index for the core/noncore
                currentRowIndex++; // increment index for the initial team
                core['data'].forEach(function (team) {
                    // iterate through teams
                    if (team.isIncluded) {
                        teamObj = _this.lkpTeams['TEAM_' + team.type];
                        // assign row index for the team
                        team['rowIndex'] = currentRowIndex;
                        // add all personnels for the current team, then increment by 1
                        if (teamObj.expChart) {
                            // determine number of groups in the current team
                            currentRowIndex += team['data'].filter(function (per) { return per.isIncluded; }).length +
                                _this.groupCount(team) + 1;
                        }
                        else {
                            currentRowIndex++;
                        }
                    }
                    _this.dataRows = currentRowIndex;
                }); // for each team in a core group
            }
        }); // for each core group
    };
    FrmChartComponent.prototype.isNewGroup = function (team, person, idxPerson) {
        var ret = false;
        var grpArr = person.grp.split(':');
        if (idxPerson == 0 && grpArr[1] != "0") {
            ret = true;
        }
        else if (idxPerson != 0) {
            var dataArray = team.data.filter(function (per) { return per.isIncluded; });
            ret = (person.grp != dataArray[idxPerson - 1]['grp']);
        }
        return ret;
    };
    FrmChartComponent.prototype.isTeamExp = function (teamID) {
        var team = this.lkpTeams["TEAM_" + String(teamID)];
        return team.expChart;
    };
    FrmChartComponent.prototype.groupRowIndex = function (team, grp) {
        var grpKey = grp.replace(':', '_') + '_rowIndex';
        var ret = team[grpKey];
        if (ret == undefined)
            ret = 0;
        ret += team.rowIndex;
        return ret;
    };
    FrmChartComponent.prototype.groupText = function (grp) {
        var ret = "";
        var grpArr = grp.split(':');
        if (grpArr[0] == "A") {
            // activity for non-core teams
            // this.lkpActivities["ACT_" + String(element['id'])]=element;
            var actObj = this.lkpActivities["ACT_" + grpArr[1]];
            if (actObj == undefined) {
                ret = "-";
            }
            else {
                ret = actObj.name;
            }
        }
        else {
            // team group for core teams
            var grpObj = this.oplandb.config['idxTeamGroups']['r' + grpArr[1]];
            if (grpObj == undefined) {
                ret = "-";
            }
            else {
                ret = grpObj.name;
            }
        }
        return ret;
    };
    FrmChartComponent.prototype.groupCount = function (team) {
        // returns the count of group per team and set the group rowIndex under the team
        var ret = 0;
        var prvGrp = "-";
        var currRow = 0;
        //team.data.forEach((per:any)=>{
        this.getTeamActivityMembers(team).forEach(function (per) {
            // count groups where personnel isIncluded flag is set
            //      if(per.isIncluded){
            var grpArr = per.grp.split(':');
            if (grpArr[1] != '0') {
                if (prvGrp != per.grp) {
                    ret++;
                    team[per.grp.replace(':', '_') + '_rowIndex'] = ret;
                    prvGrp = per.grp;
                }
            }
            currRow++;
            //      }
        });
        return ret;
    };
    FrmChartComponent.prototype.flagTitle = function (mode, compArr) {
        var ret = '';
        var cnt = this.getCount(mode, 0, compArr);
        if (cnt != 0) {
            switch (mode) {
                case "FEMALE":
                    ret = "There " + (cnt == 1 ? "is" : "are") + " " + cnt + " date" + (cnt == 1 ? "" : "s") + " where the female crew count is not an even number";
                    break;
                case "ISOVERMOB":
                    ret = "There " + (cnt == 1 ? "is" : "are") + " " + cnt + " date" + (cnt == 1 ? "" : "s") + " where the number mobilization exceeded the limit";
                    break;
                case "ISOVERDEMOB":
                    ret = "There " + (cnt == 1 ? "is" : "are") + " " + cnt + " date" + (cnt == 1 ? "" : "s") + " where the number demobilization exceeded the limit";
                    break;
                case "ISOVERNIGHT":
                    ret = "There " + (cnt == 1 ? "is" : "are") + " " + cnt + " date" + (cnt == 1 ? "" : "s") + " where the overnight slots not filled up";
                    break;
                case "ISOVERPOB":
                    ret = "There " + (cnt == 1 ? "is" : "are") + " " + cnt + " date" + (cnt == 1 ? "" : "s") + " where the number of POB's exceeded the limit";
                    break;
            }
        }
        return ret;
    };
    FrmChartComponent.prototype.getActivityBarTitle = function (act, per) {
        var actRec = this.datActivities.find(function (item) { return item.id == act.id; });
        var ret = "Name: " + per.name;
        if (actRec) {
            ret += "\nActivity: " + actRec.name + (this.oplandb.logInfo.activityMemAdmin ? "\n(Click to amend details)" : "");
        }
        return ret;
    };
    FrmChartComponent.prototype.loadActivity = function (act) {
        var _this = this;
        if (this.timeEditActivityMember) {
            clearTimeout(this.timeEditActivityMember);
        }
        var actRec = this.datActivities.find(function (item) { return item.id == act.id; });
        if (actRec) {
            localStorage.setItem("defActID", String(actRec.id));
            this.oplandb.setLoading("Loading activity " + actRec.name + ". Please wait...");
            setTimeout(function () {
                _this.oplandb.onActivityLoadAction = true;
                _this.oplandb.setOtherNavGroup();
                localStorage.setItem("currentModule", "Activities");
                _this.udf.currentModule = "Activities";
            }, 100);
        }
    };
    /* modal window events */
    FrmChartComponent.prototype.onUserChange = function (e) {
    };
    FrmChartComponent.prototype.onUserItemSelected = function (e) {
        var mem = this.oplandb.newActivityMember;
        mem.teamID = e['team']['id'];
        //if(this.oplandb.newActivityMember.id!=-1)return;  // no action if item is being modified
        var per = e['user'];
        var per2 = this.oplandb.idxPersonnel['r' + per.perID];
        mem.coyID = String(per2.companyID);
        mem.posID = String(per2.positionID);
    };
    FrmChartComponent.prototype.dateToPx = function (dt) {
        var offX = parseInt(this.contWidth('LD')) - this.dragScrollX;
        var d1 = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](__WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](this.udf.calStartDate));
        var d2 = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](dt);
        var deltaDate = __WEBPACK_IMPORTED_MODULE_1__app_globals__["daysDiff"](d1, d2);
        //udf.calStartDate
        return offX + deltaDate * this.gridHSpace();
        //return deltaDate;
    };
    FrmChartComponent.prototype.isDayVisible = function () {
        var ret = false;
        if (this.oplandb.newActivityMember.startDate.length > 0 && this.oplandb.newActivityMember.endDate.length > 0) {
            ret = this.oplandb.newActivityMember.mob && this.oplandb.newActivityMember.demob && (this.oplandb.newActivityMember.startDate == this.oplandb.newActivityMember.endDate);
        }
        return ret;
    };
    FrmChartComponent.prototype.clsYesNo = function (obj, fld) {
        return {
            'fa-check-circle': obj[fld],
            'fa-times-circle': !obj[fld],
            'text-success': obj[fld],
            'text-muted': !obj[fld]
        };
    };
    FrmChartComponent.prototype.displayLabel = function () {
        var mem = this.oplandb.newActivityMember;
        var retVal = "Select a member...";
        if (mem.teamID)
            retVal = this.oplandb.idxTeam['r' + mem.teamID]['name'];
        return retVal;
    };
    FrmChartComponent.prototype.isCoreTeam = function (teamID) {
        var team = this.oplandb.idxTeam['r' + teamID];
        if (team == undefined) {
            return true;
        }
        else {
            return team.core;
        }
    };
    FrmChartComponent.prototype.xeditActivityMember = function (act, per) {
    };
    FrmChartComponent.prototype.toggleEditMode = function (mode, conf) {
        var _this = this;
        if (mode == undefined)
            mode = "";
        if (conf == undefined)
            conf = false;
        if (mode == "save") {
            // perform save first before exiting drag and drop editing mode.
            /*
            1. Collect all modified activity assignment
            2. Iterate through all the ssignments and save until nothing is left
               to be processed (ie. filter where sdBackup is not null)
            */
            this.saveActivityMemberDragged(true);
            return;
            //if(withChange) this.arrangeData();
            // this.hovObj=null;
            // this.hovObjSrc=null;
        }
        else if (mode == "cancel") {
            if (conf) {
                this.alertPrompt = "WARNING: This action will clear changes made by drag & drop mode. Do you want to continue?";
                return;
            }
            // restore all original start and end assignment dates...
            var withChange_1 = false;
            this.datPersonnels.forEach(function (p) {
                p['activities'].forEach(function (memDet) {
                    if (memDet) {
                        if (memDet['sdBackup']) {
                            withChange_1 = true;
                            _this.undoMove(memDet, true); // arrangeData will not be called within the undoMove
                        }
                    }
                });
            });
            if (withChange_1)
                this.arrangeData();
        }
        this.oplandb.isEditing = !this.oplandb.isEditing;
    };
    FrmChartComponent.prototype.editActivityMember = function (act, per) {
        var _this = this;
        if (this.oplandb.isEditing) {
            //this.debugText += ", clicked!!"
            return;
        }
        //console.log(this.oplandb.currActivityMember);
        if (!this.oplandb.logInfo.activityMemAdmin)
            return;
        this.oplandb.activityModifiedInCalendar = true;
        // transfer/convert information to activityMember format
        var mem = this.oplandb.newActivityMember;
        this.oplandb.msgTempPromptModal = "";
        var actRec = this.datActivities.find(function (item) { return item.id == act.id; });
        if (!actRec) {
            this.oplandb.msgTempPromptModal = "!Sorry, cannot find reference activity record.";
        }
        else {
            //console.log('actRec...');
            //console.log(actRec);
            if (!this.oplandb.currActivity)
                this.oplandb.currActivity = new __WEBPACK_IMPORTED_MODULE_6__models_data_models__["c" /* dataActivity */]();
            this.oplandb.currActivity.startDate = __WEBPACK_IMPORTED_MODULE_1__app_globals__["shortToLongDate"](actRec.start);
            this.oplandb.currActivity.endDate = __WEBPACK_IMPORTED_MODULE_1__app_globals__["shortToLongDate"](actRec.end);
            //this.oplandb.currActivity=currAct;
            //console.log("current activity...");
            //console.log(this.oplandb.currActivity);
        }
        mem.coyID = String(act.coy);
        mem.demob = act.demob;
        mem.endDate = __WEBPACK_IMPORTED_MODULE_1__app_globals__["shortToLongDate"](act.end);
        mem.grpID = String(act.grpID);
        mem.id = act.acpID;
        mem.idAct = act.id;
        mem.isDay = (act.isDay != 0);
        mem.mob = act.mob;
        mem.night = act.isNight;
        mem.noPOB = act.noPOB;
        mem.perID = per.id;
        mem.posID = String(act.pos);
        mem.startDate = __WEBPACK_IMPORTED_MODULE_1__app_globals__["shortToLongDate"](act.start);
        mem.teamID = act.team;
        this.loadLookups = true;
        //mem.
        this.timeEditActivityMember = setTimeout(function () {
            //this.loadLookups=true;
            _this.oplandb.editActivityMember(_this.oplandb.newActivityMember);
        }, 100);
    };
    FrmChartComponent.prototype.btnNo = function () {
        this.alertPrompt = "";
    };
    FrmChartComponent.prototype.btnYes = function () {
        if (this.alertPrompt.toLowerCase().indexOf("saving") != -1) {
            this.saveActivityMemberDragged();
        }
        else {
            // cancelled
            this.toggleEditMode('cancel');
        }
        this.alertPrompt = "";
    };
    FrmChartComponent.prototype.saveActivityMemberDragged = function (conf) {
        var _this = this;
        if (!conf)
            conf = false;
        if (conf) {
            // if confirmation is invoked, prompt the user to confirm action
            //if(confirm('This process cannot be undone. Do you want to proceed?')){
            //  this.saveActivityMemberDragged();
            //}
            this.alertPrompt = "WARNING: Saving changes made by drag & drop cannot be undone. Do you want to continue?";
            return;
        }
        var withChange = false;
        var mems = [];
        this.datPersonnels.forEach(function (p) {
            p['activities'].forEach(function (memDet) {
                if (memDet) {
                    if (memDet['sdBackup']) {
                        withChange = true;
                        var act = _this.oplandb.datActivities.find(function (a) {
                            return a.id == memDet.id;
                        });
                        mems.push({
                            id: memDet.acpID,
                            actID: memDet.id,
                            perID: p.id,
                            teamID: memDet.team,
                            startDate: memDet.start,
                            endDate: memDet.end,
                            mob: memDet.mob,
                            demob: memDet.demob,
                            night: memDet.isNight,
                            isDay: (memDet.isDay == 1),
                            coyID: String(memDet.coy),
                            posID: String(memDet.pos),
                            groupID: String(memDet.grpID),
                            actExtracted: act.extracted,
                            actStart: __WEBPACK_IMPORTED_MODULE_1__app_globals__["longToShortDate"](act.startDate),
                            actEnd: __WEBPACK_IMPORTED_MODULE_1__app_globals__["longToShortDate"](act.endDate),
                        });
                        //mems.push(memDet);
                        //this.undoMove(memDet);
                        // invoke saving here
                    }
                }
            });
        });
        //this.debugText = JSON.stringify(mems);
        if (mems.length) {
            this.oplandb.saveActivityMember(function (data) {
                _this.hovObj = null;
                _this.hovObjSrc = null;
                _this.oplandb.isEditing = false;
                // clear all backed up information
                _this.datPersonnels.forEach(function (p) {
                    p['activities'].forEach(function (memDet) {
                        if (memDet) {
                            if (memDet['sdBackup']) {
                                memDet['sdBackup'] = null;
                                memDet['edBackup'] = null;
                                memDet['idBackup'] = null;
                            }
                        }
                    });
                });
            }, undefined, undefined, true, mems, true);
        }
        else {
            console.log('Nothing to save!');
        }
    };
    FrmChartComponent.prototype.saveActivityMember = function () {
        var _this = this;
        this.oplandb.saveActivityMember(function (data) {
            var memID = data['activityMemberId'];
            var actID = data['activityId'];
            // search activity member from the activity table if existing...
            if (_this.oplandb.datActivities.length != 0) {
                var act = _this.oplandb.datActivities.find(function (cAct) {
                    return cAct.id == actID;
                });
                if (act) {
                    var mem_1 = _this.oplandb.newActivityMember;
                    if (mem_1.id == memID) {
                        // search member by record id
                        var actMem = act.members.find(function (amem) {
                            return mem_1.id == amem.id;
                        });
                        if (actMem) {
                            var amem = actMem;
                            amem.coyID = mem_1.coyID;
                            amem.demob = mem_1.demob;
                            amem.endDate = mem_1.endDate;
                            amem.grpID = mem_1.grpID;
                            amem.id = mem_1.id;
                            amem.idAct = mem_1.idAct;
                            amem.isDay = mem_1.isDay;
                            amem.mob = mem_1.mob;
                            amem.night = mem_1.night;
                            amem.noPOB = mem_1.noPOB;
                            amem.perID = mem_1.perID;
                            amem.posID = mem_1.posID;
                            amem.startDate = mem_1.startDate;
                            amem.teamID = mem_1.teamID;
                        }
                    } // if currently edited member is the same as the response object's member ....
                } // if activity is found
            }
            _this.reloadCalendar();
        }, undefined, undefined, true);
    };
    FrmChartComponent.prototype.openCalendar = function (id) {
        __WEBPACK_IMPORTED_MODULE_1__app_globals__["setFocus"](id);
    };
    FrmChartComponent.prototype.toggleYesNo = function (obj, fld) {
        obj[fld] = !obj[fld];
    };
    //applyRemoveFilter
    FrmChartComponent.prototype.applyRemoveFilterx = function (item) {
        this.udf.calFilterDateObj = item;
    };
    FrmChartComponent.prototype.applyRemoveFilter = function (item) {
        var _this = this;
        var filterOnDate = (this.udf.calFilterDate == null);
        var dts = __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](__WEBPACK_IMPORTED_MODULE_1__app_globals__["longToDate"](item.date));
        var fdt = "";
        if (!filterOnDate) {
            // filter on 
            fdt = __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](this.udf.calFilterDate);
            filterOnDate = (fdt != dts);
        }
        if (filterOnDate) {
            // apply filter
            this.oplandb.setLoading("Applying filter. Please wait...");
            this.udf.calFilterDate = __WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](dts);
        }
        else {
            // remove filter
            this.oplandb.setLoading("Removing filter. Please wait...");
            this.udf.calFilterDate = null;
        }
        setTimeout(function () {
            _this.arrangeData();
            setTimeout(function () {
                _this.oplandb.noLoading();
            }, 50);
        }, 100);
    };
    FrmChartComponent.prototype.hoverIn = function (item) {
        this.hoveredDate = __WEBPACK_IMPORTED_MODULE_1__app_globals__["longToDate"](item.date);
    };
    FrmChartComponent.prototype.hoverOut = function (item) {
        //this.hoveredDate=null;
    };
    FrmChartComponent.prototype.getTeamActivityMembers = function (team) {
        if (team.dataIncluded == undefined) {
            team.dataIncluded = team.data.filter(function (person) { return person.isIncluded; });
        }
        return team.dataIncluded;
    };
    FrmChartComponent.prototype.getCoreNoneCoreTeams = function (core) {
        if (core.dataIncluded == undefined) {
            core.dataIncluded = core.data.filter(function (team) { return team.isIncluded; });
        }
        return core.dataIncluded;
    };
    FrmChartComponent.prototype.getIncludedCore = function (source) {
        return source.filter(function (core) { return core.isIncluded; });
    };
    FrmChartComponent.prototype.toggleFilter = function (mode) {
        var _this = this;
        var arrange = true;
        switch (mode) {
            case 'pob':
                this.udf.calFilterPOB = !this.udf.calFilterPOB;
                localStorage.setItem("calFilterPOB", this.udf.calFilterPOB ? "1" : "0");
                break;
            case 'mob':
                this.udf.calFilterMOB = !this.udf.calFilterMOB;
                localStorage.setItem("calFilterMOB", this.udf.calFilterMOB ? "1" : "0");
                break;
            case 'demob':
                this.udf.calFilterDEMOB = !this.udf.calFilterDEMOB;
                localStorage.setItem("calFilterDEMOB", this.udf.calFilterDEMOB ? "1" : "0");
                break;
            case 'night':
                this.udf.calFilterNIGHT = !this.udf.calFilterNIGHT;
                localStorage.setItem("calFilterNIGHT", this.udf.calFilterNIGHT ? "1" : "0");
                break;
            default:
                arrange = false;
                break;
        }
        if (arrange) {
            if (this.filterTimout)
                clearTimeout(this.filterTimout);
            this.filterTimout = setTimeout(function () {
                _this.oplandb.setLoading("Applying filter. Please wait...");
                setTimeout(function () {
                    _this.arrangeData();
                    _this.oplandb.noLoading();
                }, 10);
            }, 1200);
        }
    };
    FrmChartComponent.prototype.isDayTrip = function (act) {
        return act.rectSpan == 1 && act.demob && act.isDay;
    };
    FrmChartComponent.prototype.moveDragHandle = function (e, act) {
        if (!act || !e)
            return;
        var dragHandle = __WEBPACK_IMPORTED_MODULE_3_jquery__('#dragHandle');
        act = __WEBPACK_IMPORTED_MODULE_3_jquery__(act);
        // get x,y position of the hovered activity
        var pos = act.position();
        // get relative position of the cursor, relative to the upper left corner of the hovered activity bar
        // <current mouse position - x,y of the activity's upper left corner
        if (e.type == 'mousedown' || e.type == 'mouseenter') {
            // make dragHandle width the same as the activity width
            //!dragHandle.css('width',act.width());
            dragHandle.css('width', act.attr('width'));
            // set offset values
            this.dragOffX = e.x - pos.left;
            this.dragOffY = e.y - pos.top;
        }
        // reposition handle
        dragHandle.css('top', (e.y - this.dragOffY + this.dragScrollY - this.udf.bannerHeight) + 'px');
        dragHandle.css('left', (e.x - this.dragOffX + this.dragScrollX) + 'px');
    };
    FrmChartComponent.prototype.dragHandlePosition = function (mode, offset) {
        if (offset == undefined)
            offset = 0;
        var hObj = __WEBPACK_IMPORTED_MODULE_3_jquery__("#dragHandle");
        if (hObj) {
            return hObj.position()[mode] + offset;
        }
        else {
            return 0;
        }
    };
    //getActMemDetails(actID,mem)
    FrmChartComponent.prototype.getHovMemDetails = function (returnActivity) {
        if (!this.hovObjSrc)
            return null;
        if (returnActivity == undefined)
            returnActivity = false;
        var hovered = __WEBPACK_IMPORTED_MODULE_3_jquery__(this.hovObjSrc);
        var idArr = hovered.attr('id').split('_');
        var act = this.oplandb.datActivities.find(function (e) {
            return e.id == idArr[2];
        });
        if (returnActivity)
            return act;
        var mem = act.members.find(function (e) {
            return e.id == idArr[3];
        });
        if (mem) {
            var ids_1 = "";
            this.datPersonnels.forEach(function (p) {
                ids_1 += p.id + '(' + p.activities.length + '), ';
            });
            var perArr = this.datPersonnels.filter(function (p) {
                return p.id == mem.perID;
            });
            var memDet = null;
            for (var perIdx = 0; perIdx < perArr.length; perIdx++) {
                memDet = perArr[perIdx]['activities'].find(function (acp) {
                    return acp.acpID == idArr[3];
                });
                if (memDet)
                    break;
            }
            return memDet;
        }
        else {
            return null;
        }
    };
    FrmChartComponent.prototype.undoMove = function (memDet, noArrange) {
        if (noArrange == undefined)
            noArrange = false;
        if (memDet == undefined)
            memDet = this.getHovMemDetails();
        if (memDet) {
            if (memDet['sdBackup']) {
                memDet['start'] = memDet['sdBackup'];
                memDet['end'] = memDet['edBackup'];
                memDet['sdBackup'] = null;
                memDet['edBackup'] = null;
                memDet['idBackup'] = null;
            }
            this.hovObj = null;
            this.hovObjSrc = null;
            if (!noArrange)
                this.arrangeData();
        }
    };
    FrmChartComponent.prototype.hovObjLeft = function () {
        if (!this.hovObjSrc)
            return 0;
        var hov = __WEBPACK_IMPORTED_MODULE_3_jquery__("#dragHandle");
        return this.dragScrollX + hov.position().left;
    };
    FrmChartComponent.prototype.hovObjWidth = function () {
        if (!this.hovObjSrc)
            return 0;
        var hov = __WEBPACK_IMPORTED_MODULE_3_jquery__("#dragHandle");
        return hov.width();
    };
    FrmChartComponent.prototype.dragValid = function () {
        var res = "";
        var act = this.getHovMemDetails(true);
        if (!act)
            return false;
        var newDates = this.newStartAndEnd();
        var actStart = __WEBPACK_IMPORTED_MODULE_1__app_globals__["longToShortDate"](act.startDate);
        var actEnd = __WEBPACK_IMPORTED_MODULE_1__app_globals__["longToShortDate"](act.endDate);
        var isOutside = actStart > newDates.start ||
            actStart > newDates.end ||
            actEnd < newDates.start ||
            actEnd < newDates.end;
        return !isOutside;
    };
    FrmChartComponent.prototype.hovObjActLeft = function () {
        if (!this.hovObjSrc)
            return 0;
        var act = this.getHovMemDetails(true);
        if (!act)
            return 0;
        return this.dateToPx(__WEBPACK_IMPORTED_MODULE_1__app_globals__["longToShortDate"](act.startDate)) + this.dragScrollX;
    };
    FrmChartComponent.prototype.hovObjActWidth = function () {
        if (!this.hovObjSrc)
            return 0;
        var act = this.getHovMemDetails(true);
        return (this.dateToPx(__WEBPACK_IMPORTED_MODULE_1__app_globals__["longToShortDate"](act.endDate)) - this.dateToPx(__WEBPACK_IMPORTED_MODULE_1__app_globals__["longToShortDate"](act.startDate))) + this.gridHSpace();
    };
    FrmChartComponent.prototype.curMove = function () {
        //no-drop
        //return "move"
        //this.noMove = !this.dragValid()
        //return this.noMove ? "no-drop" : "move";
        return this.noMove ? "no-drop" : "move";
    };
    FrmChartComponent.prototype.curLeft = function () {
        // this.noLeft = !this.dragValid()
        return this.noLeft ? "no-drop" : "ew-resize";
    };
    FrmChartComponent.prototype.curRight = function () {
        //this.noRight = !this.dragValid()
        return this.noRight ? "no-drop" : "ew-resize";
    };
    FrmChartComponent.prototype.newStartAndEnd = function (setMemNewDates) {
        var dragHandle = __WEBPACK_IMPORTED_MODULE_3_jquery__("#dragHandle");
        var gSpace = this.gridHSpace();
        var hovered = __WEBPACK_IMPORTED_MODULE_3_jquery__(this.hovObjSrc);
        var hovPos = hovered.position();
        var handPos = dragHandle.position();
        //let hovWidth:number = hovered.width();
        var hovWidth = Number(hovered.attr('width'));
        var handWidth = dragHandle.width();
        var leftOff = Math.round((handPos.left - hovPos.left) / gSpace); // negative if set earlier
        var rightOff = Math.round(((handPos.left + handWidth) -
            (hovPos.left + hovWidth)) / gSpace); // negative if set earlier
        var memDet = this.getHovMemDetails();
        if (memDet) {
            if (setMemNewDates == undefined)
                setMemNewDates = false;
            var ds = __WEBPACK_IMPORTED_MODULE_1__app_globals__["addDays"](__WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](memDet['start']), leftOff);
            var de = __WEBPACK_IMPORTED_MODULE_1__app_globals__["addDays"](__WEBPACK_IMPORTED_MODULE_1__app_globals__["stringToDate"](memDet['end']), rightOff);
            var diff = __WEBPACK_IMPORTED_MODULE_1__app_globals__["daysDiff"](ds, de);
            var isDayTrip = (diff == 1 && ds > de);
            var isDay = isDayTrip ? 1 : 0;
            if (isDayTrip) {
                // single day assignment and will demob the same day
                //this.debugText ="Day trip";
                de = ds;
            }
            else if (diff == 0) {
                // single day assignment but will demob the next day
                //this.debugText ="Not day trip " + diff;
            }
            else {
                //this.debugText ="Else " + diff;
            }
            var newStart = __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](ds);
            var newEnd = __WEBPACK_IMPORTED_MODULE_1__app_globals__["dateToString"](de);
            this.debugText = "newStart: " + newStart + ", newEnd: " + newEnd;
            if (setMemNewDates && (newStart != memDet['start'] ||
                newEnd != memDet['end'] ||
                isDay != memDet['isDay'])) {
                if (!memDet['sdBackup']) {
                    memDet['sdBackup'] = memDet['start'];
                    memDet['edBackup'] = memDet['end'];
                    memDet['idBackup'] = memDet['isDay'];
                }
                else if (newStart == memDet['sdBackup'] &&
                    newEnd == memDet['edBackup'] &&
                    memDet['idBackup'] == isDay) {
                    // start and end date were reset to the original
                    memDet['sdBackup'] = null;
                    memDet['edBackup'] = null;
                    memDet['idBackup'] = null;
                }
                memDet['start'] = newStart;
                memDet['end'] = newEnd;
                memDet['isDay'] = isDay;
                this.arrangeData();
            }
            return { start: newStart, end: newEnd };
        }
        else {
            return {};
        }
    };
    FrmChartComponent.prototype.mouseEvent = function (e) {
        var _this = this;
        if (!this.ie())
            return; // if not editing
        var dragHandle = __WEBPACK_IMPORTED_MODULE_3_jquery__("#dragHandle");
        var gSpace = this.gridHSpace();
        var act;
        var pos;
        if (e.type == "mouseenter") {
            // assigned only to the activity bar, initializes the source object identifier hovObjSrc
            // which subsequently be used in the move/resize process
            this.oplandb.showHilights = true;
            this.hovObjSrc = e.target;
            var memDet_1 = this.getHovMemDetails();
            var act_1 = this.datActivities.find(function (a) { return a.id == memDet_1.id; });
            setTimeout(function () { _this.moveDragHandle(e, _this.hovObjSrc); }, 1);
        }
        else if (e.type == "mouseout") {
        }
        else if (e.type == "mousedown") {
            if (!this.hovObjSrc)
                return;
            // make this.hovObj same as the this.hovObjSrc to turn on the dragMask div
            this.hovObj = this.hovObjSrc;
            // get handle div id to determine the type of operation to do when mousemove event is fired
            var hId = __WEBPACK_IMPORTED_MODULE_3_jquery__(e.target).attr('id');
            // set drag mode
            this.dragMode = (hId == "dragHandle" ? "move" : (hId == "dragLeft" ? "w-resize" : "e-resize"));
            this.moveDragHandle(e, this.hovObj);
        }
        else if (e.type == "mouseup") {
            // update display only when new dates are valid
            if (this.dragValid())
                this.newStartAndEnd(true);
            this.hovObj = null;
            this.hovObjSrc = null;
        }
        else if (e.type == 'mousemove') {
            // this should be assigned only to the dragMask
            var hPos = dragHandle.position();
            var nextGridXL = void 0;
            var nextGridXR = void 0;
            var newX = void 0;
            if (this.dragMode == 'move') {
                nextGridXL = hPos.left - gSpace + this.dragScrollX;
                nextGridXR = hPos.left + gSpace + this.dragScrollX;
                newX = (e.x - this.dragOffX + this.dragScrollX);
                this.debugText = "newX: " + newX;
                if (Math.abs(newX - nextGridXL) < (gSpace / 2) || newX < nextGridXL) {
                    dragHandle.css('left', nextGridXL + 'px');
                }
                else if (Math.abs(nextGridXR - newX) < (gSpace / 2) || newX > nextGridXR) {
                    dragHandle.css('left', nextGridXR + 'px');
                }
            }
            else if (this.dragMode == 'w-resize') {
                //!let hRight:number = $(this.hovObj).position().left + $(this.hovObj).width() + this.dragScrollX;
                var hRight = __WEBPACK_IMPORTED_MODULE_3_jquery__(this.hovObj).position().left +
                    Number(__WEBPACK_IMPORTED_MODULE_3_jquery__(this.hovObj).attr('width')) + this.dragScrollX;
                nextGridXL = hPos.left - gSpace + this.dragScrollX;
                nextGridXR = hPos.left + gSpace + this.dragScrollX;
                newX = (e.x - this.dragOffX + this.dragScrollX);
                if (Math.abs(newX - nextGridXL) < (gSpace / 2) || newX < nextGridXL) {
                    dragHandle.css('left', nextGridXL + 'px');
                    dragHandle.css('width', hRight - nextGridXL + 'px');
                }
                else if (Math.abs(nextGridXR - newX) < (gSpace / 2) || newX > nextGridXR) {
                    // change left position only if new left position is less than the
                    // coordinate of the right edge of the activity bar
                    if (hRight > nextGridXR) {
                        dragHandle.css('left', nextGridXR + 'px');
                        dragHandle.css('width', hRight - nextGridXR + 'px');
                    }
                }
            }
            else if (this.dragMode == 'e-resize') {
                var aObj = __WEBPACK_IMPORTED_MODULE_3_jquery__(this.hovObj);
                var aPos = aObj.position();
                //let wd:number=aObj.width();
                var wd = Number(aObj.attr('width'));
                var dwd = dragHandle.width();
                var hRight = aPos.left + wd + this.dragScrollX;
                var hLeft = hPos.left + this.dragScrollX;
                nextGridXL = hLeft + dwd - gSpace;
                nextGridXR = hLeft + dwd + gSpace;
                newX = (e.x - this.dragOffX + this.dragScrollX + wd);
                //this.debugText = "hRight:" + hRight + ", newX:"+newX + ", nextGridXL: " + nextGridXL + ", nextGridXR: " + nextGridXR;
                // change width if ...
                var dXL = (newX - nextGridXL);
                var dXR = (nextGridXR - newX);
                var sp2 = gSpace / 2;
                if (((dXL >= 0 && dXL < sp2) || (newX < (nextGridXL - sp2))) && (nextGridXL - hLeft) >= gSpace) {
                    // shorten the handle 
                    dragHandle.css('width', nextGridXL - hLeft + 'px');
                }
                else if (dXR >= 0 && dXR < sp2 || newX > (nextGridXR + sp2)) {
                    dragHandle.css('width', nextGridXR - hLeft + 'px');
                }
                /*
        
                if(Math.abs(newX - nextGridXL)<(gSpace/2) || newX < nextGridXL){
                  dragHandle.css('width', hRight - nextGridXL +'px');
                }else if(Math.abs(nextGridXR - newX)<(gSpace/2)  || newX > nextGridXR){
                  
                  // change left position only if new left position is less than the
                  // coordinate of the right edge of the activity bar
                  dragHandle.css('width', hRight - nextGridXR +'px');
        
                }
                */
            }
        }
        else {
        }
    };
    FrmChartComponent.prototype.ie = function () {
        // returns oplandb.isEditing
        return this.oplandb.isEditing;
        //return false;
    };
    FrmChartComponent.prototype.withUndo = function () {
        if (!this.hovObjSrc)
            return false;
        var hov = __WEBPACK_IMPORTED_MODULE_3_jquery__(this.hovObjSrc);
        return (hov.attr('fill') == this.modPOBBack ||
            hov.attr('fill') == this.modNonPOBBack);
    };
    FrmChartComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-frm-chart',
            template: __webpack_require__("../../../../../src/app/components/frm-chart/frm-chart.component.html"),
            styles: [__webpack_require__("../../../../../src/app/components/frm-chart/frm-chart.component.css")]
        }),
        __metadata("design:paramtypes", [typeof (_a = (typeof __WEBPACK_IMPORTED_MODULE_1__app_globals__ !== "undefined" && __WEBPACK_IMPORTED_MODULE_1__app_globals__).AppDataService) === "function" && _a || Object, typeof (_b = (typeof __WEBPACK_IMPORTED_MODULE_2__app_udfs__ !== "undefined" && __WEBPACK_IMPORTED_MODULE_2__app_udfs__).AppUDFService) === "function" && _b || Object, typeof (_c = typeof __WEBPACK_IMPORTED_MODULE_4__angular_common_http__["a" /* HttpClient */] !== "undefined" && __WEBPACK_IMPORTED_MODULE_4__angular_common_http__["a" /* HttpClient */]) === "function" && _c || Object, typeof (_d = typeof __WEBPACK_IMPORTED_MODULE_5__app_oplandata__["a" /* AppOplpanDataService */] !== "undefined" && __WEBPACK_IMPORTED_MODULE_5__app_oplandata__["a" /* AppOplpanDataService */]) === "function" && _d || Object])
    ], FrmChartComponent);
    return FrmChartComponent;
    var _a, _b, _c, _d;
}());

//# sourceMappingURL=frm-chart.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/frm-personnel/frm-personnel.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "\r\n.tbcmp thead tr th,.tbcmp tbody tr td{\r\n    font-size:0.8em;   \r\n    padding-top:1px;\r\n    padding-bottom:2px;\r\n}\r\n.tbcmp tbody tr td input[type=\"text\"]{\r\n    font-size:0.9em;\r\n    border:none;\r\n    background: transparent;\r\n}\r\n\r\n/*\r\n.tbcmp thead tr th:nth-child(0), .tbcmp tbody tr td:nth-child(0) {\r\n    width: 250px;\r\n}\r\n.tbcmp thead tr:nth-child(1), .tbcmp tbody tr:nth-child(1) {\r\n    width: 45px;\r\n}\r\n.tbcmp thead tr:nth-child(2), .tbcmp tbody tr:nth-child(2) {\r\n    width: 100px;\r\n}\r\n.tbcmp thead tr:nth-child(3), .tbcmp tbody tr:nth-child(3) {\r\n    width: 100px;\r\n}*/", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/frm-personnel/frm-personnel.component.html":
/***/ (function(module, exports) {

module.exports = "<div class=\"card\" id=\"perCard\" style=\"margin-right:10px;height:100%;\" (window:resize)=\"onResize($event)\">\n<!--div class=\"card\" id=\"perCard\" style=\"margin-right:10px;height:100%;\"-->\n  <div class=\"card-header row py-0\" >\n    <div class=\"col float-left py-0\">\n      <h4 class=\"card-title\">\n        <i style=\"font-size:2em;margin-top:5px;margin-right:20px;\" \n        class=\"fa fa-user text-muted\" aria-hidden=\"true\"></i>Personnel Details</h4>\n    </div>\n    <div class=\" btn-group-sm col m-0 float-right text-right align-text-bottom text-info\"\n      style=\"padding-top:20px;\">\n      <a href=\"#\" \n        class=\"btn btn-link fa-style-a2\" \n        [ngClass]=\"udf.acnBtnClass('add')\"\n        (click)=\"changeRecord($event)\"\n        *ngIf=\"oplandb.logInfo.personnelAdd\"\n      >\n        <i id=\"per_add\" class=\"fa fa-plus-square-o text-success\"></i>\n      </a>\n      <a href=\"#\" class=\"btn btn-link fa-style-a2\" \n        [ngClass]=\"udf.acnBtnClass('edit')\"\n        (click)=\"changeRecord($event)\"\n        *ngIf=\"oplandb.logInfo.personnelEdit && (this.oplandb.currPersonnelObj.id!=-1) && (this.oplandb.datPersonnel.length>0)\"\n        >\n        <i id=\"per_edit\" class=\"fa fa-pencil-square-o text-info\"></i>\n      </a>\n      <a href=\"#\" class=\"btn btn-link fa-style-a2\" \n        [ngClass]=\"udf.acnBtnClass('delete')\"\n        (click)=\"btnClick('DEL')\"\n        *ngIf=\"oplandb.logInfo.personnelDelete && (this.oplandb.currPersonnelObj.id!=-1) && (this.oplandb.datPersonnel.length>0)\"\n        ><i id=\"per_delete\" class=\"fa fa-trash-o text-danger\"></i></a>\n    </div>\n  </div>\n\n\n  <div class=\"card-block py-0\" style=\"overflow-y:auto;\" >\n    <form>\n\n    <div *ngIf=\"oplandb.msgTempPrompt.length\" class=\"alert alert-danger row\"\n      role=\"alert\"\n      style=\"position:relative;top:0px;left:-5px;width:calc(100% + 40px);\">\n      <strong>Reminder: </strong>&nbsp;{{oplandb.msgTempPrompt}}\n    </div>      \n\n    <div *ngIf=\"oplandb.msgTempInfoPrompt.length\" class=\"alert alert-info row\"\n      role=\"alert\"\n      style=\"position:relative;top:0px;left:-5px;width:calc(100% + 40px);\">\n      <strong>Information: </strong>&nbsp;{{oplandb.msgTempInfoPrompt}}\n    </div>      \n\n    \n\n    <div *ngIf=\"dispAlert('delete')\"\n      class=\"alert alert-danger row my-0\" \n      role=\"alert\" style=\"position:relative;top:0px;left:-5px;width:calc(100% + 40px);\">\n      <strong>{{oplandb.alertYesNoMessage.label}}:</strong> &nbsp;{{oplandb.alertYesNoMessage.message}}&nbsp;\n      <div class=\"btn-group-sm\">\n        <button class=\"btn btn-danger\" (click)=\"btnClick('goDelete')\" style=\"cursor:pointer\">{{oplandb.alertYesNoMessage.pos}}</button>\n        <button class=\"btn btn-success\" (click)=\"btnClick('noDelete')\" style=\"cursor:pointer\">{{oplandb.alertYesNoMessage.neg}}</button>\n      </div>\n    </div>        \n\n    <div *ngIf=\"oplandb.msgPersonUpdateSuccess.length>0\"\n      class=\"alert alert-success\" \n      role=\"alert\" style=\"position:relative;top:0px;left:-20px;width:calc(100% + 40px);\">\n      {{ oplandb.msgPersonUpdateSuccess }}\n    </div>      \n\n    <div *ngIf=\"oplandb.msgPersonUpdateInstruction.length>0\"\n        class=\"alert alert-info\" \n        role=\"alert\" style=\"position:relative;top:0px;left:-20px;width:calc(100% + 40px);\">\n        {{ oplandb.msgPersonUpdateInstruction }}\n    </div>      \n\n    <div *ngIf=\"oplandb.msgPersonUpdateError.length>0\"\n        class=\"alert alert-danger\" \n        role=\"alert\" style=\"position:relative;top:0px;left:-20px;width:calc(100% + 40px);\">\n        {{ oplandb.msgPersonUpdateError }}\n    </div>\n\n    <div class=\"row\">\n      <div class=\"col d-flex p-0\"> <!--detail wrapper start-->\n        <div class=\"row\" style=\"width:100%;\">\n            <div class=\"col d-flex flex-column\">\n\n          \n\n                <!-- Full name -->\n                <div class=\"form-group m-0 col p-0\">\n                    <label class=\"m-0 mt-2\" for=\"exampleInputEmail1\">Full Name</label>\n                    <input type=\"text\" class=\"form-control form-control-sm\" id=\"firstName\"  \n                      placeholder=\"Enter full name\" \n                      [disabled]=\"udf.modePersonnel=='normal'\"\n                      [(ngModel)]=\"oplandb.currPersonnel.fullName\"\n                      id=\"personnelName\"\n                      name=\"fullName\">\n                </div>\n\n\n\n                <div class=\"d-flex flex-row m-0 px-0\" >  <!--gender and birth date start-->\n\n\n      \n                  <!-- Birth Date -->\n                  <div class=\"form-group m-0 col p-0\" >\n                      <label class=\"m-0 mt-2\" for=\"birthDate\" >Birth Date</label>\n                      <div class=\"input-group input-group-sm mb-2 mr-sm-2 m-0\" >\n                        <input type=\"text\" class=\"form-control form-control-sm\" id=\"birthDate\"  \n                          placeholder=\"Select birth date\" \n                          [disabled]=\"udf.modePersonnel=='normal'\"\n                          [(ngModel)]=\"oplandb.currPersonnel.birthDate\"\n                          [dpDayPicker]=\"udf.datePickerConfig\"\n                          onfocus=\"this.blur();\"\n                          (dblclick)=\"oplandb.currPersonnel.birthDate=''\"\n                          name=\"birthDate\">\n      \n                        <div class=\"input-group-addon\"><a href=\"#\" \n                          (click)=\"openCalendar('birthDate')\">\n                            <i class=\"fa fa-calendar text-primary\"></i></a></div>\n      \n                      </div>\n      \n      \n                  </div><!-- DATE OF BIRTH END-->\n\n                  <div *ngIf=\"true\" class=\"form-group m-0  p-0 pt-2 pl-3\" >\n                    <label class=\"m-0 pl-2 pr-1\">XBR</label>\n                    <div class=\"p-0 text-center pt-1\">\n                      <i [ngClass]=\"{'fa-check-circle':oplandb.currPersonnel.xl,'text-danger':oplandb.currPersonnel.xl,'fa-times-circle':!oplandb.currPersonnel.xl,'text-muted':!oplandb.currPersonnel.xl}\" \n                        class=\"fa\" title=\"Click to toggle value...\"\n                        (click)=\"toggleXL()\"\n                        style=\"font-size:1.3em;cursor: pointer;\"></i>\n                    </div>\n                  </div>\n\n                  <div class=\"form-group m-0  p-0 pt-2 pl-3\" >\n                      <label class=\"m-0 pl-2\">Gender</label>\n\n                      <div class=\"p-0\" >\n                          <button class=\"btn btn-link p-0 px-1\" (focus)=\"onButtonFocus($event)\" (click)=\"setGender('M')\"\n                          [style.border]=\"oplandb.currPersonnel.gender=='M'?'1px solid #c0c0c0':'1px solid white'\"\n                            style=\"border-radius:6px;cursor:pointer;\" title=\"Click to set to Male\" >\n                            <i class=\"fa fa-male\" \n                            style=\"font-size:1.65em;color:#c0c0c0;\"\n                            [style.text-shadow]=\"oplandb.currPersonnel.gender=='M' ? '1px 1px yellow' : '1px 1px white'\"\n                            [ngClass]=\"{'text-success':oplandb.currPersonnel.gender=='M'}\"></i>\n                        </button>\n      \n                        <button class=\"btn btn-link m-0 p-0 px-1\"  (focus)=\"onButtonFocus($event)\"  (click)=\"setGender('F')\"\n                          [style.border]=\"oplandb.currPersonnel.gender=='F'?'1px solid #c0c0c0':'1px solid white'\"\n                            style=\"border-radius:6px;cursor:pointer;\"  title=\"Click to set to Female\">\n                            <i class=\"fa fa-female\" \n                            style=\"font-size:1.65em;color:#c0c0c0;\"\n                            [style.text-shadow]=\"oplandb.currPersonnel.gender=='F' ? '1px 1px yellow' : '1px 1px white'\"\n                            [ngClass]=\"{'text-success':oplandb.currPersonnel.gender=='F'}\"></i>\n                        </button>\n      \n                          <input type=\"hidden\" class=\"form-control\" id=\"inpGender\" \n                            placeholder=\"Select Gender\"\n                            [disabled]=\"udf.modePersonnel=='normal'\"\n                            [(ngModel)]=\"oplandb.currPersonnel.gender\"\n                            name=\"gender\">\n  \n                          </div>                      \n\n                  </div>\n\n\n      \n\n            </div> <!--gender and birth date end-->                \n\n\n\n\n\n      \n                <!-- Functional Sponsor_Company -->\n                <div class=\"form-group col m-0 p-0\">\n                  <label class=\"m-0 mt-2\" for=\"exampleSelect1\">Functional Sponsor_Company</label>\n                  <div class=\"input-group\" >\n                    <input type=\"text\" class=\"form-control form-control-sm\" id=\"firstName\"  \n                      [disabled]=\"udf.modePersonnel=='normal'\"\n                      [value]=\"oplandb.getCompanyName(oplandb.currPersonnel.companyID)\" \n                      onfocus=\"this.blur()\">           \n      \n                    <!-- custom dropdown use \"btn-group dropdown\" class [start] -->\n                    <div class=\"btn-group dropdown\">\n                        <button type=\"button\" \n                          class=\"btn p-0 px-2 text-primary\" data-toggle=\"dropdown\" \n                          title=\"Click to select company...\"\n                          style=\"cursor: pointer;border-radius:2px;border:1px solid #c0c0c0;\"\n                          aria-haspopup=\"true\" aria-expanded=\"false\"\n                          [disabled]=\"udf.modePersonnel=='normal'\">\n                          <i class=\"fa fa-list-alt\"></i>\n                        </button>\n      \n                        <app-ctl-drop-down  class=\"dropdown-menu p-0\"\n                          [drpdData]=\"oplandb.datCompany\"\n                          [(drpdValue)]=\"oplandb.currPersonnel.companyID\" \n                          [drpdDispRows]=\"15\"\n                          \n                          style=\"z-index:10;left:-150px;\"\n                        ></app-ctl-drop-down>          \n                        \n                    </div> \n                    <!-- custom dropdown use \"btn-group dropdown\" class [end] -->            \n                  </div>  <!--input group-->\n                </div>  <!-- functional sponsor_company end-->\n      \n                <!-- position start -->\n                <div class=\"form-group col m-0 p-0\">\n      \n                    <label class=\"m-0 mt-2\" for=\"exampleSelect1\">Position</label>\n      \n                    <div class=\"input-group\" >\n        \n                      <input type=\"text\" class=\"form-control form-control-sm\" id=\"firstName\"  \n                        [disabled]=\"udf.modePersonnel=='normal'\"\n                        [value]=\"oplandb.getPositionName(oplandb.currPersonnel.positionID)\" \n                        onfocus=\"this.blur()\"\n                      >           \n                      \n                      <!-- custom dropdown use \"btn-group dropdown\" class [start] -->\n                      <div class=\"btn-group dropdown\">\n                        <button type=\"button\" \n                          title=\"Click to select position...\"\n                          class=\"btn p-0 px-2 text-primary\" data-toggle=\"dropdown\" \n                          style=\"cursor: pointer;border-radius:2px;border:1px solid #c0c0c0;\"\n                          aria-haspopup=\"true\" aria-expanded=\"false\"\n                          [disabled]=\"udf.modePersonnel=='normal'\"\n                          >\n                          <i class=\"fa fa-list-alt\"></i>\n                        </button>\n        \n                        <app-ctl-drop-down  class=\"dropdown-menu p-0\"\n                          [drpdData]=\"oplandb.datPosition\"\n                          [(drpdValue)]=\"oplandb.currPersonnel.positionID\" \n                          [drpdDispRows]=\"15\"\n                          \n                          style=\"z-index:10;left:-150px;\"\n                        ></app-ctl-drop-down>          \n                        \n                      </div> \n                      <!-- custom dropdown use \"btn-group dropdown\" class [end] -->\n        \n                    </div>\n      \n                </div> <!-- position end -->\n      \n          \n\n            </div>\n\n        </div>\n      </div> <!-- Detail wrapper end -->\n\n      <div class=\"p-0 mr-3\" style=\"width:485px;position:relative;\">\n        <span *ngIf=\"oplandb.currPersonnel.id==-1\"\n            title=\"This section is only editable after saving the new personnel record...\" \n            style=\"position:absolute;z-index:10;font-size:11em;opacity: 0.2;margin-left:150px;\">\n            <i class=\"fa fa-ban text-info\"></i></span>\n          <!--label class=\"m-0 mt-1\" for=\"exampleSelect1\" style=\"font-size: 0.9m;\">Compentency</label-->\n\n          <table id=\"comp\"\n          class=\"table table-striped table-sm table-hover table-bordered tbcmp tbc m-0 mt-4\">\n          <thead>\n              <tr>\n                <th style=\"width:150px;\">Certificate</th>\n                <th style=\"width:40px;\" class=\"px-0\">Req</th>\n                <th style=\"width:40px;\" class=\"px-0\">Freq</th>\n                <th style=\"width:125px;\">Issued</th>\n                <th style=\"width:125px;\">Expiry</th>\n              </tr>\n          </thead>\n          <tbody style=\"overflow-y: hidden;height:208px;\">\n            <tr *ngFor=\"let row of certSource();trackBy:oplandb.trackByItem;let ridx=index;\"\n              (click)=\"curComp(row)\"\n            >\n              <td style=\"width:150px;\">{{row.name}}</td>\n              <td class=\"text-center\" style=\"width:40px;\">\n                  <i  \n                  [style.cursor]=\"row.mandatory ? 'default' : 'pointer'\"\n                  (click)=toggleRequired(row);\n                  [title]=\"row.mandatory ? 'This is a mandatory certificate...' :'Toggle competency required flag...'\"\n                  [class.fa-times-circle]=\"!row.required\"\n                  [class.text-muted]=\"!row.required\"\n                  [class.fa-check-circle]=\"row.required\"\n                  [class.text-success]=\"row.required\"\n                  class=\"fa\"></i>              \n              </td>\n              <td  class=\"text-center p-0\" style=\"width:40px;\">\n                {{row.frequency ? row.frequency + (row.frequency>1?' yrs':' yr') : '&nbsp;'}}\n                <!--span *ngIf=\"!calEdit(row)\">{{row.frequency ? row.frequency : '&nbsp;'}}</span>\n                <select class=\"form-control form-control-sm p-0 m-0\" \n                style=\"height:22px;width:38px;font-size:9pt;\" \n                  [(ngModel)]=\"row.frequency\" \n                  [name]=\"'frequency_'+ridx\"\n                  *ngIf=\"calEdit(row)\">\n                  <option value=\"0\"></option>\n                  <option value=\"1\">1</option>\n                  <option value=\"2\">2</option>\n                  <option value=\"3\">3</option>\n                  <option value=\"4\">4</option>\n                  <option value=\"5\">5</option>\n                  <option value=\"6\">6</option>\n                </select-->\n              </td>\n              <td [ngClass]=\"{'text-right':calEdit(row),'text-center':!calEdit(row)}\" \n               style=\"width:125px;\">\n                  <span *ngIf=\"!calEdit(row)\">{{row.issued}}</span>\n                  <div class=\"input-group\" *ngIf=\"calEdit(row)\">\n                      <span class=\"text-muted\" style=\"cursor:pointer; font-weight: normal;font-size:0.8em;\" \n                        title=\"Clear issued date\" (click)=\"clearEntry(row,'issued')\">\n                          <i class=\"fa fa-times\"></i>\n                      </span>\n                      <input class=\"form-control form-control-sm text-center m-0 p-0 border-0\" \n                        [disabled]=\"udf.modePersonnel=='normal' || oplandb.currPersonnel.id==-1\"\n                        type=\"text\" \n                        id=\"issued-date\"\n                        [(ngModel)]=\"row.issued\" \n                      [name]=\"'issued_'+ridx\"\n                      [dpDayPicker]=\"udf.datePickerConfig\"\n                      onfocus=\"this.blur();\" style=\"width:90px;\">\n  \n                      <span style=\"cursor:pointer;\" \n                      (click)=\"openCalendar('issued-date')\"\n                      title=\"Select expiry date...\"><i class=\"fa fa-calendar\"></i></span>                    \n  \n                    </div>\n              </td>\n              <td [ngClass]=\"{'text-right':calEdit(row),'text-center':!calEdit(row)}\" \n               style=\"width:125px;\">\n                  <!--span *ngIf=\"!calEdit(row)\">{{row.expiry}}</span-->\n                  <span *ngIf=\"!calEdit(row)\">{{displayExpiry(row)}}</span>\n\n\n                  <div class=\"input-group\" *ngIf=\"calEdit(row)\">\n                    <span class=\"text-muted\" style=\"cursor:pointer; font-weight: normal;font-size:0.8em;\" \n                      title=\"Clear expiry date\" (click)=\"clearEntry(row,'expiry')\">\n                        <i class=\"fa fa-times\"></i>\n                    </span>\n                    <input class=\"form-control form-control-sm text-center m-0 p-0 border-0\" \n                    [disabled]=\"udf.modePersonnel=='normal' || oplandb.currPersonnel.id==-1\"\n                      type=\"text\" \n                      id=\"expiry-date\"\n\n                     [(ngModel)]=\"row.expiry\" \n                    [name]=\"'expiry_'+ridx\"\n                    [dpDayPicker]=\"udf.datePickerConfig\"\n                    onfocus=\"this.blur();\" style=\"width:90px;\">\n                    <span style=\"cursor:pointer;\" \n                    (click)=\"openCalendar('expiry-date')\"\n                    title=\"Select expiry date...\"\n                    \n                    ><i class=\"fa fa-calendar\"></i></span>                    \n                  </div>\n\n              </td>\n            </tr>\n          </tbody>\n          </table>\n      </div>\n    </div>\n       \n      <!-- Current and upcoming activities -->\n      <div class=\"row\">\n          <div class=\"form-group col my-0\" >\n            <div class=\"row\">\n                <label class=\"col\">Current and upcoming activities:</label>\n                <label class=\"col text-right\" *ngIf=\"oplandb.currPersonnel.maxDays\">Maximum number of days offshore&nbsp;<span style=\"font-size:0.75em;color:white;\" class=\"class=badge badge-pill\" [ngClass]=\"{'badge-info':oplandb.currPersonnel.maxDays<=21,'badge-danger':oplandb.currPersonnel.maxDays>21}\">{{ oplandb.currPersonnel.maxDays }}</span></label>\n            </div>\n            \n            <table id=\"acts\" \n              class=\"table table-striped table-sm table-hover table-bordered tbc my-0 m-0\"\n              >\n              <thead>\n                  <tr>\n                    <th [style.width]=\"actsColsPX[0]+'px'\" >Activity</th>\n                    <th [style.width]=\"actsColsPX[1]+'px'\">Start</th>\n                    <th [style.width]=\"actsColsPX[2]+'px'\">End</th>\n                    <th [style.width]=\"actsColsPX[3]+'px'\">MOB</th>\n                    <th [style.width]=\"actsColsPX[4]+'px'\">DEMOB</th>\n                    <th [style.width]=\"actsColsPX[5]+'px'\">Night Shift</th>\n                  </tr>\n              </thead>\n              <tbody [style.height]=\"actsBodyHeight+'px'\">\n                <tr *ngFor=\"let act of oplandb.currPersonnel.activities;trackBy:oplandb.trackByItem;let actIdx=index;\">\n                  <td [style.width]=\"actsColsPX[0]+'px'\">{{ oplandb.lkpActivityName['a_'+act.aid] }}</td>\n                  <td [style.width]=\"actsColsPX[1]+'px'\" style=\"text-align:center;\">{{act.start}}</td>\n                  <td [style.width]=\"actsColsPX[2]+'px'\" style=\"text-align:center;\">{{act.end}}</td>\n  \n                  <td [style.width]=\"actsColsPX[3]+'px'\" style=\"text-align:center;\">\n                    <i \n                    [class.fa-times-circle]=\"act.mob!='1'\"\n                    [class.text-muted]=\"act.mob!='1'\"\n                    [class.fa-check-circle]=\"act.mob=='1'\"\n                    [class.text-success]=\"act.mob=='1'\"\n                    class=\"fa\"></i>\n                  </td>\n                  <td [style.width]=\"actsColsPX[4]+'px'\" style=\"text-align:center;\">\n                    <i \n                    [class.fa-times-circle]=\"act.demob!='1'\"\n                    [class.text-muted]=\"act.demob!='1'\"\n                    [class.fa-check-circle]=\"act.demob=='1'\"\n                    [class.text-success]=\"act.demob=='1'\"\n                    class=\"fa\"></i>\n                  </td>\n                  <td [style.width]=\"actsColsPX[5]+'px'\" style=\"text-align:center;\">\n                    <i \n                    [class.fa-times-circle]=\"act.night!='1'\"\n                    [class.text-muted]=\"act.night!='1'\"\n                    [class.fa-check-circle]=\"act.night=='1'\"\n                    [class.text-success]=\"act.night=='1'\"\n                    class=\"fa\"></i>\n                  </td>\n          \n                </tr>\n              </tbody>              \n            </table>\n\n          </div>\n        </div>\n          \n    </form>\n\n  </div>\n\n\n  <div class=\"card-footer\" *ngIf=\"udf.modePersonnel!='normal' && !udf.isPersonnelDeleteAlert()\">\n      <div class=\"form-buttons align-bottom\" >\n          <button type=\"button\" class=\"btn btn-secondary\"\n            [class.disabled]=\"udf.modePersonnel=='normal'\"\n            (click)=\"cancelUpdate()\">Cancel</button>\n            <button type=\"button\" class=\"btn btn-primary\"\n            (click)=\"savePersonnel()\" >Save</button>\n        </div>\n    </div>\n</div>"

/***/ }),

/***/ "../../../../../src/app/components/frm-personnel/frm-personnel.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return FrmPersonnelComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__models_data_models__ = __webpack_require__("../../../../../src/app/models/data-models.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2_jquery__ = __webpack_require__("../../../../jquery/dist/jquery.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2_jquery___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_2_jquery__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3__app_globals__ = __webpack_require__("../../../../../src/app/app.globals.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_4__app_udfs__ = __webpack_require__("../../../../../src/app/app.udfs.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_5__app_oplandata__ = __webpack_require__("../../../../../src/app/app.oplandata.ts");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};






var FrmPersonnelComponent = (function () {
    function FrmPersonnelComponent(appDataService, oplandb, udf) {
        this.appDataService = appDataService;
        this.oplandb = oplandb;
        this.udf = udf;
        this.actsCols = ["100%", 100, 100, 100, 100, 100];
        this.actsBodyHeight = 150;
        this.lkpCompanies = [
            { id: 1, name: "Soga" },
            { id: 2, name: "DOF" },
            { id: 3, name: "SPEX" }
        ];
        this.lkpPositions = [
            { id: 1, name: "Electrical" },
            { id: 2, name: "Assessor" },
            { id: 3, name: "Integrity Engineer" },
            { id: 4, name: "Programmer" }
        ];
        this.actionMode = "add";
        this.isDirty = false;
        this.loadLookups = false;
    }
    FrmPersonnelComponent.prototype.ngOnInit = function () {
        var _this = this;
        this.oplandb.itemsGroupings = [];
        this.loadLookups = false;
        this.oplandb.clearAllSubscription();
        this.oplandb.getPersonnel(false, function (data) {
            _this.oplandb.getPersonnelActiveAcivities();
            _this.oplandb.setPersonnelNavGroup();
        });
        this.setTableColWidths();
        this.oplandb.getCredentials();
        this.udf.isPersonnelNormal(true);
    };
    FrmPersonnelComponent.prototype.onResize = function (e) {
        this.setTableColWidths();
    };
    FrmPersonnelComponent.prototype.setTableColWidths = function () {
        this.actsColsPX = this.udf.getAbsWidthVals(this.actsCols, __WEBPACK_IMPORTED_MODULE_2_jquery__('#acts').width());
        this.actsBodyHeight = this.udf.getReactiveTableHeight('acts', 'perCard', 40, 35);
    };
    FrmPersonnelComponent.prototype.savePersonnel = function () {
        if (this.udf.modePersonnel == "normal")
            return;
        this.oplandb.savePersonnel();
    };
    FrmPersonnelComponent.prototype.cancelUpdate = function () {
        this.udf.modePersonnel = "normal";
        if (this.oplandb.currPersonnelObj.id != -1) {
            this.oplandb.resetMessages();
            this.oplandb.scatterPersonnel();
        }
    };
    FrmPersonnelComponent.prototype.clearEntry = function (row, fld) {
        row[fld] = "";
    };
    FrmPersonnelComponent.prototype.dispAlert = function (mode) {
        return this.udf.isPersonnelDeleteAlert();
    };
    FrmPersonnelComponent.prototype.calEdit = function (row) {
        if (this.udf.modePersonnel == 'normal')
            return false;
        if (!row.current)
            return false;
        if (this.oplandb.currPersonnel.id < 0)
            return false;
        //if(!row.required) return false;
        return true;
    };
    FrmPersonnelComponent.prototype.curComp = function (row) {
        var cRow = this.oplandb.datCertificates.find(function (e) {
            return e.current;
        });
        if (cRow) {
            if (cRow.id != row.id) {
                cRow.current = false;
            }
        }
        row.current = true;
    };
    FrmPersonnelComponent.prototype.btnClick = function (mode) {
        switch (mode) {
            case 'DEL':
                if (this.oplandb.currPersonnelObj.id == -1) {
                    this.oplandb.setTempMessage(undefined, "Please a personnel to delete.");
                }
                else {
                    this.udf.isPersonnelDeleteAlert(true);
                }
                break;
            case 'goDelete':
                this.oplandb.deletePersonnel();
                break;
            case 'noDelete':
                this.udf.isPersonnelNormal(true);
                break;
        }
    };
    FrmPersonnelComponent.prototype.onButtonFocus = function (e) {
        //b.blur();
        if (this.udf.modePersonnel == 'normal')
            e.target.blur();
    };
    FrmPersonnelComponent.prototype.setGender = function (gender) {
        if (this.udf.modePersonnel == 'normal')
            return;
        this.oplandb.currPersonnel.gender = gender;
    };
    FrmPersonnelComponent.prototype.toggleXL = function () {
        if (this.udf.modePersonnel == 'normal')
            return;
        this.oplandb.currPersonnel.xl = !this.oplandb.currPersonnel.xl;
    };
    FrmPersonnelComponent.prototype.changeRecord = function (evt) {
        this.loadLookups = true;
        this.oplandb.dataActionClicked(evt);
    };
    FrmPersonnelComponent.prototype.toggleRequired = function (row) {
        if (this.udf.modePersonnel == 'normal' || this.oplandb.currPersonnel.id == -1 || row.mandatory == true)
            return;
        row.required = !row.required;
    };
    FrmPersonnelComponent.prototype.changeDate = function (row, fld) {
    };
    FrmPersonnelComponent.prototype.certSource = function () {
        return this.oplandb.datCertificates.slice(1, 10);
    };
    FrmPersonnelComponent.prototype.openCalendar = function (id) {
        __WEBPACK_IMPORTED_MODULE_3__app_globals__["setFocus"](id);
    };
    FrmPersonnelComponent.prototype.displayExpiry = function (cert) {
        var ret = '';
        if (cert.expiry) {
            ret = cert.expiry;
        }
        else if (cert.issued) {
            //let c:any=this.oplandb.datCertificates.find(e=>{ return e.id==cert.id});
            //if(c){
            // check if frequency is set
            //  console.log("Certificate:")
            //  console.log(c)
            //}
            if (cert.frequency != 0) {
                ret = "( " + __WEBPACK_IMPORTED_MODULE_3__app_globals__["dateToString"](__WEBPACK_IMPORTED_MODULE_3__app_globals__["addDays"](__WEBPACK_IMPORTED_MODULE_3__app_globals__["stringToDate"](__WEBPACK_IMPORTED_MODULE_3__app_globals__["longToShortDate"](cert.issued)), 365.25 * cert.frequency), true) + " )";
            }
        }
        return ret;
    };
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", typeof (_a = typeof __WEBPACK_IMPORTED_MODULE_1__models_data_models__["d" /* dataPersonnel */] !== "undefined" && __WEBPACK_IMPORTED_MODULE_1__models_data_models__["d" /* dataPersonnel */]) === "function" && _a || Object)
    ], FrmPersonnelComponent.prototype, "currPersonnel", void 0);
    FrmPersonnelComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-frm-personnel',
            template: __webpack_require__("../../../../../src/app/components/frm-personnel/frm-personnel.component.html"),
            styles: [__webpack_require__("../../../../../src/app/components/frm-personnel/frm-personnel.component.css")]
        }),
        __metadata("design:paramtypes", [typeof (_b = (typeof __WEBPACK_IMPORTED_MODULE_3__app_globals__ !== "undefined" && __WEBPACK_IMPORTED_MODULE_3__app_globals__).AppDataService) === "function" && _b || Object, typeof (_c = typeof __WEBPACK_IMPORTED_MODULE_5__app_oplandata__["a" /* AppOplpanDataService */] !== "undefined" && __WEBPACK_IMPORTED_MODULE_5__app_oplandata__["a" /* AppOplpanDataService */]) === "function" && _c || Object, typeof (_d = (typeof __WEBPACK_IMPORTED_MODULE_4__app_udfs__ !== "undefined" && __WEBPACK_IMPORTED_MODULE_4__app_udfs__).AppUDFService) === "function" && _d || Object])
    ], FrmPersonnelComponent);
    return FrmPersonnelComponent;
    var _a, _b, _c, _d;
}());

//# sourceMappingURL=frm-personnel.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/frm-settings/frm-settings.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/frm-settings/frm-settings.component.html":
/***/ (function(module, exports) {

module.exports = "<div class=\"card\" style=\"height:100%;\" id=\"settingsCard\" (window:resize)=\"onResize($event)\">\n    <div class=\"card-header row py-0\">\n        <div class=\"col float-left py-2\">\n            <h4 class=\"card-title\">Settings</h4>\n        </div>\n    </div>\n  <div class=\"card-block py-2 px-5\" style=\"overflow-y:auto;position:relative;\" id=\"myBLOCK\">\n    <div *ngIf=\"oplandb.msgTempPrompt.length\" class=\"alert alert-danger row\"\n      role=\"alert\"\n      style=\"position:relative;top:0px;left:-5px;width:calc(100% + 0px);position:absolute;z-index:10;left:15px;\">\n      <strong>Invald data:&nbsp;&nbsp;&nbsp; </strong>&nbsp;&nbsp;{{oplandb.msgTempPrompt}}\n    </div>    \n\n    <div *ngIf=\"oplandb.msgRequestPrompt.length\" class=\"alert alert-info row\"\n      role=\"alert\"\n      style=\"position:relative;top:0px;left:-5px;width:calc(100% + 0px);position:absolute;z-index:10;left:15px;\">\n      <strong>Processing:&nbsp;&nbsp;&nbsp; </strong>&nbsp;&nbsp;{{oplandb.msgRequestPrompt}}\n    </div>    \n\n    <div *ngIf=\"oplandb.msgTempInfoPrompt.length\" class=\"alert alert-success row\"\n      role=\"alert\"\n      style=\"position:relative;top:0px;left:-5px;width:calc(100% + 0px);position:absolute;z-index:10;left:15px;\">\n      <strong>Information:&nbsp;&nbsp;&nbsp; </strong>&nbsp;&nbsp;{{oplandb.msgTempInfoPrompt}}\n    </div>    \n  \n    <form>\n      <div class=\"row\" [style.height]=\"(blockHeight)+'px'\">\n        <div class=\"col p-0 m-0 px-1\">\n          <div class=\"card row px-0 mx-1\">\n            <div class=\"card-header text-center\">\n                <h6>{{oplandb.config.stdPOBLimit.title}}</h6>\n            </div>\n            <div class=\"card-block p-0\">\n              <table  class=\"table table-sm table-hover table-bordered tbc m-0\">\n                <thead>\n                  <tr><th class=\"border-0\">Date</th><th class=\"border-0\">Limit</th></tr>\n                </thead>\n                <tbody  [style.height]=\"tblHalfHeight()+'px'\">\n                    <tr *ngFor=\"let lim of oplandb.config.stdPOBLimit.limits;let idx=index;\" \n                      (mouseenter)=\"mouseEnter('SPOB',idx)\" (mouseleave)=\"mouseLeave()\"  >\n                      <td class=\"text-center\">{{showLongDate(lim.setDate)}}</td>\n                      <td class=\"text-center\" style=\"position:relative;\">{{lim.value}}\n                        <i \n                        (click)=\"removeSelectedConfig('SPOB', idx);\"\n                        [ngStyle]=\"getDeleteStyle('SPOB', idx)\"\n                        [title]=\"getDeleteTitle('SPOB', idx)\"\n                        class=\"fa fa-times-circle float-right\" aria-hidden=\"true\" \n                        style=\"width:20px;font-size:1.2em;\"></i>\n                    </td>\n                    </tr>\n                </tbody>\n                <tfoot *ngIf=\"oplandb.logInfo.settingsAdmin\">\n                    <tr style=\"display:flex;\">\n                      <td style=\"flex-grow:1;\" class=\"text-center p-0\">\n                        <div class=\"input-group input-group-sm mb-2 mr-sm-2 mb-sm-0\" >\n                            <input class=\"form-control text-center my-0 border-0\" \n                              type=\"text\" \n                              id=\"stdPOBLimits\"\n                              name=\"stdPOBsetDate\"\n                              onfocus=\"this.blur();\" \n                              [disabled]=\"oplandb.isPosting\"\n                              [(ngModel)]=\"oplandb.config.stdPOBLimit.newLimit.setDate\" \n                              [dpDayPicker]=\"udf.datePickerConfig\"\n                              placeholder=\"Select date...\"\n                              >\n                            <div class=\"input-group-addon border-1\">\n                              <a href=\"#\" (click)=\"showCalendar('stdPOBLimits')\">\n                                <i class=\"fa fa-calendar text-primary\"></i>\n                              </a>\n                            </div>\n                          </div>                    \n                      </td>\n                      <td style=\"flex-grow:1;\" class=\"text-center p-0\">\n                        <input class=\"form-control form-control-sm m-0 border-0 text-center\" \n                        placeholder=\"Enter limit...\"\n                        [disabled]=\"oplandb.isPosting\"\n                        name=\"stdPOBValue\"\n                        [(ngModel)]=\"oplandb.config.stdPOBLimit.newLimit.value\" \n                        type=\"text\">\n                        \n                      </td>\n                      <td class=\"p-0\" title=\"Add new standard POB limit ...\" style=\"width:25px;cursor:pointer;\">\n                        <button class=\"btn btn-link m-0\" type=\"button\" \n                          style=\"font-size:18px;height:22px;padding:0px;padding-left:2px;cursor:pointer;\"\n                          [disabled]=\"oplandb.isPosting\" (click)=\"oplandb.addConfigValue('SPOB')\">\n                          <i class=\"fa fa-plus p-0 m-0 text-info\" aria-hidden=\"true\"></i>\n                        </button>\n                      </td>\n                          \n                    </tr>\n                </tfoot>\n              </table>\n            </div>\n          </div>\n          <div class=\"card row px-0 mx-1 my-1\">\n            <div class=\"card-header text-center\">\n                <h6>{{oplandb.config.upmPOBLimit.title}}</h6>\n            </div>\n            <div class=\"card-block p-0\">\n              <table  class=\"table table-sm table-hover table-bordered tbc m-0\">\n                <thead>\n                <tr><th class=\"border-0\">Date</th><th class=\"border-0\">Limit</th></tr>\n                </thead>\n                <tbody  [style.height]=\"tblHalfHeight()+'px'\">\n                  <tr *ngFor=\"let lim of oplandb.config.upmPOBLimit.limits;let idx=index;\"\n                  (mouseenter)=\"mouseEnter('UPOB',idx)\" (mouseleave)=\"mouseLeave()\">\n                      <td class=\"text-center\">{{showLongDate(lim.setDate)}}</td>\n                      <td class=\"text-center\">{{lim.value}}\n                        <i \n                        (click)=\"removeSelectedConfig('UPOB', idx);\"\n                        [ngStyle]=\"getDeleteStyle('UPOB', idx)\"\n                        [title]=\"getDeleteTitle('UPOB', idx)\"\n                        class=\"fa fa-times-circle float-right\" aria-hidden=\"true\" \n                        style=\"width:20px;font-size:1.2em;\"></i>\n                  \n                  </td>\n                  </tr>\n                </tbody>\n                <tfoot  *ngIf=\"oplandb.logInfo.settingsAdmin\">\n                    <tr style=\"display:flex;\">\n                      <td style=\"flex-grow:1;\" class=\"text-center p-0\">\n                        <div class=\"input-group input-group-sm mb-2 mr-sm-2 mb-sm-0\" >\n                            <input class=\"form-control text-center my-0 border-0\" \n                              type=\"text\" \n                              id=\"upmPOBLimits\"\n                              name=\"upmPOBDate\"\n                              onfocus=\"this.blur();\" \n                              [disabled]=\"oplandb.isPosting\"\n                              [(ngModel)]=\"oplandb.config.upmPOBLimit.newLimit.setDate\" \n                              [dpDayPicker]=\"udf.datePickerConfig\"\n                              placeholder=\"Select date...\"\n                              >\n                            <div class=\"input-group-addon border-1\">\n                              <a href=\"#\" (click)=\"showCalendar('upmPOBLimits')\">\n                                <i class=\"fa fa-calendar text-primary\"></i>\n                              </a>\n                            </div>\n                          </div>                    \n                      </td>\n                      <td style=\"flex-grow:1;\" class=\"text-center p-0\">\n                        <input class=\"form-control form-control-sm m-0 border-0 text-center\" \n                        placeholder=\"Enter limit...\"\n                        name=\"upmPOBValue\"\n                        [disabled]=\"oplandb.isPosting\"\n                        [(ngModel)]=\"oplandb.config.upmPOBLimit.newLimit.value\" \n                        type=\"text\">\n                        \n                      </td>\n                      <td class=\"p-0\" title=\"Add new upmanning POB limit ...\" style=\"width:25px;cursor:pointer;\">\n                        <button class=\"btn btn-link m-0\" type=\"button\" \n                          style=\"font-size:18px;height:22px;padding:0px;padding-left:2px;cursor:pointer;\"\n                          [disabled]=\"oplandb.isPosting\" (click)=\"oplandb.addConfigValue('UPOB')\">\n                          <i class=\"fa fa-plus p-0 m-0 text-info\" aria-hidden=\"true\"></i>\n                        </button>\n                      </td>                      \n                    </tr>\n                </tfoot>\n                </table>\n            </div>\n          </div>\n        </div>\n        <div class=\"card col px-0 mx-1\">\n          <div class=\"card-header text-center\">\n              <h6>{{oplandb.config.upmanPeriods.title}}</h6>\n          </div>\n          <div class=\"card-block p-0\">\n            <table  class=\"table table-sm table-hover table-bordered tbc m-0\">\n              <thead>\n                <tr><th class=\"border-0\">Start Date</th><th class=\"border-0\">End Date</th></tr>\n              </thead>\n              <tbody  [style.height]=\"(tblFullHeight())+'px'\">\n                <tr *ngFor=\"let per of oplandb.config.upmanPeriods.periods;let idx=index;\"\n                (mouseenter)=\"mouseEnter('UPER',idx)\" (mouseleave)=\"mouseLeave()\">\n                  <td class=\"text-center\">{{showLongDate(per.startDate)}}</td>\n                  <td class=\"text-center\">{{showLongDate(per.endDate)}}\n                    <i \n                    (click)=\"removeSelectedConfig('UPER', idx);\"\n                    [ngStyle]=\"getDeleteStyle('UPER', idx)\"\n                    [title]=\"getDeleteTitle('UPER', idx)\"\n                    class=\"fa fa-times-circle float-right\" aria-hidden=\"true\" \n                    style=\"width:20px;font-size:1.2em;\"></i>\n          \n                    </td></tr>\n              </tbody>\n              <tfoot *ngIf=\"oplandb.logInfo.settingsAdmin\">\n                  <tr style=\"display:flex;\">\n                    <td style=\"flex-grow:1;\" class=\"text-center p-0\">\n                      <div class=\"input-group input-group-sm mb-2 mr-sm-2 mb-sm-0\" >\n                          <input class=\"form-control text-center my-0 border-0\" \n                            type=\"text\" \n                            id=\"upmPerStart\"\n                            name=\"startDate\"\n                            onfocus=\"this.blur();\" \n                            [disabled]=\"oplandb.isPosting\"\n                            [(ngModel)]=\"oplandb.config.upmanPeriods.newPeriod.startDate\" \n                            [dpDayPicker]=\"udf.datePickerConfig\"\n                            placeholder=\"Select date...\"\n                            >\n                          <div class=\"input-group-addon border-1\">\n                            <a href=\"#\" (click)=\"showCalendar('upmPerStart')\">\n                              <i class=\"fa fa-calendar text-primary\"></i>\n                            </a>\n                          </div>\n                        </div>                    \n                    </td>\n\n                    <td style=\"flex-grow:1;\" class=\"text-center p-0\">\n                      <div class=\"input-group input-group-sm mb-2 mr-sm-2 mb-sm-0\" >\n                          <input class=\"form-control text-center my-0 border-0\" \n                            type=\"text\" \n                            id=\"upmPerEnd\"\n                            name=\"endDate\"\n                            onfocus=\"this.blur();\" \n                            [disabled]=\"oplandb.isPosting\"\n                            [(ngModel)]=\"oplandb.config.upmanPeriods.newPeriod.endDate\" \n                            [dpDayPicker]=\"udf.datePickerConfig\"\n                            placeholder=\"Select date...\"\n                            >\n                          <div class=\"input-group-addon border-1\">\n                            <a href=\"#\" (click)=\"showCalendar('upmPerEnd')\">\n                              <i class=\"fa fa-calendar text-primary\"></i>\n                            </a>\n                          </div>\n                        </div>                    \n                    </td>\n  \n                    <td class=\"p-0\" title=\"Add upmanning period ...\" style=\"width:25px;cursor:pointer;\">\n                      <button class=\"btn btn-link m-0\" type=\"button\" \n                        style=\"font-size:18px;height:22px;padding:0px;padding-left:2px;cursor:pointer;\"\n                        [disabled]=\"oplandb.isPosting\" (click)=\"oplandb.addConfigValue('UPER')\">\n                        <i class=\"fa fa-plus p-0 m-0 text-info\" aria-hidden=\"true\"></i>\n                      </button>\n                    </td>\n                  </tr>\n                </tfoot>\n            </table>\n          </div>\n        </div>\n\n        <div class=\"col p-0 mx-1\">\n\n          <div class=\"card row px-0 mx-1\">\n            <div class=\"card-header text-center\">\n                <h6>{{oplandb.config.mobLimit.title}}</h6>\n            </div>\n            <div class=\"card-block p-0\">\n              <table  class=\"table table-sm table-hover table-bordered tbc m-0\">\n                <thead>\n                  <tr><th class=\"border-0\">Date</th><th class=\"border-0\">Limit</th></tr>\n                </thead>\n                <tbody  [style.height]=\"tblHalfHeight()+'px'\">\n                    <tr *ngFor=\"let lim of oplandb.config.mobLimit.limits;let idx=index;\"\n                    (mouseenter)=\"mouseEnter('MLIM',idx)\" (mouseleave)=\"mouseLeave()\">\n                        <td class=\"text-center\">{{showLongDate(lim.setDate)}}</td>\n                        <td class=\"text-center\">{{lim.value}}\n                          <i \n                          (click)=\"removeSelectedConfig('MLIM', idx);\"\n                          [ngStyle]=\"getDeleteStyle('MLIM', idx)\"\n                          [title]=\"getDeleteTitle('MLIM', idx)\"\n                          class=\"fa fa-times-circle float-right\" aria-hidden=\"true\" \n                          style=\"width:20px;font-size:1.2em;\"></i>\n  \n                        </td>\n                    </tr>\n                  </tbody>\n                  <tfoot *ngIf=\"oplandb.logInfo.settingsAdmin\">\n                    <tr style=\"display:flex;\">\n                      <td style=\"flex-grow:1;\" class=\"text-center p-0\">\n                        <div class=\"input-group input-group-sm mb-2 mr-sm-2 mb-sm-0\" >\n                            <input class=\"form-control text-center my-0 border-0\" \n                              type=\"text\" \n                              id=\"mobLimits\"\n                              name=\"mobDate\"\n                              onfocus=\"this.blur();\" \n                              [disabled]=\"oplandb.isPosting\"\n                              [(ngModel)]=\"oplandb.config.mobLimit.newLimit.setDate\" \n                              [dpDayPicker]=\"udf.datePickerConfig\"\n                              placeholder=\"Select date...\"\n                              >\n                            <div class=\"input-group-addon border-1\">\n                              <a href=\"#\" (click)=\"showCalendar('mobLimits')\">\n                                <i class=\"fa fa-calendar text-primary\"></i>\n                              </a>\n                            </div>\n                          </div>                    \n                      </td>\n                      <td style=\"flex-grow:1;\" class=\"text-center p-0\">\n                        <input class=\"form-control form-control-sm m-0 border-0 text-center\" \n                        placeholder=\"Enter limit...\"\n                        name=\"mobValue\"\n                        [disabled]=\"oplandb.isPosting\"\n                        [(ngModel)]=\"oplandb.config.mobLimit.newLimit.value\" \n                        type=\"text\">\n                        \n                      </td>\n\n                      <td class=\"p-0\" title=\"Add new Mob limit ...\" style=\"width:25px;cursor:pointer;\">\n                        <button class=\"btn btn-link m-0\" type=\"button\" \n                          style=\"font-size:18px;height:22px;padding:0px;padding-left:2px;cursor:pointer;\"\n                          [disabled]=\"oplandb.isPosting\" (click)=\"oplandb.addConfigValue('MLIM')\">\n                          <i class=\"fa fa-plus p-0 m-0 text-info\" aria-hidden=\"true\"></i>\n                        </button>\n                      </td>                      \n                    </tr>\n                  </tfoot>\n              </table>\n            </div>\n          </div>\n\n          <div class=\"card row px-0 mx-1 my-1\">\n            <div class=\"card-header text-center\">\n                <h6>{{oplandb.config.demobLimit.title}}</h6>\n            </div>\n            <div class=\"card-block p-0\">\n              <table  class=\"table table-sm table-hover table-bordered tbc m-0\">\n                <thead>\n                  <tr><th class=\"border-0\">Date</th><th class=\"border-0\">Limit</th></tr>\n                </thead>\n                <tbody  [style.height]=\"tblHalfHeight()+'px'\">\n                  <tr *ngFor=\"let lim of oplandb.config.demobLimit.limits;let idx=index;\"\n                  (mouseenter)=\"mouseEnter('DLIM',idx)\" (mouseleave)=\"mouseLeave()\">\n                      <td class=\"text-center\">{{showLongDate(lim.setDate)}}</td>\n                      <td class=\"text-center\">{{lim.value}}\n                          <i \n                          (click)=\"removeSelectedConfig('DLIM', idx);\"\n                          [ngStyle]=\"getDeleteStyle('DLIM', idx)\"\n                          [title]=\"getDeleteTitle('DLIM', idx)\"\n                          class=\"fa fa-times-circle float-right\" aria-hidden=\"true\" \n                          style=\"width:20px;font-size:1.2em;\"></i>\n                        <!--i *ngIf=\"currRowID==('DLIM' + idx) && !oplandb.isPosting\" \n                          style=\"cursor:pointer;font-size:1.2em;width:20px;\"\n                          title=\"Remove entry...\"\n                          \n                          class=\"fa fa-times-circle text-danger float-right\" \n                          aria-hidden=\"true\"></i--> \n                      \n                      </td>\n                  </tr>\n                </tbody>\n                <tfoot *ngIf=\"oplandb.logInfo.settingsAdmin\">\n                  <tr style=\"display:flex;\">\n                    <td style=\"flex-grow:1;\" class=\"text-center p-0\">\n                      <div class=\"input-group input-group-sm mb-2 mr-sm-2 mb-sm-0\" >\n                          <input class=\"form-control text-center my-0 border-0\" \n                            type=\"text\" \n                            id=\"demobLimits\"\n                            name=\"demobDate\"\n                            onfocus=\"this.blur();\" \n                            [disabled]=\"oplandb.isPosting\"\n                            [(ngModel)]=\"oplandb.config.demobLimit.newLimit.setDate\" \n                            [dpDayPicker]=\"udf.datePickerConfig\"\n                            placeholder=\"Select date...\"\n                            >\n                          <div class=\"input-group-addon border-1\">\n                            <a href=\"#\" (click)=\"showCalendar('demobLimits')\">\n                              <i class=\"fa fa-calendar text-primary\"></i>\n                            </a>\n                          </div>\n                        </div>                    \n                    </td>\n                    <td style=\"flex-grow:1;\" class=\"text-center p-0\">\n                      <input class=\"form-control form-control-sm m-0 border-0 text-center\" \n                      placeholder=\"Enter limit...\"\n                      [disabled]=\"oplandb.isPosting\"\n                      name=\"demobValue\"\n                      [(ngModel)]=\"oplandb.config.demobLimit.newLimit.value\" \n                      type=\"text\">\n                      \n                    </td>\n\n                    <td class=\"p-0\" title=\"Add new Demob limit ...\" style=\"width:25px;cursor:pointer;\">\n                      <button class=\"btn btn-link m-0\" type=\"button\" \n                        style=\"font-size:18px;height:22px;padding:0px;padding-left:2px;cursor:pointer;\"\n                        [disabled]=\"oplandb.isPosting\" (click)=\"oplandb.addConfigValue('DLIM')\">\n                        <i class=\"fa fa-plus p-0 m-0 text-info\" aria-hidden=\"true\"></i>\n                      </button>\n                    </td>                      \n                    \n                  </tr>\n                </tfoot>\n              </table>\n            </div>\n          </div>\n        \n        </div>\n\n      </div>\n    </form>\n  </div>\n  <div class=\"card-footer\" *ngIf=\"false\">\n    <div class=\"form-buttons align-bottom\" >\n      ...\n    </div>\n  </div>   \n</div>"

/***/ }),

/***/ "../../../../../src/app/components/frm-settings/frm-settings.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return FrmSettingsComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1_jquery__ = __webpack_require__("../../../../jquery/dist/jquery.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1_jquery___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_1_jquery__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__app_globals__ = __webpack_require__("../../../../../src/app/app.globals.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3__app_udfs__ = __webpack_require__("../../../../../src/app/app.udfs.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_4__app_oplandata__ = __webpack_require__("../../../../../src/app/app.oplandata.ts");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};





var FrmSettingsComponent = (function () {
    function FrmSettingsComponent(udf, oplandb) {
        this.udf = udf;
        this.oplandb = oplandb;
        this.currRowID = "";
        this.blockHeight = 500;
    }
    FrmSettingsComponent.prototype.ngOnInit = function () {
        var _this = this;
        this.oplandb.clearAllSubscription();
        setTimeout(function () { _this.onResize(null); }, 50);
        this.oplandb.getConfig();
    };
    FrmSettingsComponent.prototype.onResize = function (e) {
        this.blockHeight = __WEBPACK_IMPORTED_MODULE_1_jquery__('#myBLOCK').height();
    };
    FrmSettingsComponent.prototype.tblFullHeight = function () {
        return this.blockHeight - 115 + (!this.oplandb.logInfo.settingsAdmin ? 32 : 0);
    };
    FrmSettingsComponent.prototype.tblHalfHeight = function () {
        return ((this.blockHeight - 115 * 2) / 2 - 2) + (!this.oplandb.logInfo.settingsAdmin ? 32 : 0);
    };
    FrmSettingsComponent.prototype.showLongDate = function (dt, initVal) {
        if (__WEBPACK_IMPORTED_MODULE_2__app_globals__ == undefined)
            return "";
        if (initVal == undefined)
            initVal = "01 Jan 1980";
        var ret = __WEBPACK_IMPORTED_MODULE_2__app_globals__["dateToString"](dt, true);
        return ret == initVal ? "( Initial )" : ret;
    };
    FrmSettingsComponent.prototype.mouseEnter = function (type, id) {
        this.currRowID = type + id;
    };
    FrmSettingsComponent.prototype.mouseLeave = function () {
        this.currRowID = "";
    };
    FrmSettingsComponent.prototype.showCalendar = function (id) {
        __WEBPACK_IMPORTED_MODULE_2__app_globals__["setFocus"](id);
    };
    FrmSettingsComponent.prototype.removeSelectedConfig = function (type, idx) {
        var cfgObj = this.oplandb.getConfigObject(type);
        var isNoDelete = false;
        if (cfgObj == undefined)
            return;
        if (type != 'UPER') {
            // suppress deletion of inital values
            var lastIndex = cfgObj['limits'].length - 1;
            if (lastIndex != idx) {
                // not the last item of the limits array
                isNoDelete = ((type + idx) != this.currRowID);
            }
        }
        if (!isNoDelete) {
            this.oplandb.removeConfig(type + idx);
        }
        else {
            alert('not deleting...');
        }
    };
    FrmSettingsComponent.prototype.getDeleteStyle = function (type, idx) {
        var cfgObj = this.oplandb.getConfigObject(type);
        var isNoDelete = true;
        if (type != 'UPER' && this.oplandb.logInfo.settingsAdmin) {
            var lastIndex = cfgObj['limits'].length - 1;
            if (lastIndex != idx) {
                // not the last item of the limits array
                isNoDelete = ((type + idx) != this.currRowID);
            }
        }
        else if (type == 'UPER' && this.oplandb.logInfo.settingsAdmin) {
            isNoDelete = ((type + idx) != this.currRowID);
        }
        return {
            'color': isNoDelete ? 'transparent' : this.udf.defColorDanger,
            'cursor': isNoDelete ? 'default' : 'pointer'
        };
    };
    FrmSettingsComponent.prototype.getDeleteTitle = function (type, idx) {
        return '';
    };
    FrmSettingsComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-frm-settings',
            template: __webpack_require__("../../../../../src/app/components/frm-settings/frm-settings.component.html"),
            styles: [__webpack_require__("../../../../../src/app/components/frm-settings/frm-settings.component.css")]
        }),
        __metadata("design:paramtypes", [typeof (_a = (typeof __WEBPACK_IMPORTED_MODULE_3__app_udfs__ !== "undefined" && __WEBPACK_IMPORTED_MODULE_3__app_udfs__).AppUDFService) === "function" && _a || Object, typeof (_b = typeof __WEBPACK_IMPORTED_MODULE_4__app_oplandata__["a" /* AppOplpanDataService */] !== "undefined" && __WEBPACK_IMPORTED_MODULE_4__app_oplandata__["a" /* AppOplpanDataService */]) === "function" && _b || Object])
    ], FrmSettingsComponent);
    return FrmSettingsComponent;
    var _a, _b;
}());

//# sourceMappingURL=frm-settings.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/frm-table-panels/frm-table-panels.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/frm-table-panels/frm-table-panels.component.html":
/***/ (function(module, exports) {

module.exports = "\n<div id=\"leftHeader\" style=\"position:absolute;top:0px;left:0px;background:yellow;overflow:hidden;\"\n[style.width]=\"vertSeparatorPos + 'px'\"\n[style.height]=\"headSeparatorPos + 'px'\"\n(window:resize)=\"onResize($event)\"\n>\n<div style=\"background:mediumseagreen;width:800px;height:900px;\">\n    {{contWidth() + ', ' + contHeight()}}\n</div>\n</div>\n<div id=\"leftDetails\" style=\"position:absolute;left:0px;background:cyan;overflow:auto;\"\n[style.top]=\"headSeparatorPos + 'px'\"\n[style.width]=\"vertSeparatorPos + 'px'\"\n[style.height]=\"(windowHeight- headSeparatorPos) + 'px'\"\n>\n    <div style=\"background:moccasin;width:800px;height:4000px;\">\n        left details\n    </div>\n</div>\n<div id=\"rightHeader\" style=\"position:absolute;top:0px;background:lime;overflow:hidden;\"\n[style.left]=\"vertSeparatorPos + 'px'\"\n[style.width]=\"(windowWidth-vertSeparatorPos) + 'px'\"\n[style.height]=\"headSeparatorPos + 'px'\"\n>\n    <div style=\"background: mediumturquoise;width:4000px;height:200px;white-space:nowrap;\">\n        The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog. \n    </div>\n</div>\n<div id=\"rightDetails\" \nstyle=\"position:absolute;background:magenta;overflow:auto;white-space:normal;\"\n[style.top]=\"headSeparatorPos + 'px'\"\n[style.left]=\"vertSeparatorPos + 'px'\"\n[style.width]=\"(windowWidth-vertSeparatorPos-1) + 'px'\"\n[style.height]=\"(windowHeight- headSeparatorPos) + 'px'\"\n(scroll)=\"onElementScroll($event)\"\n>\n<div style=\"width:2048px;height:4000px;background:beige;white-space:nowrap;\">\nThe quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog. \n</div>\n</div>"

/***/ }),

/***/ "../../../../../src/app/components/frm-table-panels/frm-table-panels.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return FrmTablePanelsComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1_jquery__ = __webpack_require__("../../../../jquery/dist/jquery.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1_jquery___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_1_jquery__);
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};


var FrmTablePanelsComponent = (function () {
    function FrmTablePanelsComponent() {
        this.vertSeparatorPos = 400;
        this.headSeparatorPos = 100;
        this.divMain = __WEBPACK_IMPORTED_MODULE_1_jquery__('#d');
    }
    FrmTablePanelsComponent.prototype.contWidth = function () {
        //return $(document).width();
        return window.innerWidth;
    };
    FrmTablePanelsComponent.prototype.contHeight = function () {
        return __WEBPACK_IMPORTED_MODULE_1_jquery__(document).height();
    };
    FrmTablePanelsComponent.prototype.ngOnInit = function () {
        this.divMain = __WEBPACK_IMPORTED_MODULE_1_jquery__('#d');
        var cont = __WEBPACK_IMPORTED_MODULE_1_jquery__('#d');
        this.setWindowSizeVars();
    };
    FrmTablePanelsComponent.prototype.iClick = function (e) {
        alert('hhh');
    };
    FrmTablePanelsComponent.prototype.onResize = function (e) {
        this.setWindowSizeVars();
    };
    FrmTablePanelsComponent.prototype.setWindowSizeVars = function () {
        this.windowWidth = __WEBPACK_IMPORTED_MODULE_1_jquery__(window).width();
        this.windowHeight = __WEBPACK_IMPORTED_MODULE_1_jquery__(window).height();
    };
    FrmTablePanelsComponent.prototype.onElementScroll = function (e) {
        if (e.srcElement.id == 'rightDetails') {
            __WEBPACK_IMPORTED_MODULE_1_jquery__('#rightHeader').scrollLeft(__WEBPACK_IMPORTED_MODULE_1_jquery__(e.srcElement).scrollLeft());
        }
        else if (e.srcElement.id == 'leftDetails') {
            __WEBPACK_IMPORTED_MODULE_1_jquery__('#leftHeader').scrollLeft(__WEBPACK_IMPORTED_MODULE_1_jquery__(e.srcElement).scrollLeft());
        }
    };
    FrmTablePanelsComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-frm-table-panels',
            template: __webpack_require__("../../../../../src/app/components/frm-table-panels/frm-table-panels.component.html"),
            styles: [__webpack_require__("../../../../../src/app/components/frm-table-panels/frm-table-panels.component.css")]
        }),
        __metadata("design:paramtypes", [])
    ], FrmTablePanelsComponent);
    return FrmTablePanelsComponent;
}());

//# sourceMappingURL=frm-table-panels.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/frm-team/frm-team.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/frm-team/frm-team.component.html":
/***/ (function(module, exports) {

module.exports = "<!--div style=\"position:absolute;top:0px;left:0px;z-index:10;width:20px;height:324px;background:cyan;\">\n        &nbsp;\n    </div>\n    <div style=\"position:absolute;top:0px;left:80px;z-index:10;width:20px;height:359px;background:lime;\">\n            &nbsp;\n        </div>\n        <div style=\"position:absolute;top:0px;left:30px;z-index:10;width:20px;height:84px;background:red;\">\n            &nbsp;\n        </div>    \n        <div style=\"position:absolute;top:0px;left:60px;z-index:10;width:20px;height:59px;background:yellow;\">\n                &nbsp;\n            </div-->            \n<div class=\"card\" style=\"height:100%;\" id=\"teamCard\" (window:resize)=\"onResize($event)\">\n    <div class=\"card-header row py-0\">\n        <div class=\"col float-left py-0\">\n            <h4 class=\"card-title\"><i style=\"font-size:2em;margin-top:5px;margin-right:20px;\" \n                class=\"fa fa-users text-muted\" aria-hidden=\"true\"></i>Team Details</h4>\n        </div>\n\n        <div class=\" btn-group-sm col m-0 float-right text-right align-text-bottom text-info\"\n        style=\"padding-top:20px;\">\n\n\n        <button \n          class=\"btn btn-link\" \n          [ngClass]=\"udf.acnBtnClass('add')\"\n          (click)=\"btnClick('ADD')\"\n          title=\"Add new team...\"\n          [disabled]=\"!disableControls('header')\"\n          *ngIf=\"oplandb.logInfo.teamAdd\"\n          >\n          <i id=\"per_add\" class=\"fa fa-plus-square-o text-success\" style=\"font-size:20px;cursor:pointer;\"></i>\n        </button>\n        <button class=\"btn btn-link\" \n          [ngClass]=\"udf.acnBtnClass('edit')\"\n          (click)=\"btnClick('EDIT')\"\n          title=\"Edit team details...\"\n          [disabled]=\"!disableControls('header')\"\n          *ngIf=\"oplandb.logInfo.teamEdit\"\n          >\n          <i id=\"per_edit\" class=\"fa fa-pencil-square-o text-info\" style=\"font-size:20px;cursor:pointer;\"></i>\n        </button>\n        <button class=\"btn btn-link\" \n          [ngClass]=\"udf.acnBtnClass('delete')\"\n          (click)=\"btnClick('DEL')\"\n          [disabled]=\"!disableControls('header')\"\n          title=\"Delete selected activity...\"\n          *ngIf=\"oplandb.logInfo.teamDelete\"\n        ><i id=\"per_delete\" class=\"fa fa-trash-o text-danger\" style=\"font-size:20px;cursor:pointer;\"></i></button>\n    </div>\n      \n        \n    </div>\n  <div class=\"card-block py-0\" style=\"overflow-y:auto;\" id=\"myBLOCK\">\n      <form>\n\n\n            <div *ngIf=\"dispAlert('delete')\"\n            class=\"alert alert-danger row my-0\" \n            role=\"alert\" style=\"position:relative;top:0px;left:-5px;width:calc(100% + 40px);\">\n            <strong>{{oplandb.alertYesNoMessage.label}}:</strong> &nbsp;{{oplandb.alertYesNoMessage.message}}&nbsp;\n            <div class=\"btn-group-sm\">\n              <button class=\"btn btn-danger\" (click)=\"btnClick('goDelete')\" style=\"cursor:pointer\">{{oplandb.alertYesNoMessage.pos}}</button>\n              <button class=\"btn btn-secondary\" (click)=\"btnClick('noDelete')\" style=\"cursor:pointer\">{{oplandb.alertYesNoMessage.neg}}</button>\n            </div>\n        </div>  \n        \n        <div *ngIf=\"udf.isTeamNew()\" class=\"alert alert-info row\"\n          role=\"alert\"\n          style=\"position:relative;top:0px;left:-5px;width:calc(100% + 40px);\">\n          <strong>Attention: </strong>&nbsp;Team members can only be added after the new team basic info has been saved.\n        </div>\n                \n        <div *ngIf=\"oplandb.msgTempPrompt.length\" class=\"alert alert-danger row\"\n          role=\"alert\"\n          style=\"position:relative;top:0px;left:-5px;width:calc(100% + 40px);\">\n          <strong>Reminder: </strong>&nbsp;{{oplandb.msgTempPrompt}}\n        </div>\n\n        <div *ngIf=\"oplandb.msgTempInfoPrompt.length\" class=\"alert alert-info row\"\n            role=\"alert\"\n            style=\"position:relative;top:0px;left:-5px;width:calc(100% + 40px);\">\n            <strong>Information: </strong>&nbsp;{{oplandb.msgTempInfoPrompt}}\n        </div>      \n          \n\n        <div *ngIf=\"oplandb.msgRequestPrompt.length\" class=\"alert alert-info row\"\n          role=\"alert\"\n          style=\"position:relative;top:0px;left:-5px;width:calc(100% + 40px);\">\n          <strong>Server Request: </strong>&nbsp;{{oplandb.msgRequestPrompt}}\n        </div>   \n\n        <div class=\"row\">\n            <div class=\"col\">\n          <div class=\"form-group my-0\">\n              <label class=\"form-label py-0 my-0\" for=\"exampleInputEmail1\">Team Name</label>\n              <input id=\"teamName\" type=\"text\" class=\"form-control form-control-sm\"\n                [disabled]=\"disableControls('header')\"\n                placeholder=\"Enter team name\"\n                [(ngModel)]=\"oplandb.currTeam.name\"\n                name=\"name\"\n              >\n          </div>\n      \n          <div class=\"form-group\">\n            <label for=\"exampleTextarea\" class=\"form-label py-0 my-0\">Description</label>\n            <textarea class=\"form-control form-control-sm\" rows=\"3\" \n            placeholder=\"Enter team description\"\n            [disabled]=\"disableControls('header')\"\n            [(ngModel)]=\"oplandb.currTeam.description\"\n            name=\"description\"\n            ></textarea>\n          </div>\n        </div>\n        <div style=\"width:150px;padding-top:25px;\" class=\"text-center\">\n            <label for=\"isCore\" class=\"form-label py-0 my-0\">Core team?</label><br>\n            <input id=\"isCore\" type=\"checkbox\" name=\"core\"\n                [(ngModel)]=\"oplandb.currTeam.core\"\n            [disabled]=\"disableControls('header')\"\n            ><br>\n            <label for=\"isUpMan\" class=\"form-label py-0 my-0\">Up-Manning?</label><br>\n            <input id=\"isUpMan\" type=\"checkbox\"\n                [(ngModel)]=\"oplandb.currTeam.upman\"  name=\"upman\"\n            [disabled]=\"disableControls('header')\"\n            >\n        </div>\n\n        </div>\n\n          <div class=\"row\">\n              <div class=\"form-group col\">\n                \n                <div class=\"row\">\n                    <label class=\"col\" for=\"exampleInputPassword1\">Members:</label>\n                        <div *ngIf=\"oplandb.logInfo.teamMemAdmin\">\n                            <div class=\"col text-right\">\n                                <app-ctl-user-select \n                                [width]=\"teamColsPX[0]+'px'\" \n                                [height]=\"250\" \n                                [(counter)]=\"selectedUser\" \n                    \n                                [lookup]=\"oplandb.idxPersonnel\"\n                                [displayField]=\"'fullName'\"\n                                [valueField]=\"'id'\"\n                    \n                                [disableSelect]=\"disableControls()\"\n                                \n                                [source]=\"oplandb.userSelection\"\n                                [groupName]=\"'name'\"\n                                [dataName]=\"'data'\"\n                                [valueColumn]=\"'id'\"\n                                [displayColumn]=\"\"\n                                (counterChange)=\"onUserChange($event)\"\n                                (itemSelected)=\"onUserItemSelected($event)\"\n                                \n                                >\n                                </app-ctl-user-select>  \n                                <button class=\"btn btn-link btn-sm\" \n                                    style=\"font-size:20px;cursor:pointer;\"\n                                    title=\"Add selected user as member\"\n                                    (click)=\"addMemberClick()\"\n                                    [disabled]=\"!disableControls('header') || selectedUser==0\"\n                                >\n                                <i class=\"fa fa-plus text-info\"></i>\n                                </button>\n                        </div>\n                </div>               \n            </div>\n            \n                  <table id=\"team\" class=\"table table-sm table-hover table-bordered tbc my-0\">\n                    <thead>\n                        <tr>\n                          <th [style.width]=\"teamColsPX[0]+'px'\" >Name</th>\n                          <th [style.width]=\"teamColsPX[1]+'px'\">Company</th>\n                          <th [style.width]=\"teamColsPX[2]+'px'\">Position</th>\n                          <th *ngIf=\"oplandb.logInfo.teamMemAdmin\" [style.width]=\"teamColsPX[3]+'px'\">Del</th>\n                        </tr>\n                    </thead>\n                    <tbody [style.height]=\"teamBodyHeight+'px'\">\n                        <div *ngFor=\"let mem of oplandb.currTeam.members;trackBy:oplandb.trackByItem;\">\n                        <tr>\n                          <td [style.width]=\"teamColsPX[0]+'px'\">{{getTeamMemberDetails(mem,'name')}}</td>\n                          <td [style.width]=\"teamColsPX[1]+'px'\" style=\"text-align:left;\">{{getTeamMemberDetails(mem,'coy')}}</td>\n                          <td [style.width]=\"teamColsPX[2]+'px'\" style=\"text-align:left;\">{{getTeamMemberDetails(mem,'pos')}}</td>\n\n                          <td  *ngIf=\"oplandb.logInfo.teamMemAdmin\"\n                            [style.width]=\"teamColsPX[3]+'px'\" style=\"text-align:center;\" class=\"acn\">\n                              <button \n                                [disabled]=\"!disableControls('header')\"\n                                class=\"btn btn-link btn-sm text-danger\"\n                                (click)=\"btnClick('del',mem)\"\n                                style=\"cursor:pointer;\"\n                                >\n                                  <i class=\"fa fa-times\"></i></button>\n                          </td>\n                        </tr>\n                      </div>\n                    </tbody>\n                  </table>\n\n              </div>\n            </div>\n              \n            <!--div style=\"display:flex;flex-direction:row;background:magenta;width:100%;height:150px;\">\n                <div style=\"background: red;flex-grow:0.5\">1</div>\n                <div style=\"background: lime;flex-grow:1\">2</div>\n                <div style=\"background: yellow;flex-grow:2\">3</div>\n                <div style=\"background: blue;flex-grow:3\">4</div>\n                <div style=\"background: cyan;flex-grow:4\">5</div>\n            </div-->\n      </form>\n  </div>\n  <div class=\"card-footer\" *ngIf=\"!disableControls('header') && !udf.isTeamDeleteAlert() && !udf.isTeamMemberDeleteAlert()\">\n    <div class=\"form-buttons align-bottom\" >\n        <button type=\"button\" class=\"btn btn-secondary\"\n        [disabled]=\"disableControls('header')\"\n        (click)=\"btnClick('CANCEL')\">Cancel</button>\n      <button type=\"button\" class=\"btn btn-primary\"\n        [disabled]=\"disableControls('header')\"\n        (click)=\"btnClick('SAVE')\" >Save</button>\n    </div>\n</div>   \n</div>"

/***/ }),

/***/ "../../../../../src/app/components/frm-team/frm-team.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return FrmTeamComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__models_data_models__ = __webpack_require__("../../../../../src/app/models/data-models.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2_jquery__ = __webpack_require__("../../../../jquery/dist/jquery.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2_jquery___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_2_jquery__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3__app_globals__ = __webpack_require__("../../../../../src/app/app.globals.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_4__app_udfs__ = __webpack_require__("../../../../../src/app/app.udfs.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_5__app_oplandata__ = __webpack_require__("../../../../../src/app/app.oplandata.ts");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};






var FrmTeamComponent = (function () {
    function FrmTeamComponent(udf, oplandb) {
        this.udf = udf;
        this.oplandb = oplandb;
        this.teamCols = ["40%", "30%", "30%", 40];
        this.teamBodyHeight = 150;
        this.selectedUser = 0;
    }
    FrmTeamComponent.prototype.ngOnInit = function () {
        this.oplandb.itemsGroupings = [];
        this.oplandb.clearAllSubscription();
        this.oplandb.getCredentials();
        this.oplandb.setOtherNavGroup();
        this.setTableColWidths();
        this.udf.isTeamNormal(true);
    };
    FrmTeamComponent.prototype.ngOnDestroy = function () {
        this.oplandb.clearSubscription(this.subsProcessData);
    };
    FrmTeamComponent.prototype.disableControls = function (mode) {
        var retVal = false;
        if (mode == undefined)
            mode = '';
        if (mode == 'header') {
            retVal = this.udf.isTeamNormal();
            //    }else{
            //      retVal = this.udf.isTeamMemberNormal();
        }
        return retVal || this.oplandb.isRequesting() || this.alertModal();
    };
    FrmTeamComponent.prototype.setTableColWidths = function () {
        var ftr = __WEBPACK_IMPORTED_MODULE_2_jquery__('.card-footer');
        var footOff = 0;
        if (ftr.length)
            footOff = ftr.height();
        this.teamBodyHeight = this.udf.getReactiveTableHeight('team', 'teamCard', 25, 90) - footOff;
        var body = __WEBPACK_IMPORTED_MODULE_2_jquery__('#team').children('tbody');
        var isOverFlow = (body.height() < body.prop('scrollHeight'));
        var refWidth = __WEBPACK_IMPORTED_MODULE_2_jquery__('#team').width() - 22; //-(isOverFlow ? 22 : 0);
        var inpArr = this.oplandb.cloneObject(this.teamCols);
        if (!this.oplandb.logInfo.teamMemAdmin)
            inpArr[inpArr.length - 1] = 0;
        this.teamColsPX = this.udf.getAbsWidthVals(inpArr, refWidth, 0);
    };
    FrmTeamComponent.prototype.onResize = function (e) {
        this.setTableColWidths();
    };
    FrmTeamComponent.prototype.alertModal = function () {
        return false;
    };
    FrmTeamComponent.prototype.getTeamMemberDetails = function (mem, property) {
        var per = this.oplandb.idxPersonnel['r' + mem.perID];
        if (property == 'name') {
            return per.fullName;
        }
        else if (property == 'coy') {
            return this.oplandb.getCompanyName(per.companyID);
        }
        else if (property == 'pos') {
            return this.oplandb.getPositionName(per.positionID);
        }
        else {
            return property;
        }
    };
    FrmTeamComponent.prototype.dispAlert = function (mode) {
        /*if(mode=='delete'){
    
        }*/
        if (this.udf.isTeamDeleteAlert() || this.udf.isTeamMemberDeleteAlert()) {
            return (mode == 'delete');
        }
        return false;
    };
    FrmTeamComponent.prototype.onUserChange = function (e) {
    };
    FrmTeamComponent.prototype.onUserItemSelected = function (e) {
    };
    FrmTeamComponent.prototype.addMemberClick = function () {
        var _this = this;
        if (this.selectedUser == 0) {
            this.oplandb.setTempMessage("msgTempPrompt", "Please select a member to add.");
            return;
        }
        var per = this.oplandb.idxPersonnel['r' + this.selectedUser];
        //[tid] Long, [tmmid] Long, [pid] Long, [shift] LongText, [sbtid] Long;
        var mem = new __WEBPACK_IMPORTED_MODULE_1__models_data_models__["h" /* teamMember */]();
        mem.id = -1;
        mem.perID = this.selectedUser;
        mem.teamID = this.oplandb.currTeamObj.id;
        mem.shift = "";
        mem.subTeamID = 0;
        mem.sort = "0";
        this.oplandb.msgRequestPrompt = "Adding team member '" + per['fullName'] + "'. Please wait...";
        this.subsProcessData = this.oplandb.processData({
            method: "saveTeamMember",
            data: mem,
            isPosting: true,
            processMessage: this.oplandb.msgRequestPrompt,
            subscription: this.subsProcessData,
            onSuccess: function (data) {
                _this.oplandb.setTempMessage("msgTempInfoPrompt", "Team member '" + per['fullName'] + "' has been added to the team.");
                var newMem = new __WEBPACK_IMPORTED_MODULE_1__models_data_models__["h" /* teamMember */]();
                newMem.perID = per.id;
                newMem.shift = "";
                newMem.subTeamID = 0;
                newMem.teamID = _this.oplandb.currTeam.id;
                newMem.id = data['resInt32'];
                var obj = _this.oplandb.currTeam.members.find(function (item) { return item.perID == _this.selectedUser; });
                _this.oplandb.currTeam.members.push(newMem);
                _this.oplandb.currTeamObj.members = _this.oplandb.cloneObject(_this.oplandb.currTeam.members);
                _this.oplandb.currTeam.members.forEach(function (item) {
                    var per = _this.oplandb.idxPersonnel['r' + item.perID];
                    item.sort = per['fullName'];
                });
                __WEBPACK_IMPORTED_MODULE_3__app_globals__["sortBy"](_this.oplandb.currTeam.members, "sort");
                _this.oplandb.setTeamUserSelection();
            }
        });
        /**
      
         
      
          let newMem:teamMember=new teamMember();
          newMem.perID=per.id;
          newMem.shift="";
          newMem.subTeamID=0;
      
          let obj:any = this.oplandb.currTeam.members.find(item=>item.perID == this.selectedUser);
      
          this.oplandb.currTeam.members.push(newMem);
      
          this.oplandb.currTeam.members.forEach(item=>{
            let per:dataPersonnel = this.oplandb.idxPersonnel['r'+item.perID]
            item.sort=per['fullName'];
          })
          G.sortBy(this.oplandb.currTeam.members,"sort");
      
          this.oplandb.setTeamUserSelection();
      
           */
        this.selectedUser = 0;
    };
    FrmTeamComponent.prototype.btnClick = function (type, item) {
        switch (type) {
            case 'ok':
                this.oplandb.saveActivityMember();
                break;
            case 'cancel':
                this.oplandb.cancelActivityMember();
                break;
            case 'edit':
                this.oplandb.editActivityMember(item);
                break;
            case 'del':
                this.oplandb.alertYesNoMessage['message'] = "You are about to remove '" +
                    this.oplandb.idxPersonnel['r' + item['perID']].fullName + "' as member of the team. Do you want to proceed?";
                this.udf.isTeamMemberDeleteAlert(true);
                this.oplandb.currDelObj = item;
                //this.oplandb.deleteTeamMember(item);
                break;
            case 'SAVE':
                this.oplandb.saveTeam();
                break;
            case 'CANCEL':
                this.oplandb.cancelTeam();
                break;
            case 'ADD':
                this.oplandb.newTeam();
                __WEBPACK_IMPORTED_MODULE_3__app_globals__["setFocus"]('teamName');
                break;
            case 'DEL':
                if (this.oplandb.currTeam.id == -1) {
                    this.oplandb.setTempMessage(undefined, "Please select a team to delete.");
                    return;
                }
                this.oplandb.alertYesNoMessage['message'] = "You are about to delete the current team and its related data. Do you want to proceed?";
                this.udf.isTeamDeleteAlert(true);
                break;
            case 'EDIT':
                this.udf.isTeamEditing(true);
                __WEBPACK_IMPORTED_MODULE_3__app_globals__["setFocus"]('teamName');
                break;
            case 'goDelete':
                // put delete action here then normalize mode
                this.oplandb.deleteTeam(undefined, true);
                break;
            case 'noDelete':
                // cancel delete request
                this.udf.isTeamNormal(true);
                break;
            default:
        }
    };
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["ViewChildren"])('input'),
        __metadata("design:type", Object)
    ], FrmTeamComponent.prototype, "vc", void 0);
    FrmTeamComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-frm-team',
            template: __webpack_require__("../../../../../src/app/components/frm-team/frm-team.component.html"),
            styles: [__webpack_require__("../../../../../src/app/components/frm-team/frm-team.component.css")]
        }),
        __metadata("design:paramtypes", [typeof (_a = (typeof __WEBPACK_IMPORTED_MODULE_4__app_udfs__ !== "undefined" && __WEBPACK_IMPORTED_MODULE_4__app_udfs__).AppUDFService) === "function" && _a || Object, typeof (_b = typeof __WEBPACK_IMPORTED_MODULE_5__app_oplandata__["a" /* AppOplpanDataService */] !== "undefined" && __WEBPACK_IMPORTED_MODULE_5__app_oplandata__["a" /* AppOplpanDataService */]) === "function" && _b || Object])
    ], FrmTeamComponent);
    return FrmTeamComponent;
    var _a, _b;
}());

//# sourceMappingURL=frm-team.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/nav-list/nav-list.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "li div{\r\n    cursor: default;\r\n}\r\n\r\nli.act{\r\n    cursor: pointer;\r\n}", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/nav-list/nav-list.component.html":
/***/ (function(module, exports) {

module.exports = "<div class=\"card\" style=\"height:100%;margin-left:20px;\" >\n  <div class=\"card-header\">\n    <div class=\"row\">\n      <div class=\"col\">\n    <div class=\"input-group\" style=\"display:flex;\">\n      <input (change)=\"onSearchTextChange()\" \n          style=\"flex-grow:1;flex-shrink:1;\" class=\"sform-control px-2\"\n              placeholder=\"Search...\"\n              [(ngModel)]=\"this.udf.searchText\"\n              name=\"searchText\"\n      >\n    <div class=\"input-group-addon\" \n      (click)=\"onSerchClicked(true)\" \n      style=\"cursor:pointer;\"><i class=\"fa fa-times\"></i></div>\n    <div class=\"input-group-addon\" \n        (change)=\"onSerchClicked()\" \n        style=\"cursor:pointer;\"><i class=\"fa fa-search\"></i></div>\n    </div>\n  </div>\n  </div>\n    <div class=\"float-left p-0 m-0\">\n      <div class=\"dropdown\" *ngIf=\"udf.currentModule=='By Name' || udf.currentModule=='By Company' || udf.currentModule=='By Position'\">\n          <button class=\"btn btn-sm dropdown-toggle btn-secondary\" \n            type=\"button\" \n            id=\"dropdownMenuButton\" \n            data-toggle=\"dropdown\" \n            aria-haspopup=\"true\" \n            aria-expanded=\"false\"\n            style=\"margin-top:4px;background:#f7f7f9;border:none;\"\n          >\n          {{ udf.currentModule }}\n          </button>\n          <div class=\"dropdown-menu\" aria-labelledby=\"dropdownMenuButton\">\n              <!--a class=\"dropdown-item\"  (click)=\"onSortClicked($event)\"  href=\"#\">Activities</a>\n              <div class=\"dropdown-divider\"></div>\n              <a class=\"dropdown-item\"  (click)=\"onSortClicked($event)\"  href=\"#\">Teams</a>\n              <div class=\"dropdown-divider\"></div>\n              <a class=\"dropdown-item disabled\" href=\"#\"  style=\"color:#292b2c;\">Personnel</a-->\n            <a class=\"dropdown-item\" *ngIf=\"udf.currentModule != 'By Name'\" (click)=\"onSortClicked($event)\" href=\"#\" style=\"padding-left:40px;\">By Name</a>\n            <a class=\"dropdown-item\" *ngIf=\"udf.currentModule != 'By Company'\"  (click)=\"onSortClicked($event)\" href=\"#\" style=\"padding-left:40px;\">By Company</a>\n            <a class=\"dropdown-item\"  *ngIf=\"udf.currentModule != 'By Position'\" (click)=\"onSortClicked($event)\" href=\"#\" style=\"padding-left:40px;\">By Position</a>\n          </div>\n      </div> \n    </div>\n    <div class=\"float-right\" *ngIf=\"false\">\n      <a class=\"btn btn-sm btn-link fa-style-a px-0 mx-1\" href=\"#\" (click)=\"expandToggle($event)\">\n      <i class=\"fa mx-4\" \n        [class.fa-compress]=\"expAll\"\n        [class.fa-expand]=\"!expAll\"\n        aria-hidden=\"true\"></i>\n      </a>\n    </div>\n  </div>\n\n  <div class=\"card-block px-0 py-0\"  style=\"overflow-y: auto;\">\n    \n    <!--Activities navigation-->\n    <ul class=\"list-group\" *ngIf=\"udf.currentModule=='Activities'\">\n      <li  class=\"list-group-item p-2\" \n      [style.background]=\"udf.cardHeaderBackground\"\n      *ngFor=\"let yr of oplandb.activityGroupings\">\n        <div>\n            <i class=\"fa p-0 pr-1 pl-1\" \n              (click)=\"toggleActivityGroup('YEAR',yr)\"\n              [ngClass]=\"{'fa-angle-right':!yr.exp, 'fa-angle-down':yr.exp}\"\n              style=\"cursor:pointer;\" \n              aria-hidden=\"true\"></i>\n              <i class=\"fa fa-calendar-o pr-2\" aria-hidden=\"true\"></i>{{yr.year}}\n        </div>\n        <ul class=\"list-group\" *ngIf=\"yr.exp\">\n          <li class=\"list-group-item p-1 m-0  border-left-0 border-top-0 border-right-0\" \n            style=\"font-size:0.85em;\"\n            *ngFor=\"let mo of yr.months\">\n            <div>\n                <i class=\"fa pr-2 pl-2 m-0\" \n                  (click)=\"toggleActivityGroup('MONTH',mo)\"\n                  [ngClass]=\"{'fa-angle-right':!mo.exp, 'fa-angle-down':mo.exp}\"\n                  style=\"cursor:pointer;\" \n                  aria-hidden=\"true\"></i>\n                  <i class=\"fa fa-calendar pr-2\" aria-hidden=\"true\"></i>{{getMonthName(mo.month)}}\n            </div>\n            <ul class=\"list-group ml-3\" *ngIf=\"mo.exp\">\n              <li class=\"list-group-item p-1 border-top-0 border-left-0 border-right-0\" \n                *ngFor=\"let pri of mo.priorities\">\n                  <div>\n                      <i class=\"fa fa-angle-right pr-2 pl-2\" \n                        (click)=\"toggleActivityGroup('PRIORITY',pri)\"\n                        [ngClass]=\"{'fa-angle-right':!pri.exp, 'fa-angle-down':pri.exp}\"\n                        style=\"cursor:pointer;\" \n                        aria-hidden=\"true\"></i>{{pri.name}}\n                  </div>\n                  <ul class=\"list-group\"  *ngIf=\"pri.exp\">\n                    <li class=\"list-group-item m-0 border-0 p-1 pl-4 act\" \n                      (click)=\"oplandb.scatterData(act)\"\n                      *ngFor=\"let act of pri.activities\"\n                      [style.background]=\"act.id==oplandb.currActivity.id ? udf.selectedTextHilight : 'white'\"\n                      [style.color]=\"act.id==oplandb.currActivity.id ? 'white' : 'black'\"\n                      style=\"border-radius:0px;\"\n                      >\n                        <i class=\"fa fa-tasks mr-2\" aria-hidden=\"true\"></i>{{act.name}}\n                    </li>\n                  </ul>\n              </li>\n            </ul>\n          </li>\n        </ul>\n    </li>\n    </ul>\n    <ul class=\"list-group\" *ngIf=\"udf.currentModule!='Activities'\">\n        <li class=\"list-group-item cusg\" \n          *ngFor=\"let group of oplandb.itemsGroupings\"   \n          [style.background]=\"udf.currGroup==group.name ? '#f7f7f9' : 'white'\" >\n          <div class=\"col\" style=\"display:flex;justify-content:space-between;\">\n            <label>{{ group.name ? group.name : '[unassigned]' }}</label>\n            <a (click)=\"groupExpCol($event)\"  *ngIf=\"udf.currGroup!=group.name\" class=\"fa-style-a\" href=\"#\">\n              <i class=\"fa fa-caret-square-o-down\"></i>\n            </a>\n            <a (click)=\"groupExpCol($event)\"  *ngIf=\"udf.currGroup==group.name\" class=\"fa-style-a\" href=\"#\">\n              <i class=\"fa fa-caret-square-o-up\"></i>\n            </a>\n          </div>\n          <!--currPersonnelObj-->\n          <ul class=\"list-group\" *ngIf=\"udf.currGroup==group.name\" >\n            <li class=\"list-group-item cusi\"\n                  *ngFor=\"let item of group.data;let iIdx=index;\" \n                  [ngStyle]=\"oplandb.getSelectedStyle(item)\"\n                  (click)=\"oplandb.scatterData(item)\"\n              >\n              <label style=\"font-weight:normal\" >\n                {{ item[oplandb.filterFieldName()] + (oplandb.hideIds ? '' : '  (' + item.id  +')') }}\n              </label>\n              <span *ngIf=\"item.activityCount\" class=\"badge badge-pill\" [ngClass]=\"{'badge-info':item.maxDays<=21,'badge-danger':item.maxDays>21}\">{{ item.activityCount }}</span>\n            </li>\n          </ul>\n        </li>\n    </ul>\n  </div>\n</div>"

/***/ }),

/***/ "../../../../../src/app/components/nav-list/nav-list.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return NavListComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__app_globals__ = __webpack_require__("../../../../../src/app/app.globals.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2_jquery__ = __webpack_require__("../../../../jquery/dist/jquery.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2_jquery___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_2_jquery__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3__app_udfs__ = __webpack_require__("../../../../../src/app/app.udfs.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_4__app_oplandata__ = __webpack_require__("../../../../../src/app/app.oplandata.ts");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};





var NavListComponent = (function () {
    //  constructor(public appDataService:Global.AppDataService) { }
    function NavListComponent(appDataService, oplandb, udf) {
        this.appDataService = appDataService;
        this.oplandb = oplandb;
        this.udf = udf;
        this.itemClicked = new __WEBPACK_IMPORTED_MODULE_0__angular_core__["EventEmitter"]();
        this.sortSelectionClicked = new __WEBPACK_IMPORTED_MODULE_0__angular_core__["EventEmitter"]();
        this.thisComponent = new __WEBPACK_IMPORTED_MODULE_0__angular_core__["EventEmitter"]();
        this.expAll = false;
        this.navSortBy = "Activities";
        this.monthNames = [
            "January", "February", "March",
            "April", "May", "June",
            "July", "August", "September",
            "October", "November", "December"
        ];
        this.lkpCompanies = [
            { id: 1, name: "Soga" },
            { id: 2, name: "DOF" },
            { id: 3, name: "SPEX" }
        ];
        this.lkpPositions = [
            { id: 1, name: "Electrical" },
            { id: 2, name: "Assessor" },
            { id: 3, name: "Integrity Engineer" },
            { id: 4, name: "Programmer" }
        ];
    }
    NavListComponent.prototype.itemClass = function (item, idx) {
        return {
            "list-group-item": true,
            "cusi": (idx != 3 ? true : false),
            "cusg2": (idx == 3 ? true : false)
        };
    };
    NavListComponent.prototype.itemGroupName = function (value, group) {
        var obj = null;
        if (group == 'position') {
            obj = this.oplandb.datPosition.find(function (item) { return item['id'] == value; });
        }
        else if (group == 'company') {
            obj = this.oplandb.datCompany.find(function (item) { return item['id'] == value; });
        }
        if (obj) {
            return obj['name'];
        }
        else {
            return "unknown";
        }
    };
    NavListComponent.prototype.onItemClicked = function (e, item) {
        // when a person's name was clicked on the navigation panel
        this.itemClicked.emit(item);
    };
    NavListComponent.prototype.onSortClicked = function (e) {
        //this.sortSelectionClicked.emit(this.navSortBy);
        this.udf.searchText = '';
        this.udf.currentModule = e.srcElement.innerText;
        this.oplandb.setPersonnelNavGroup(e.srcElement.innerText);
        //alert('this.udf.currentModule:'+ this.udf.currentModule)
    };
    NavListComponent.prototype.onSerchClicked = function (clear) {
        if (clear == undefined)
            clear = false;
        if (clear)
            this.udf.searchText = "";
        this.oplandb.refreshNavGroup();
        //this.oplandb.setPersonnelNavGroup();
    };
    NavListComponent.prototype.onSearchTextChange = function () {
        this.oplandb.refreshNavGroup();
        //this.oplandb.setPersonnelNavGroup();
    };
    NavListComponent.prototype.ngOnInit = function () {
        var _this = this;
        var filterString = "a";
        this.udf.searchText = "";
        this.oplandb.getCompanies(true, function (data) {
            _this.oplandb.getPositions(true, function (data) {
                _this.oplandb.getPersonnel(false, function (data) {
                    _this.oplandb.setPersonnelNavGroup();
                });
            });
        });
        //this.getPersonnelsRaw();
        /*
            this.items[0]['actCount']=10;
            this.items[0]['nameGroup']='X';
            this.items[0]['lastName']='XYtel';
        
            this.items.sort(function(a,b) {return (a['lastName'] > b['lastName']) ? 1 : ((b['lastName'] > a['lastName']) ? -1 : 0);})
            
            this.itemsByLastName=this.groupBy(this.items,'nameGroup');
          */
    };
    NavListComponent.prototype.groupExpCol = function (e) {
        if (this.udf.modePersonnel != "normal")
            return;
        var iElem = __WEBPACK_IMPORTED_MODULE_2_jquery__(e.srcElement);
        if (iElem.attr('class').indexOf('-down') != -1) {
            var lbl = iElem.parent().parent().find('label').html();
            this.udf.currGroup = lbl;
        }
        else {
            this.udf.currGroup = "-";
        }
    };
    NavListComponent.prototype.expandToggle = function (e) {
        if (this.udf.modePersonnel != "normal")
            return;
        //expAll
        this.expAll = !this.expAll;
    };
    NavListComponent.prototype.ngAfterContentInit = function () {
        //this.thisComponent.emit(this);
    };
    //var item = myArray.find(item => item.id === 2);
    NavListComponent.prototype.isActivityModule = function () {
        //this.udf
        return false;
    };
    NavListComponent.prototype.getMonthName = function (mo) {
        return this.monthNames[mo];
    };
    NavListComponent.prototype.toggleActivityGroup = function (type, grp) {
        var actg = this.oplandb.activityGroupings;
        var parent = grp.parent;
        if (type == 'YEAR') {
            actg.forEach(function (yr) {
                if (yr.exp && yr.year != grp.year)
                    yr.exp = false;
            });
        }
        else if (type == 'MONTH') {
            parent['months'].forEach(function (mo) {
                if (mo.exp && mo.month != grp.month)
                    mo.exp = false;
            });
        }
        else if (type == 'PRIORITY') {
            parent['priorities'].forEach(function (pri) {
                if (pri.exp && pri.priority != grp.priority)
                    pri.exp = false;
            });
        }
        grp.exp = !grp.exp;
    };
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Output"])(),
        __metadata("design:type", Object)
    ], NavListComponent.prototype, "itemClicked", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Output"])(),
        __metadata("design:type", Object)
    ], NavListComponent.prototype, "sortSelectionClicked", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Output"])(),
        __metadata("design:type", Object)
    ], NavListComponent.prototype, "thisComponent", void 0);
    NavListComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-nav-list',
            template: __webpack_require__("../../../../../src/app/components/nav-list/nav-list.component.html"),
            styles: [__webpack_require__("../../../../../src/app/components/nav-list/nav-list.component.css")]
        }),
        __metadata("design:paramtypes", [typeof (_a = (typeof __WEBPACK_IMPORTED_MODULE_1__app_globals__ !== "undefined" && __WEBPACK_IMPORTED_MODULE_1__app_globals__).AppDataService) === "function" && _a || Object, typeof (_b = typeof __WEBPACK_IMPORTED_MODULE_4__app_oplandata__["a" /* AppOplpanDataService */] !== "undefined" && __WEBPACK_IMPORTED_MODULE_4__app_oplandata__["a" /* AppOplpanDataService */]) === "function" && _b || Object, typeof (_c = typeof __WEBPACK_IMPORTED_MODULE_3__app_udfs__["AppUDFService"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_3__app_udfs__["AppUDFService"]) === "function" && _c || Object])
    ], NavListComponent);
    return NavListComponent;
    var _a, _b, _c;
}());

//# sourceMappingURL=nav-list.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/our-table/our-table.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "th{\r\n    text-align:center;\r\n}\r\n\r\n", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/our-table/our-table.component.html":
/***/ (function(module, exports) {

module.exports = "<!--table class=\"table table-sm table-hover table-bordered\">\n  <thead>\n    <tr>\n      <th class=\"text-center\" *ngFor=\"let col of tableDef.columnDefs\"\n        [style.width]=\"col.width ? col.width+'px' : 'auto'\"\n        [style.display]=\"col.visible ? '' : 'none'\"\n        >{{col.headerText}}</th>\n    </tr>\n  </thead>\n  <tbody>\n    <tr *ngFor=\"let row of dataRows\">\n      <td *ngFor=\"let col of tableDef.columnDefs\"\n        [style.display]=\"col.visible ? '' : 'none'\"\n        [style.text-align]=\"col.align\"\n        [style.width]=\"col.width ? col.width+'px' : 'auto'\"\n        [innerHTML]=\"row[col.dataField]\"\n        ></td>\n    </tr>\n  </tbody>\n</table-->\n<div (scroll)=\"onScroll($event)\"\n  style=\"overflow:auto;width:600px;height:340px;position:relative;top:50px;left:50px;\">\n\n<!--frozen columns header-->\n<table *ngIf=\"isWithFrozen\"  [ngClass]=\"getClass()\"\n  [style.height.px]=\"colHeaderHeight\"\n  [style.width.px]=\"rowHeaderWidth\"\n  [style.margin-left.px]=\"scrollLeft\"\n  [style.margin-top.px]=\"scrollTop\"\n    style=\"position:absolute;z-index:3;background:white;\"\n  >\n  <thead>\n    <!--tr><th colspan=\"2\" class=\"text-center\">Frozen</th></tr>\n    <tr><th style=\"width:80px;\" class=\"text-center\">ID</th><th class=\"text-center\" style=\"width:120px;\">Name</th></tr-->\n    <our-table-header [inputTemplate]=\"frozenHeader\"></our-table-header>      \n  </thead>\n</table>\n\n<!--scrolling columns header table table-sm table-hover table-bordered-->\n<table  [ngClass]=\"getClass()\"\n[style.height.px]=\"colHeaderHeight\"\n[style.width.px]=\"800\"\n[style.margin-left.px]=\"rowHeaderWidth\"\n[style.margin-top.px]=\"scrollTop\"\nstyle=\"position:absolute;z-index:2;background:white;\"\n>\n<thead>\n  <!--tr><th colspan=\"2\" class=\"text-center\">Scrolling</th></tr>\n  <tr>\n    <th  style=\"width:250px;\" class=\"text-center\">Description</th>\n    <th class=\"text-center\">URL</th></tr-->\n    <our-table-header [inputTemplate]=\"scrollingHeader\"></our-table-header>\n</thead>\n</table>\n\n<!--frozen columns-->\n<table *ngIf=\"isWithFrozen\" [ngClass]=\"getClass()\"\n  style=\"position:absolute;z-index:2;background:white;\"\n  [style.margin-top.px]=\"colHeaderHeight\"\n  [style.margin-left.px]=\"scrollLeft\"\n  [style.width.px]=\"rowHeaderWidth\">\n  <tbody>\n    <!--tr *ngFor=\"let row of dataRows\">\n        <td style=\"width:80px;\">{{row.id}}</td>\n        <td style=\"width:120px;\">{{row.name}}</td>\n    </tr-->\n  <our-table-body [inputTemplate]=\"frozenDetails\"></our-table-body>\n  </tbody>\n</table>\n\n<!--scrolling columns-->\n<table  [ngClass]=\"getClass()\"\nstyle=\"position:absolute;z-index:1;\"\n[style.width.px]=\"800\"\n[style.margin-top.px]=\"colHeaderHeight\"\n[style.margin-left.px]=\"rowHeaderWidth\">\n<tbody>\n  <!--tr *ngFor=\"let row of dataRows\">\n      <td style=\"width:250px;\">{{row.desc}}</td>\n      <td style=\"width:550px;\"[innerHTML]=\"row.url\"></td>\n  </tr-->\n  <our-table-body [inputTemplate]=\"scrollingDetails\"></our-table-body>\n</tbody>\n</table>\n</div>"

/***/ }),

/***/ "../../../../../src/app/components/our-table/our-table.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "c", function() { return OurTableComponent; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "d", function() { return OurTableHeaderComponent; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return OurTableBodyComponent; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "b", function() { return OurTableColumnDef; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1_jquery__ = __webpack_require__("../../../../jquery/dist/jquery.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1_jquery___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_1_jquery__);
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};


var OurTableComponent = (function () {
    function OurTableComponent() {
        this.isWithFrozen = false;
        this.dataRows = [
            { id: 1, name: "Item 1", desc: "Item 1 description", url: '<a href="https://wwww.yahoo.com">Yahoo!!!</a>' },
            { id: 2, name: "Item 2", desc: "Item 2 description", url: "https://wwww.google.com" },
            { id: 3, name: "Item 3", desc: "Item 3 description", url: "https://wwww.bpi.com" },
            { id: 4, name: "Item 4", desc: "Item 4 description", url: "https://wwww.i-rbi.com" },
            { id: 5, name: "Item 5", desc: "Item 5 description", url: "https://wwww.soga.com" },
            { id: 6, name: "Item 1", desc: "Item 1 description", url: '<a href="https://wwww.yahoo.com">Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliq</a>' },
            { id: 7, name: "Item 2", desc: "Item 2 description", url: "https://wwww.google.com" },
            { id: 8, name: "Item 3", desc: "Item 3 description", url: "https://wwww.bpi.com" },
            { id: 9, name: "Item 4", desc: "Item 4 description", url: "https://wwww.i-rbi.com" },
            { id: 10, name: "Item 5", desc: "Item 5 description", url: "https://wwww.soga.com" },
        ];
        this.columnDefs = [
            {
                name: "col1",
                headerText: "ID",
                freeze: true,
                type: "string",
                dataField: "id",
                width: 120,
                align: "center",
                visible: false
            },
            {
                name: "col2",
                headerText: "Name",
                freeze: false,
                type: "string",
                dataField: "name",
                sortable: true,
                sort: "asc",
                order: 1,
                width: 120,
                visible: true
            },
            {
                name: "col3",
                headerText: "Description",
                freeze: false,
                type: "string",
                dataField: "desc",
                sortable: true,
                sort: "asc",
                order: 2,
                visible: true
            },
            {
                name: "col4",
                headerText: "Yes\/No",
                freeze: false,
                type: "string",
                dataField: "url",
                sortable: true,
                sort: "asc",
                order: 3,
                width: 200,
                visible: true,
                html: true
            }
        ];
        this.tableDef = {
            columnDefs: this.columnDefs
        };
        //rowHeaderWidth:number=this.isWithFrozen ? 200 : 0;
        this.rowHeaderWidth = 0;
        this.colHeaderHeight = 68;
        this.scrollLeft = 0;
        this.scrollTop = 0;
    }
    OurTableComponent.prototype.ngOnInit = function () {
        var _this = this;
        this.isWithFrozen = (this.frozenHeader != null && this.frozenDetails != null);
        this.rowHeaderWidth = this.isWithFrozen ? 200 : 0;
        setTimeout(function () { _this.setRowHeight(); }, 0);
    };
    OurTableComponent.prototype.onScroll = function (e) {
        var dv = __WEBPACK_IMPORTED_MODULE_1_jquery__(e.srcElement);
        this.scrollLeft = dv.scrollLeft();
        this.scrollTop = dv.scrollTop();
    };
    OurTableComponent.prototype.setRowHeight = function () {
        var tbls = __WEBPACK_IMPORTED_MODULE_1_jquery__('table');
        var LL_TR = __WEBPACK_IMPORTED_MODULE_1_jquery__(tbls[2]).find('tr');
        var LR_TR = __WEBPACK_IMPORTED_MODULE_1_jquery__(tbls[3]).find('tr');
        for (var i = 0; i < LL_TR.length; i++) {
            var ltr = __WEBPACK_IMPORTED_MODULE_1_jquery__(LL_TR[i]);
            var rtr = __WEBPACK_IMPORTED_MODULE_1_jquery__(LR_TR[i]);
            var ltrh = ltr.height();
            var rtrh = rtr.height();
            if (ltrh < rtrh) {
                ltr.css('height', rtrh + 'px');
            }
            else if (ltrh > rtrh) {
                rtr.css('height', ltrh + 'px');
            }
        }
    };
    OurTableComponent.prototype.getClass = function () {
        return {
            "table": true,
            "table-sm": true,
            "table-hover": true,
            "table-bordered": true,
        };
    };
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", typeof (_a = typeof __WEBPACK_IMPORTED_MODULE_0__angular_core__["TemplateRef"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_0__angular_core__["TemplateRef"]) === "function" && _a || Object)
    ], OurTableComponent.prototype, "scrollingHeader", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", typeof (_b = typeof __WEBPACK_IMPORTED_MODULE_0__angular_core__["TemplateRef"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_0__angular_core__["TemplateRef"]) === "function" && _b || Object)
    ], OurTableComponent.prototype, "frozenHeader", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", typeof (_c = typeof __WEBPACK_IMPORTED_MODULE_0__angular_core__["TemplateRef"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_0__angular_core__["TemplateRef"]) === "function" && _c || Object)
    ], OurTableComponent.prototype, "scrollingDetails", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", typeof (_d = typeof __WEBPACK_IMPORTED_MODULE_0__angular_core__["TemplateRef"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_0__angular_core__["TemplateRef"]) === "function" && _d || Object)
    ], OurTableComponent.prototype, "frozenDetails", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Object)
    ], OurTableComponent.prototype, "tableDef", void 0);
    OurTableComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'our-table',
            template: __webpack_require__("../../../../../src/app/components/our-table/our-table.component.html"),
            styles: [__webpack_require__("../../../../../src/app/components/our-table/our-table.component.css")]
        }),
        __metadata("design:paramtypes", [])
    ], OurTableComponent);
    return OurTableComponent;
    var _a, _b, _c, _d;
}());

var OurTableHeaderComponent = (function () {
    function OurTableHeaderComponent() {
    }
    OurTableHeaderComponent.prototype.ngOnInit = function () {
    };
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", typeof (_a = typeof __WEBPACK_IMPORTED_MODULE_0__angular_core__["TemplateRef"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_0__angular_core__["TemplateRef"]) === "function" && _a || Object)
    ], OurTableHeaderComponent.prototype, "inputTemplate", void 0);
    OurTableHeaderComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'our-table-header',
            template: "\n  <ng-container *ngTemplateOutlet=\"inputTemplate\"></ng-container>\n  ",
            styles: [__webpack_require__("../../../../../src/app/components/our-table/our-table.component.css")]
        }),
        __metadata("design:paramtypes", [])
    ], OurTableHeaderComponent);
    return OurTableHeaderComponent;
    var _a;
}());

var OurTableBodyComponent = (function () {
    function OurTableBodyComponent() {
    }
    OurTableBodyComponent.prototype.ngOnInit = function () {
    };
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", typeof (_a = typeof __WEBPACK_IMPORTED_MODULE_0__angular_core__["TemplateRef"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_0__angular_core__["TemplateRef"]) === "function" && _a || Object)
    ], OurTableBodyComponent.prototype, "inputTemplate", void 0);
    OurTableBodyComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'our-table-body',
            template: "\n  <ng-container *ngTemplateOutlet=\"inputTemplate\"></ng-container>\n  ",
            styles: [__webpack_require__("../../../../../src/app/components/our-table/our-table.component.css")]
        }),
        __metadata("design:paramtypes", [])
    ], OurTableBodyComponent);
    return OurTableBodyComponent;
    var _a;
}());

var OurTableColumnDef = (function () {
    function OurTableColumnDef() {
        this.name = "";
        this.freeze = false;
        this.type = "string";
        this.readOnly = true;
        this.width = 120;
        this.sort = "";
        this.sortable = false;
        this.hasFilter = true;
        this.visible = true;
        this.order = 0;
        this.align = "";
        this.html = false;
    }
    return OurTableColumnDef;
}());

//# sourceMappingURL=our-table.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/pnl-vert-split/pnl-vert-split.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/pnl-vert-split/pnl-vert-split.component.html":
/***/ (function(module, exports) {

module.exports = "<div class=\" px-0 py-0 mx-0 my-0\" style=\"position:relative;display:flex;flex-grow:1;flex-shrink:1;flex-basis:0;background:transparent;\">\n  <div class=\" px-0 py-0 mx-0 my-0\" \n    [style.width]=\"splitHandleLeft\"\n    style=\"background:transparent;color:black;\n      height:100%;\">\n      <div #leftContainer>\n\n      </div>\n  </div>\n  <div (mousedown)=\"onMouseDown($event)\" (mouseup)=\"onMouseUp($event)\" \n    class=\" px-0 py-0 mx-0 my-0\" \n    [style.width]=\"splitHandleWidth\"\n    style=\"background:gray;cursor:e-resize;height:100%;font-size:0px;\">&nbsp;\n  </div>\n  <div #rightContainer class=\" px-0 py-0 mx-0 my-0\" style=\"background:yellow;color:black;flex-grow:1;flex-shrink:1;height:100%;\">{{rightComponent}}</div>\n  <div *ngIf=\"displayMask\" (mouseup)=\"onMaskMouseUp($event)\" (mousemove)=\"onMaskMouseMove($event)\"\n    style= \"position:absolute;top:0px;left:0px;\n            width:100%;height:100%;\n            background:black;\n            opacity:0;\n            z-index:10;\n            cursor:e-resize;\">\n  </div>\n  <div *ngIf=\"displayMask\"\n    [style.width]=\"splitHandleWidth\"\n    [style.left]=\"splitHandleLeft\"\n    style= \"position:absolute;top:0px;\n            height:100%;\n            background:gray;\n            z-index:5;\n            cursor:e-resize;\">&nbsp;\n  </div>          \n</div>\n\n"

/***/ }),

/***/ "../../../../../src/app/components/pnl-vert-split/pnl-vert-split.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return PnlVertSplitComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};

var PnlVertSplitComponent = (function () {
    function PnlVertSplitComponent(resolver, renderer, elementRef) {
        this.resolver = resolver;
        this.leftStatus = 'none';
        this.rightStatus = 'none';
        this.maskDisplayStyle = 'inline';
        this.displayMask = false;
        this.splitHandleWidth = "5px";
        this.splitHandleLeft = "500px";
        this.leftComponentLoaded = new __WEBPACK_IMPORTED_MODULE_0__angular_core__["EventEmitter"]();
    }
    PnlVertSplitComponent.prototype.ngOnInit = function () {
    };
    PnlVertSplitComponent.prototype.ngAfterContentInit = function () {
        var leftContainerFactory = this.resolver.resolveComponentFactory(this.leftComponent.component);
        var data = this.leftComponent.data;
        var comp = this.leftContainer.createComponent(leftContainerFactory);
        var k;
        for (k in data) {
            comp.instance[k] = data[k];
        }
        this.leftComponentLoaded.emit(comp.instance);
    };
    PnlVertSplitComponent.prototype.onMouseDown = function (e) {
        this.leftStatus = 'left mouse button pressed';
        this.displayMask = true;
    };
    PnlVertSplitComponent.prototype.onMouseUp = function (e) {
        this.leftStatus = 'left mouse button released';
    };
    PnlVertSplitComponent.prototype.onMaskMouseUp = function (e) {
        this.displayMask = false;
    };
    PnlVertSplitComponent.prototype.onMaskMouseMove = function (e) {
        this.splitHandleLeft = (e.offsetX - parseFloat(this.splitHandleWidth) / 2) + 'px';
    };
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Object)
    ], PnlVertSplitComponent.prototype, "leftComponent", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Object)
    ], PnlVertSplitComponent.prototype, "rightComponent", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Output"])(),
        __metadata("design:type", Object)
    ], PnlVertSplitComponent.prototype, "leftComponentLoaded", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["ViewChild"])('leftContainer', { read: __WEBPACK_IMPORTED_MODULE_0__angular_core__["ViewContainerRef"] }),
        __metadata("design:type", Object)
    ], PnlVertSplitComponent.prototype, "leftContainer", void 0);
    PnlVertSplitComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-pnl-vert-split',
            template: __webpack_require__("../../../../../src/app/components/pnl-vert-split/pnl-vert-split.component.html"),
            styles: [__webpack_require__("../../../../../src/app/components/pnl-vert-split/pnl-vert-split.component.css")]
        }),
        __metadata("design:paramtypes", [typeof (_a = typeof __WEBPACK_IMPORTED_MODULE_0__angular_core__["ComponentFactoryResolver"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_0__angular_core__["ComponentFactoryResolver"]) === "function" && _a || Object, typeof (_b = typeof __WEBPACK_IMPORTED_MODULE_0__angular_core__["Renderer"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_0__angular_core__["Renderer"]) === "function" && _b || Object, typeof (_c = typeof __WEBPACK_IMPORTED_MODULE_0__angular_core__["ElementRef"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_0__angular_core__["ElementRef"]) === "function" && _c || Object])
    ], PnlVertSplitComponent);
    return PnlVertSplitComponent;
    var _a, _b, _c;
}());

//# sourceMappingURL=pnl-vert-split.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/pobprofile/pobprofile.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/pobprofile/pobprofile.component.html":
/***/ (function(module, exports) {

module.exports = "<button style=\"position:absolute;top:0px;left:0px;z-index:50;left:470px;cursor:pointer;padding:0px;\"\nclass=\"btn btn-link\"\ndata-toggle=\"tooltip\"\n[style.top]=\"'7px'\"\n(click)=\"reloadPOBProfile()\"\ntitle=\"Reload POB Profile new specified scope dates ...\"\n[disabled]=\"disableReload()\"\n><i class=\"fa fa-refresh\" aria-hidden=\"true\"\nstyle=\"font-size:20px;\"\n></i></button>\n\n<!--button style=\"position:absolute;top:0px;left:0px;z-index:50;left:500px;cursor:pointer;padding:0px;\"\nclass=\"btn btn-link\"\ndata-toggle=\"tooltip\"\n[style.top]=\"'9px'\"\n(click)=\"downloadProfile()\"\ntitle=\"Save POB Profile as image ...\"\n><i class=\"fa fa-download\" aria-hidden=\"true\"\nstyle=\"font-size:20px;\"\n></i></button-->\n\n<label *ngIf=\"statusText.length!=0\" class=\"form-label text-danger\"\nstyle=\"position:absolute;z-index:50;left:500px;font-size:0.9em;width:400px;font-weight:bold;\"\n[style.top]=\"'7px'\"\n>\n{{statusText}}\n</label>\n<div class=\"card\" style=\"position:absolute;left:15px;\"\n    [style.height]=\"(oplandb.winHeight - udf.bannerHeight)+'px'\"\n    [style.top]=\"udf.bannerMenuHeight + 'px'\"\n>\n    <div class=\"card-header py-0\">\n        <h5 class=\"my-1\">{{ getCardTitle() }}</h5>\n    </div>\n    <div class=\"card-block\"\n        [style.width]=\"(oplandb.winWidth)+'px'\"\n        [style.height]=\"(oplandb.winHeight-udf.bannerHeight)+'px'\"\n        [style.top]=\"udf.bannerMenuHeight + 'px'\"\n        style=\"background: white;overflow-y:auto;\">\n\n        <canvas id=\"canvas\" #cnv>{{ myChart }}</canvas>\n    </div>\n\n</div>\n"

/***/ }),

/***/ "../../../../../src/app/components/pobprofile/pobprofile.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return PobprofileComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1_chart_js__ = __webpack_require__("../../../../chart.js/src/chart.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1_chart_js___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_1_chart_js__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__app_udfs__ = __webpack_require__("../../../../../src/app/app.udfs.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3__app_oplandata__ = __webpack_require__("../../../../../src/app/app.oplandata.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_4_jquery__ = __webpack_require__("../../../../jquery/dist/jquery.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_4_jquery___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_4_jquery__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_5__app_globals__ = __webpack_require__("../../../../../src/app/app.globals.ts");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};






var PobprofileComponent = (function () {
    function PobprofileComponent(udf, oplandb) {
        this.udf = udf;
        this.oplandb = oplandb;
        this.chartColors = ["#3c6494", "#963d3b", "#799244", "#634d7e", "#39869b", "#c27535", "#4978b1", "#46a1b9", "#e78c41", "#7e9bc8", "#ca7e7d", "#aec683", "#9b89b3", "#7cbbcf", "#f8aa79", "#ff0000", "#c00000", "#745995"];
        this.myChart = null;
        this.chart = []; // This will hold our chart info
        this.chartData = { datasets: [] };
        this.profileData = { "message": "", "cod": "200", "city_id": 2643743, "calctime": 0.0875, "cnt": 3, "list": [{ "main": { "temp": 279.946, "temp_min": 100.946, "temp_max": 279.946, "pressure": 1016.76, "sea_level": 1024.45, "grnd_level": 1016.76, "humidity": 100 }, "wind": { "speed": 4.59, "deg": 163.001 }, "clouds": { "all": 92 }, "weather": [{ "id": 500, "main": "Rain", "description": "light rain", "icon": "10n" }], "rain": { "3h": 2.69 }, "dt": 1485717216 }, { "main": { "temp": 282.597, "temp_min": 282.597, "temp_max": 282.597, "pressure": 1012.12, "sea_level": 1019.71, "grnd_level": 1012.12, "humidity": 98 }, "wind": { "speed": 4.04, "deg": 226 }, "clouds": { "all": 92 }, "weather": [{ "id": 500, "main": "Rain", "description": "light rain", "icon": "10n" }], "rain": { "3h": 0.405 }, "dt": 1485745061 }, { "main": { "temp": 279.38, "pressure": 1011, "humidity": 93, "temp_min": 278.15, "temp_max": 280.15 }, "wind": { "speed": 2.6, "deg": 30 }, "clouds": { "all": 90 }, "weather": [{ "id": 701, "main": "Mist", "description": "mist", "icon": "50d" }, { "id": 741, "main": "Fog", "description": "fog", "icon": "50d" }], "dt": 1485768552 }] };
        this.statusText = "";
        this.link = null;
    }
    PobprofileComponent.prototype.ngOnInit = function () {
        this.udf.resetPOBProfileScope();
        this.reloadPOBProfile();
    };
    PobprofileComponent.prototype.ngOnDestroy = function () {
        this.oplandb.clearSubscription(this.subsReload);
    };
    PobprofileComponent.prototype.reloadPOBProfile = function () {
        var _this = this;
        this.udf.profStartDate = __WEBPACK_IMPORTED_MODULE_5__app_globals__["longToDate"](this.udf.profDispStartDate);
        this.udf.profEndDate = __WEBPACK_IMPORTED_MODULE_5__app_globals__["longToDate"](this.udf.profDispEndDate);
        this.oplandb.processData({
            method: "getPOBProfile",
            processMessage: "Extracting POB Profile data. Please wait...",
            subscription: this.subsReload,
            data: {
                startDate: __WEBPACK_IMPORTED_MODULE_5__app_globals__["dateToString"](this.udf.profStartDate),
                endDate: __WEBPACK_IMPORTED_MODULE_5__app_globals__["dateToString"](this.udf.profEndDate)
            },
            onSuccess: function (data) {
                _this.chartData = { datasets: [] };
                var teams = data['teams'];
                _this.chartData['labels'] = [];
                if (data['d0'] != undefined) {
                    var stdLimit = [];
                    var upLimit = [];
                    var sd = _this.udf.profStartDate;
                    for (var idx = 1; idx <= data['d0'].length; idx++) {
                        var dt = new Date(sd.getFullYear(), sd.getMonth(), sd.getDate() + idx - 1);
                        var lbl = dt.toLocaleDateString('en-GB', { year: '2-digit', month: 'short', day: 'numeric' });
                        _this.chartData['labels'].push(lbl.replace(/ /gi, '-'));
                        var std = data['std'][idx - 1];
                        var upm = data['upm'][idx - 1];
                        stdLimit.push(std); // standard
                        upLimit.push(Math.max(upm - std, 0)); // upmanning less standard
                    }
                    for (var idx = 0; idx < teams.length; idx++) {
                        // split team info to separate team name from the switches that come with it
                        // (ie. tean field indicator, sort field, ..., etc)
                        var tmArr = teams[idx].split('^');
                        _this.chartData['datasets'].push({
                            label: tmArr[tmArr.length - 1],
                            data: data["d" + idx],
                            backgroundColor: _this.chartColors[idx]
                        });
                    } // for ...
                    _this.chartData['datasets'].push({
                        type: 'line',
                        label: "STD POB Limit",
                        //display:'linear',
                        data: stdLimit,
                        fill: false,
                        lineTension: 0,
                        borderColor: "rgba(255,0,0,0.7)",
                        backgroundColor: 'white',
                        borderWidth: '3',
                        pointRadius: 0
                    });
                    _this.chartData['datasets'].push({
                        type: 'line',
                        label: "UpManning Limit",
                        data: upLimit,
                        fill: false,
                        lineTension: 0,
                        borderColor: "rgba(255,0,0,0.9)",
                        backgroundColor: 'white',
                        borderWidth: '3',
                        pointRadius: 0
                    });
                    setTimeout(function () { _this.generateProfile(); }, 100);
                }
                else {
                    // no data found
                    _this.statusText = "Sorry no data found within the selected date scope !";
                }
                /**
                 *
                 *
                 this.config['stdPOBLimit']['limits']=[];
            this.config['upmPOBLimit']['limits']=[];
            this.config['mobLimit']['limits']=[];
            this.config['demobLimit']['limits']=[];
            this.config['upmanPeriods']['periods']=[];
        
                 console.log("Configuration data...")
        
                 console.log(this.oplandb.config)
                 setTimeout(()=>{
                  console.log(this.oplandb.config.stdPOBLimit.limits)
                  console.log(this.oplandb.config.upmPOBLimit.limits)
                  console.log(this.oplandb.config.upmanPeriods.periods)
                  console.log(this.oplandb.config.upmanPeriods.periods)
        
                 },2000)
                 */
            } // on success 2018-09-27
        }); // process data
    };
    PobprofileComponent.prototype.downloadProfile = function () {
        if (!this.cnv) {
            this.statusText = "Cannot save empty chart !";
            return;
        }
        var canvas = this.cnv.nativeElement;
        if (this.link == null) {
            this.link = document.createElement('a');
            this.link.id = "downloadProfile";
            this.link.innerHTML = 'download image';
            this.link.addEventListener('click', function (ev) {
                this.link.href = canvas.toDataURL();
                this.link.download = "mypainting.png";
            }, false);
            document.body.appendChild(this.link);
        }
        __WEBPACK_IMPORTED_MODULE_4_jquery__("#downloadProfile").trigger("click");
    };
    PobprofileComponent.prototype.generateProfile = function () {
        if (this.myChart != null) {
            this.myChart.destroy();
        }
        this.myChart = new __WEBPACK_IMPORTED_MODULE_1_chart_js__["Chart"]('canvas', {
            type: 'bar',
            title: {
                text: "SWP POB Profile"
            },
            data: this.chartData,
            options: {
                title: {
                    display: true,
                    text: this.getCardTitle(),
                    fontSize: 18,
                    position: 'bottom'
                },
                legend: {
                    display: true,
                    position: 'bottom'
                },
                scales: {
                    xAxes: [{ stacked: true }],
                    yAxes: [{ stacked: true }],
                    maintainAspectRatio: false,
                    responsive: true
                }
            }
        });
    };
    PobprofileComponent.prototype.getCardTitle = function () {
        var sd = this.udf.profStartDate;
        var ed = this.udf.profEndDate;
        return "SWP POB Profile (" + this.udf.profDispStartDate + " to " + this.udf.profDispEndDate + ")";
    };
    PobprofileComponent.prototype.disableReload = function () {
        // check if no change on the current filter scope
        var sd = __WEBPACK_IMPORTED_MODULE_5__app_globals__["longToDate"](this.udf.profDispStartDate);
        var ed = __WEBPACK_IMPORTED_MODULE_5__app_globals__["longToDate"](this.udf.profDispEndDate);
        var retVal = ((this.udf.profDispStartDate == __WEBPACK_IMPORTED_MODULE_5__app_globals__["dateToString"](this.udf.profStartDate, true)) &&
            (this.udf.profDispEndDate == __WEBPACK_IMPORTED_MODULE_5__app_globals__["dateToString"](this.udf.profEndDate, true)));
        this.statusText = '';
        if (!retVal) {
            // check if invalid date
            retVal = (ed < sd);
            if (retVal)
                this.statusText = '( Start date cannot be later than end date )';
        }
        else {
            // date was changed
            this.statusText = '';
        }
        if (!retVal) {
            // check maximum dates
            retVal = (__WEBPACK_IMPORTED_MODULE_5__app_globals__["daysDiff"](sd, ed) > this.udf.profScopeLimit);
            if (retVal)
                this.statusText = '( Scope cannot be more than ' + this.udf.profScopeLimit + ' days )';
        }
        //alert(ed-sd);
        //this.statusText=String(G.daysDiff(sd,ed) + ' , ' + this.udf.calScopeLimit);
        return retVal;
    };
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["ViewChild"])('cnv'),
        __metadata("design:type", typeof (_a = typeof __WEBPACK_IMPORTED_MODULE_0__angular_core__["ElementRef"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_0__angular_core__["ElementRef"]) === "function" && _a || Object)
    ], PobprofileComponent.prototype, "cnv", void 0);
    PobprofileComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-pobprofile',
            template: __webpack_require__("../../../../../src/app/components/pobprofile/pobprofile.component.html"),
            styles: [__webpack_require__("../../../../../src/app/components/pobprofile/pobprofile.component.css")]
        }),
        __metadata("design:paramtypes", [typeof (_b = typeof __WEBPACK_IMPORTED_MODULE_2__app_udfs__["AppUDFService"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_2__app_udfs__["AppUDFService"]) === "function" && _b || Object, typeof (_c = typeof __WEBPACK_IMPORTED_MODULE_3__app_oplandata__["a" /* AppOplpanDataService */] !== "undefined" && __WEBPACK_IMPORTED_MODULE_3__app_oplandata__["a" /* AppOplpanDataService */]) === "function" && _c || Object])
    ], PobprofileComponent);
    return PobprofileComponent;
    var _a, _b, _c;
}());

//# sourceMappingURL=pobprofile.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/report-header/report-header.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/report-header/report-header.component.html":
/***/ (function(module, exports) {

module.exports = "<div class=\"d-flex\">\n  <img src=\"assets/images/shell_logo_d.png\" style=\"margin:5px;\" alt=\"image\">\n  <label class=\"h2 p-0 pt-3 pl-3\" \n    style=\"text-shadow: 1px 2px #c0c0c0;color:#296192;font-weight:bold;\">\n      {{headerText}}</label>\n</div>"

/***/ }),

/***/ "../../../../../src/app/components/report-header/report-header.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return ReportHeaderComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};

var ReportHeaderComponent = (function () {
    function ReportHeaderComponent() {
    }
    ReportHeaderComponent.prototype.ngOnInit = function () {
    };
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", String)
    ], ReportHeaderComponent.prototype, "headerText", void 0);
    ReportHeaderComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-report-header',
            template: __webpack_require__("../../../../../src/app/components/report-header/report-header.component.html"),
            styles: [__webpack_require__("../../../../../src/app/components/report-header/report-header.component.css")]
        }),
        __metadata("design:paramtypes", [])
    ], ReportHeaderComponent);
    return ReportHeaderComponent;
}());

//# sourceMappingURL=report-header.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/temp-component/temp-component.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/temp-component/temp-component.component.html":
/***/ (function(module, exports) {

module.exports = "<p>\n  temp-component works!\n</p>\n"

/***/ }),

/***/ "../../../../../src/app/components/temp-component/temp-component.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return TempComponentComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};

var TempComponentComponent = (function () {
    function TempComponentComponent() {
    }
    TempComponentComponent.prototype.ngOnInit = function () {
    };
    TempComponentComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-temp-component',
            template: __webpack_require__("../../../../../src/app/components/temp-component/temp-component.component.html"),
            styles: [__webpack_require__("../../../../../src/app/components/temp-component/temp-component.component.css")]
        }),
        __metadata("design:paramtypes", [])
    ], TempComponentComponent);
    return TempComponentComponent;
}());

//# sourceMappingURL=temp-component.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/test-child/test-child.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/test-child/test-child.component.html":
/***/ (function(module, exports) {

module.exports = "<ng-container *ngTemplateOutlet=\"template\">\n    \n</ng-container>"

/***/ }),

/***/ "../../../../../src/app/components/test-child/test-child.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return TestChildComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};

var TestChildComponent = (function () {
    function TestChildComponent() {
        this.prompt = "Sample text...";
        this.disp = "";
        this.paneUL = false;
        this.paneLL = false;
        this.paneUR = false;
        this.paneLR = false;
    }
    TestChildComponent.prototype.ngOnInit = function () {
    };
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", typeof (_a = typeof __WEBPACK_IMPORTED_MODULE_0__angular_core__["TemplateRef"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_0__angular_core__["TemplateRef"]) === "function" && _a || Object)
    ], TestChildComponent.prototype, "template", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", String)
    ], TestChildComponent.prototype, "prompt", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", String)
    ], TestChildComponent.prototype, "disp", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Boolean)
    ], TestChildComponent.prototype, "paneUL", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Boolean)
    ], TestChildComponent.prototype, "paneLL", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Boolean)
    ], TestChildComponent.prototype, "paneUR", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Boolean)
    ], TestChildComponent.prototype, "paneLR", void 0);
    TestChildComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-test-child',
            template: __webpack_require__("../../../../../src/app/components/test-child/test-child.component.html"),
            styles: [__webpack_require__("../../../../../src/app/components/test-child/test-child.component.css")]
        }),
        __metadata("design:paramtypes", [])
    ], TestChildComponent);
    return TestChildComponent;
    var _a;
}());

//# sourceMappingURL=test-child.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/test-counter/test-counter.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return TestCounterComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};

var TestCounterComponent = (function () {
    function TestCounterComponent() {
        this.count = 0;
        this.counterChange = new __WEBPACK_IMPORTED_MODULE_0__angular_core__["EventEmitter"]();
    }
    Object.defineProperty(TestCounterComponent.prototype, "counter", {
        get: function () {
            return this.count;
        },
        enumerable: true,
        configurable: true
    });
    TestCounterComponent.prototype.increment = function () {
        this.count = this.count + 1;
        this.counterChange.emit(this.count);
    };
    TestCounterComponent.prototype.decrement = function () {
        this.count = this.count - 1;
        this.counterChange.emit(this.count);
    };
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Output"])(),
        __metadata("design:type", typeof (_a = typeof __WEBPACK_IMPORTED_MODULE_0__angular_core__["EventEmitter"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_0__angular_core__["EventEmitter"]) === "function" && _a || Object)
    ], TestCounterComponent.prototype, "counterChange", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", Object),
        __metadata("design:paramtypes", [])
    ], TestCounterComponent.prototype, "counter", null);
    TestCounterComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'countercomponent',
            template: "\n    <button (click)='increment()'>Increment</button>\n    {{count}}\n    <button (click)='decrement()'>Decrement</button>\n    "
        }),
        __metadata("design:paramtypes", [])
    ], TestCounterComponent);
    return TestCounterComponent;
    var _a;
}());

//# sourceMappingURL=test-counter.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/test-master/test-master.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/test-master/test-master.component.html":
/***/ (function(module, exports) {

module.exports = "\n\n<app-test-child #childComponent1 [paneUL]=\"true\" [template]=\"childTemplate\" [prompt]=\"'Hello world'\"></app-test-child>\n<app-test-child #childComponent2 [paneLL]=\"true\" [template]=\"childTemplate\" [prompt]=\"'The quick brown fox jumps over the lazy dog'\"></app-test-child>\n\n<!--p>\n  test-master works!\n</p>\n<app-test-child #childComponent1 [template]=\"childTemplate\" [prompt]=\"'Hello world'\"></app-test-child>\n<app-test-child #childComponent2  [prompt]=\"'The quick brown fox jumps over the lazy dog'\"></app-test-child-->"

/***/ }),

/***/ "../../../../../src/app/components/test-master/test-master.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return TestMasterComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__test_child_test_child_component__ = __webpack_require__("../../../../../src/app/components/test-child/test-child.component.ts");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};


var TestMasterComponent = (function () {
    function TestMasterComponent() {
    }
    TestMasterComponent.prototype.ngOnInit = function () {
    };
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["ViewChild"])('childComponent1'),
        __metadata("design:type", typeof (_a = typeof __WEBPACK_IMPORTED_MODULE_1__test_child_test_child_component__["a" /* TestChildComponent */] !== "undefined" && __WEBPACK_IMPORTED_MODULE_1__test_child_test_child_component__["a" /* TestChildComponent */]) === "function" && _a || Object)
    ], TestMasterComponent.prototype, "tableUL", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["ViewChild"])('childComponent2'),
        __metadata("design:type", typeof (_b = typeof __WEBPACK_IMPORTED_MODULE_1__test_child_test_child_component__["a" /* TestChildComponent */] !== "undefined" && __WEBPACK_IMPORTED_MODULE_1__test_child_test_child_component__["a" /* TestChildComponent */]) === "function" && _b || Object)
    ], TestMasterComponent.prototype, "tableLL", void 0);
    __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Input"])(),
        __metadata("design:type", typeof (_c = typeof __WEBPACK_IMPORTED_MODULE_0__angular_core__["TemplateRef"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_0__angular_core__["TemplateRef"]) === "function" && _c || Object)
    ], TestMasterComponent.prototype, "childTemplate", void 0);
    TestMasterComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-test-master',
            template: __webpack_require__("../../../../../src/app/components/test-master/test-master.component.html"),
            styles: [__webpack_require__("../../../../../src/app/components/test-master/test-master.component.css")]
        }),
        __metadata("design:paramtypes", [])
    ], TestMasterComponent);
    return TestMasterComponent;
    var _a, _b, _c;
}());

//# sourceMappingURL=test-master.component.js.map

/***/ }),

/***/ "../../../../../src/app/components/test-panels/test-panels.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/components/test-panels/test-panels.component.html":
/***/ (function(module, exports) {

module.exports = "    <div class=\"header\">Header ({{selectedDate}})</div>\n    <div>\n        <input class=\"form-control form-control-sm\" \n          style=\"text-align:center;width:120px;\"\n          [(ngModel)]=\"selectedDate\" \n          [dpDayPicker]=\"datePickerConfig\" />\n    </div>\n"

/***/ }),

/***/ "../../../../../src/app/components/test-panels/test-panels.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return TestPanelsComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1_jquery__ = __webpack_require__("../../../../jquery/dist/jquery.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1_jquery___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_1_jquery__);
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};


var TestPanelsComponent = (function () {
    function TestPanelsComponent() {
        this.selectedDate = "20 Oct 2017";
        this.datePickerConfig = {
            format: "D MMM YYYY",
            locale: "en"
        };
    }
    TestPanelsComponent.prototype.ngOnInit = function () {
        //alert($('.dp-picker-input').length);
        __WEBPACK_IMPORTED_MODULE_1_jquery__('.dp-picker-input').addClass('form-control form-control-sm').css('width', '120px').css('text-align', 'center');
    };
    TestPanelsComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-test-panels',
            template: __webpack_require__("../../../../../src/app/components/test-panels/test-panels.component.html"),
            styles: [__webpack_require__("../../../../../src/app/components/test-panels/test-panels.component.css")]
        }),
        __metadata("design:paramtypes", [])
    ], TestPanelsComponent);
    return TestPanelsComponent;
}());

//# sourceMappingURL=test-panels.component.js.map

/***/ }),

/***/ "../../../../../src/app/directives/test-dir.directive.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return TestDirDirective; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};

var TestDirDirective = (function () {
    function TestDirDirective(viewContainer, template) {
        this.viewContainer = viewContainer;
        this.template = template;
        this.data = [
            { id: 1, name: "Item 1" },
            { id: 2, name: "Item 2" },
            { id: 3, name: "Item 3" },
            { id: 4, name: "Item 4" },
        ];
    }
    TestDirDirective.prototype.ngOnInit = function () {
        var condition = true;
        if (condition) {
            this.viewContainer.createEmbeddedView(this.template);
        }
        else {
            this.viewContainer.clear();
        }
    };
    TestDirDirective = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Directive"])({
            selector: '[testDir]'
        }),
        __metadata("design:paramtypes", [typeof (_a = typeof __WEBPACK_IMPORTED_MODULE_0__angular_core__["ViewContainerRef"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_0__angular_core__["ViewContainerRef"]) === "function" && _a || Object, typeof (_b = typeof __WEBPACK_IMPORTED_MODULE_0__angular_core__["TemplateRef"] !== "undefined" && __WEBPACK_IMPORTED_MODULE_0__angular_core__["TemplateRef"]) === "function" && _b || Object])
    ], TestDirDirective);
    return TestDirDirective;
    var _a, _b;
}());

//# sourceMappingURL=test-dir.directive.js.map

/***/ }),

/***/ "../../../../../src/app/models/data-models.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "c", function() { return dataActivity; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "b", function() { return activityMember; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return activityConflict; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "d", function() { return dataPersonnel; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "e", function() { return dataTeam; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "h", function() { return teamMember; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "g", function() { return lookupTemplate; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "f", function() { return loginDetails; });
var dataActivity = (function () {
    function dataActivity() {
        this.id = -1;
        this.site = "SWP";
        this.name = "";
        this.description = "";
        this.startDate = "";
        this.endDate = "";
        this.type = 0;
        this.showInChart = true;
        this.noPOBCount = false;
        this.ready = "Amber";
        this.members = [];
        this.conflicts = null;
        this.parent = null;
        this.vessel = false;
        this.upmanning = false;
        this.extracted = new Date();
        this.removed = false;
        this.obsolete = false;
    }
    return dataActivity;
}());

var activityMember = (function () {
    function activityMember() {
        this.id = -1;
        this.idAct = -1;
        this.perID = 0;
        this.startDate = "";
        this.endDate = "";
        this.mob = true;
        this.demob = true;
        this.night = false;
        this.isDay = false;
        this.noPOB = false;
        this.coyID = "0";
        this.posID = "0";
        this.grpID = "0";
        this.sort = "";
        this.conflicts = null;
        this.warnMessage = "";
        this.infoMessage = "";
        this.confMessage = "";
        this.warnCount = 0;
        this.infoCount = 0;
        this.confCount = 0;
    }
    return activityMember;
}());

var activityConflict = (function () {
    function activityConflict() {
    }
    return activityConflict;
}());

var dataPersonnel = (function () {
    function dataPersonnel() {
        this.status = 1;
        this.id = -1;
        this.companyID = "0";
        this.positionID = "0";
        this.gender = "";
        this.birthDate = "";
        this.maxDays = 0;
        this.xl = false;
        this.certificates = [];
    }
    return dataPersonnel;
}());

var dataTeam = (function () {
    function dataTeam() {
        this.id = -1;
        this.name = "";
        this.description = "";
        this.order = 1;
        this.core = true;
        this.upman = false;
        this.members = [];
        this.group = "";
        this.expView = false;
    }
    return dataTeam;
}());

var teamMember = (function () {
    function teamMember() {
        this.sort = "";
    }
    return teamMember;
}());

var lookupTemplate = (function () {
    function lookupTemplate() {
        this.order = 0;
    }
    return lookupTemplate;
}());

var loginDetails = (function () {
    function loginDetails() {
        this.taken = false;
        this.uid = -1;
        this.login = "visitor";
        this.fullName = "Visitor";
        this.role = "read-only";
        this.login = "visitor";
        this.fullName = "Visitor";
        this.role = "read-only";
        this.teamAdd = false;
        this.teamEdit = false;
        this.teamDelete = false;
        this.teamMemAdmin = false;
        this.activityAdd = false;
        this.activityEdit = false;
        this.activityDelete = false;
        this.activityMemAdmin = false;
        this.personnelAdd = false;
        this.personnelEdit = false;
        this.personnelDelete = false;
        this.personnelDelete = false;
        this.settingsAdmin = false;
        this.userSession = null;
        /*
                this.teamAdd=true;
                this.teamEdit=true;
                this.teamDelete=true;
                this.teamMemAdmin=true;
        
                this.activityAdd=true;
                this.activityEdit=true;
                this.activityDelete=true;
                this.activityMemAdmin=true;
        
                this.personnelAdd=true;
                this.personnelEdit=true;
                this.personnelDelete=true;
        
                this.settingsAdmin=true;
                */
    }
    return loginDetails;
}());

//# sourceMappingURL=data-models.js.map

/***/ }),

/***/ "../../../../../src/app/test-window-freeze/test-window-freeze.component.css":
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__("../../../../css-loader/lib/css-base.js")(false);
// imports


// module
exports.push([module.i, "", ""]);

// exports


/*** EXPORTS FROM exports-loader ***/
module.exports = module.exports.toString();

/***/ }),

/***/ "../../../../../src/app/test-window-freeze/test-window-freeze.component.html":
/***/ (function(module, exports) {

module.exports = "<div style=\"position:absolute;width:1024px;height:768px;top:50px;left:50px;background:yellow;overflow:auto;\"\n(scroll)=\"onElementScroll($event)\"\n>\n    <div name=\"ul_pane\"\n    style=\"z-index:2;position:absolute;background:lime;opacity:0.6;\"\n    [style.width]=\"rowHeaderWidth+'px'\"\n    [style.height]=\"colHeaderHeight+'px'\"\n    >\n        Upper Left Header\n    </div>\n\n    <div name=\"ll_pane\"\n    style=\"color:white;z-index:1;position:absolute;background:olive;opacity:0.6;\"\n    [style.margin-top]=\"colHeaderHeight+'px'\"\n    [style.height]=\"'2048px'\"\n    [style.width]=\"rowHeaderWidth +'px'\"\n    >\n        Lower Left Header\n    </div>    \n    \n\n    <div name=\"ur_pane\"\n    style=\"color:white;z-index:1;position:absolute;background:magenta;opacity:0.6;\"\n    [style.margin-left]=\"rowHeaderWidth+'px'\"\n    [style.height]=\"colHeaderHeight+'px'\"\n    [style.width]=\"'1000px'\"\n    >\n        Upper right Header\n    </div>\n\n    \n    <div name=\"lr_pane\"\n    style=\"position:absolute;background:cyan;opacity:0.5;width:1000px;height:2048px;\"\n        [style.margin-left]=\"rowHeaderWidth+'px'\"\n        [style.margin-top]=\"colHeaderHeight+'px'\"\n    >\nDetails\n    </div>\n</div>"

/***/ }),

/***/ "../../../../../src/app/test-window-freeze/test-window-freeze.component.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return TestWindowFreezeComponent; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1_jquery__ = __webpack_require__("../../../../jquery/dist/jquery.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1_jquery___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_1_jquery__);
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};


var TestWindowFreezeComponent = (function () {
    function TestWindowFreezeComponent() {
        this.rowHeaderWidth = 200;
        this.colHeaderHeight = 100;
    }
    TestWindowFreezeComponent.prototype.ngOnInit = function () {
    };
    TestWindowFreezeComponent.prototype.onElementScroll = function (e) {
        var cont = __WEBPACK_IMPORTED_MODULE_1_jquery__(e.srcElement);
        //console.log(cont.scrollTop())
        //console.log(cont.scrollLeft())
        var ul = cont.find('[name="ul_pane"]');
        ul.css('top', cont.scrollTop());
        ul.css('left', cont.scrollLeft());
        var ur = cont.find('[name="ur_pane"]');
        ur.css('top', cont.scrollTop());
        var ll = cont.find('[name="ll_pane"]');
        ll.css('left', cont.scrollLeft());
    };
    TestWindowFreezeComponent = __decorate([
        Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["Component"])({
            selector: 'app-test-window-freeze',
            template: __webpack_require__("../../../../../src/app/test-window-freeze/test-window-freeze.component.html"),
            styles: [__webpack_require__("../../../../../src/app/test-window-freeze/test-window-freeze.component.css")]
        }),
        __metadata("design:paramtypes", [])
    ], TestWindowFreezeComponent);
    return TestWindowFreezeComponent;
}());

//# sourceMappingURL=test-window-freeze.component.js.map

/***/ }),

/***/ "../../../../../src/environments/environment.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "a", function() { return environment; });
// The file contents for the current environment will overwrite these during build.
// The build system defaults to the dev environment which uses `environment.ts`, but if you do
// `ng build --env=prod` then `environment.prod.ts` will be used instead.
// The list of which env maps to which file can be found in `.angular-cli.json`.
var environment = {
    production: false
};
//# sourceMappingURL=environment.js.map

/***/ }),

/***/ "../../../../../src/main.ts":
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
Object.defineProperty(__webpack_exports__, "__esModule", { value: true });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__angular_core__ = __webpack_require__("../../../core/@angular/core.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__angular_platform_browser_dynamic__ = __webpack_require__("../../../platform-browser-dynamic/@angular/platform-browser-dynamic.es5.js");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__app_app_module__ = __webpack_require__("../../../../../src/app/app.module.ts");
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3__environments_environment__ = __webpack_require__("../../../../../src/environments/environment.ts");




if (__WEBPACK_IMPORTED_MODULE_3__environments_environment__["a" /* environment */].production) {
    Object(__WEBPACK_IMPORTED_MODULE_0__angular_core__["enableProdMode"])();
}
Object(__WEBPACK_IMPORTED_MODULE_1__angular_platform_browser_dynamic__["a" /* platformBrowserDynamic */])().bootstrapModule(__WEBPACK_IMPORTED_MODULE_2__app_app_module__["a" /* AppModule */])
    .catch(function (err) { return console.log(err); });
//# sourceMappingURL=main.js.map

/***/ }),

/***/ "../../../../moment/locale recursive ^\\.\\/.*$":
/***/ (function(module, exports, __webpack_require__) {

var map = {
	"./af": "../../../../moment/locale/af.js",
	"./af.js": "../../../../moment/locale/af.js",
	"./ar": "../../../../moment/locale/ar.js",
	"./ar-dz": "../../../../moment/locale/ar-dz.js",
	"./ar-dz.js": "../../../../moment/locale/ar-dz.js",
	"./ar-kw": "../../../../moment/locale/ar-kw.js",
	"./ar-kw.js": "../../../../moment/locale/ar-kw.js",
	"./ar-ly": "../../../../moment/locale/ar-ly.js",
	"./ar-ly.js": "../../../../moment/locale/ar-ly.js",
	"./ar-ma": "../../../../moment/locale/ar-ma.js",
	"./ar-ma.js": "../../../../moment/locale/ar-ma.js",
	"./ar-sa": "../../../../moment/locale/ar-sa.js",
	"./ar-sa.js": "../../../../moment/locale/ar-sa.js",
	"./ar-tn": "../../../../moment/locale/ar-tn.js",
	"./ar-tn.js": "../../../../moment/locale/ar-tn.js",
	"./ar.js": "../../../../moment/locale/ar.js",
	"./az": "../../../../moment/locale/az.js",
	"./az.js": "../../../../moment/locale/az.js",
	"./be": "../../../../moment/locale/be.js",
	"./be.js": "../../../../moment/locale/be.js",
	"./bg": "../../../../moment/locale/bg.js",
	"./bg.js": "../../../../moment/locale/bg.js",
	"./bn": "../../../../moment/locale/bn.js",
	"./bn.js": "../../../../moment/locale/bn.js",
	"./bo": "../../../../moment/locale/bo.js",
	"./bo.js": "../../../../moment/locale/bo.js",
	"./br": "../../../../moment/locale/br.js",
	"./br.js": "../../../../moment/locale/br.js",
	"./bs": "../../../../moment/locale/bs.js",
	"./bs.js": "../../../../moment/locale/bs.js",
	"./ca": "../../../../moment/locale/ca.js",
	"./ca.js": "../../../../moment/locale/ca.js",
	"./cs": "../../../../moment/locale/cs.js",
	"./cs.js": "../../../../moment/locale/cs.js",
	"./cv": "../../../../moment/locale/cv.js",
	"./cv.js": "../../../../moment/locale/cv.js",
	"./cy": "../../../../moment/locale/cy.js",
	"./cy.js": "../../../../moment/locale/cy.js",
	"./da": "../../../../moment/locale/da.js",
	"./da.js": "../../../../moment/locale/da.js",
	"./de": "../../../../moment/locale/de.js",
	"./de-at": "../../../../moment/locale/de-at.js",
	"./de-at.js": "../../../../moment/locale/de-at.js",
	"./de-ch": "../../../../moment/locale/de-ch.js",
	"./de-ch.js": "../../../../moment/locale/de-ch.js",
	"./de.js": "../../../../moment/locale/de.js",
	"./dv": "../../../../moment/locale/dv.js",
	"./dv.js": "../../../../moment/locale/dv.js",
	"./el": "../../../../moment/locale/el.js",
	"./el.js": "../../../../moment/locale/el.js",
	"./en-au": "../../../../moment/locale/en-au.js",
	"./en-au.js": "../../../../moment/locale/en-au.js",
	"./en-ca": "../../../../moment/locale/en-ca.js",
	"./en-ca.js": "../../../../moment/locale/en-ca.js",
	"./en-gb": "../../../../moment/locale/en-gb.js",
	"./en-gb.js": "../../../../moment/locale/en-gb.js",
	"./en-ie": "../../../../moment/locale/en-ie.js",
	"./en-ie.js": "../../../../moment/locale/en-ie.js",
	"./en-nz": "../../../../moment/locale/en-nz.js",
	"./en-nz.js": "../../../../moment/locale/en-nz.js",
	"./eo": "../../../../moment/locale/eo.js",
	"./eo.js": "../../../../moment/locale/eo.js",
	"./es": "../../../../moment/locale/es.js",
	"./es-do": "../../../../moment/locale/es-do.js",
	"./es-do.js": "../../../../moment/locale/es-do.js",
	"./es.js": "../../../../moment/locale/es.js",
	"./et": "../../../../moment/locale/et.js",
	"./et.js": "../../../../moment/locale/et.js",
	"./eu": "../../../../moment/locale/eu.js",
	"./eu.js": "../../../../moment/locale/eu.js",
	"./fa": "../../../../moment/locale/fa.js",
	"./fa.js": "../../../../moment/locale/fa.js",
	"./fi": "../../../../moment/locale/fi.js",
	"./fi.js": "../../../../moment/locale/fi.js",
	"./fo": "../../../../moment/locale/fo.js",
	"./fo.js": "../../../../moment/locale/fo.js",
	"./fr": "../../../../moment/locale/fr.js",
	"./fr-ca": "../../../../moment/locale/fr-ca.js",
	"./fr-ca.js": "../../../../moment/locale/fr-ca.js",
	"./fr-ch": "../../../../moment/locale/fr-ch.js",
	"./fr-ch.js": "../../../../moment/locale/fr-ch.js",
	"./fr.js": "../../../../moment/locale/fr.js",
	"./fy": "../../../../moment/locale/fy.js",
	"./fy.js": "../../../../moment/locale/fy.js",
	"./gd": "../../../../moment/locale/gd.js",
	"./gd.js": "../../../../moment/locale/gd.js",
	"./gl": "../../../../moment/locale/gl.js",
	"./gl.js": "../../../../moment/locale/gl.js",
	"./gom-latn": "../../../../moment/locale/gom-latn.js",
	"./gom-latn.js": "../../../../moment/locale/gom-latn.js",
	"./he": "../../../../moment/locale/he.js",
	"./he.js": "../../../../moment/locale/he.js",
	"./hi": "../../../../moment/locale/hi.js",
	"./hi.js": "../../../../moment/locale/hi.js",
	"./hr": "../../../../moment/locale/hr.js",
	"./hr.js": "../../../../moment/locale/hr.js",
	"./hu": "../../../../moment/locale/hu.js",
	"./hu.js": "../../../../moment/locale/hu.js",
	"./hy-am": "../../../../moment/locale/hy-am.js",
	"./hy-am.js": "../../../../moment/locale/hy-am.js",
	"./id": "../../../../moment/locale/id.js",
	"./id.js": "../../../../moment/locale/id.js",
	"./is": "../../../../moment/locale/is.js",
	"./is.js": "../../../../moment/locale/is.js",
	"./it": "../../../../moment/locale/it.js",
	"./it.js": "../../../../moment/locale/it.js",
	"./ja": "../../../../moment/locale/ja.js",
	"./ja.js": "../../../../moment/locale/ja.js",
	"./jv": "../../../../moment/locale/jv.js",
	"./jv.js": "../../../../moment/locale/jv.js",
	"./ka": "../../../../moment/locale/ka.js",
	"./ka.js": "../../../../moment/locale/ka.js",
	"./kk": "../../../../moment/locale/kk.js",
	"./kk.js": "../../../../moment/locale/kk.js",
	"./km": "../../../../moment/locale/km.js",
	"./km.js": "../../../../moment/locale/km.js",
	"./kn": "../../../../moment/locale/kn.js",
	"./kn.js": "../../../../moment/locale/kn.js",
	"./ko": "../../../../moment/locale/ko.js",
	"./ko.js": "../../../../moment/locale/ko.js",
	"./ky": "../../../../moment/locale/ky.js",
	"./ky.js": "../../../../moment/locale/ky.js",
	"./lb": "../../../../moment/locale/lb.js",
	"./lb.js": "../../../../moment/locale/lb.js",
	"./lo": "../../../../moment/locale/lo.js",
	"./lo.js": "../../../../moment/locale/lo.js",
	"./lt": "../../../../moment/locale/lt.js",
	"./lt.js": "../../../../moment/locale/lt.js",
	"./lv": "../../../../moment/locale/lv.js",
	"./lv.js": "../../../../moment/locale/lv.js",
	"./me": "../../../../moment/locale/me.js",
	"./me.js": "../../../../moment/locale/me.js",
	"./mi": "../../../../moment/locale/mi.js",
	"./mi.js": "../../../../moment/locale/mi.js",
	"./mk": "../../../../moment/locale/mk.js",
	"./mk.js": "../../../../moment/locale/mk.js",
	"./ml": "../../../../moment/locale/ml.js",
	"./ml.js": "../../../../moment/locale/ml.js",
	"./mr": "../../../../moment/locale/mr.js",
	"./mr.js": "../../../../moment/locale/mr.js",
	"./ms": "../../../../moment/locale/ms.js",
	"./ms-my": "../../../../moment/locale/ms-my.js",
	"./ms-my.js": "../../../../moment/locale/ms-my.js",
	"./ms.js": "../../../../moment/locale/ms.js",
	"./my": "../../../../moment/locale/my.js",
	"./my.js": "../../../../moment/locale/my.js",
	"./nb": "../../../../moment/locale/nb.js",
	"./nb.js": "../../../../moment/locale/nb.js",
	"./ne": "../../../../moment/locale/ne.js",
	"./ne.js": "../../../../moment/locale/ne.js",
	"./nl": "../../../../moment/locale/nl.js",
	"./nl-be": "../../../../moment/locale/nl-be.js",
	"./nl-be.js": "../../../../moment/locale/nl-be.js",
	"./nl.js": "../../../../moment/locale/nl.js",
	"./nn": "../../../../moment/locale/nn.js",
	"./nn.js": "../../../../moment/locale/nn.js",
	"./pa-in": "../../../../moment/locale/pa-in.js",
	"./pa-in.js": "../../../../moment/locale/pa-in.js",
	"./pl": "../../../../moment/locale/pl.js",
	"./pl.js": "../../../../moment/locale/pl.js",
	"./pt": "../../../../moment/locale/pt.js",
	"./pt-br": "../../../../moment/locale/pt-br.js",
	"./pt-br.js": "../../../../moment/locale/pt-br.js",
	"./pt.js": "../../../../moment/locale/pt.js",
	"./ro": "../../../../moment/locale/ro.js",
	"./ro.js": "../../../../moment/locale/ro.js",
	"./ru": "../../../../moment/locale/ru.js",
	"./ru.js": "../../../../moment/locale/ru.js",
	"./sd": "../../../../moment/locale/sd.js",
	"./sd.js": "../../../../moment/locale/sd.js",
	"./se": "../../../../moment/locale/se.js",
	"./se.js": "../../../../moment/locale/se.js",
	"./si": "../../../../moment/locale/si.js",
	"./si.js": "../../../../moment/locale/si.js",
	"./sk": "../../../../moment/locale/sk.js",
	"./sk.js": "../../../../moment/locale/sk.js",
	"./sl": "../../../../moment/locale/sl.js",
	"./sl.js": "../../../../moment/locale/sl.js",
	"./sq": "../../../../moment/locale/sq.js",
	"./sq.js": "../../../../moment/locale/sq.js",
	"./sr": "../../../../moment/locale/sr.js",
	"./sr-cyrl": "../../../../moment/locale/sr-cyrl.js",
	"./sr-cyrl.js": "../../../../moment/locale/sr-cyrl.js",
	"./sr.js": "../../../../moment/locale/sr.js",
	"./ss": "../../../../moment/locale/ss.js",
	"./ss.js": "../../../../moment/locale/ss.js",
	"./sv": "../../../../moment/locale/sv.js",
	"./sv.js": "../../../../moment/locale/sv.js",
	"./sw": "../../../../moment/locale/sw.js",
	"./sw.js": "../../../../moment/locale/sw.js",
	"./ta": "../../../../moment/locale/ta.js",
	"./ta.js": "../../../../moment/locale/ta.js",
	"./te": "../../../../moment/locale/te.js",
	"./te.js": "../../../../moment/locale/te.js",
	"./tet": "../../../../moment/locale/tet.js",
	"./tet.js": "../../../../moment/locale/tet.js",
	"./th": "../../../../moment/locale/th.js",
	"./th.js": "../../../../moment/locale/th.js",
	"./tl-ph": "../../../../moment/locale/tl-ph.js",
	"./tl-ph.js": "../../../../moment/locale/tl-ph.js",
	"./tlh": "../../../../moment/locale/tlh.js",
	"./tlh.js": "../../../../moment/locale/tlh.js",
	"./tr": "../../../../moment/locale/tr.js",
	"./tr.js": "../../../../moment/locale/tr.js",
	"./tzl": "../../../../moment/locale/tzl.js",
	"./tzl.js": "../../../../moment/locale/tzl.js",
	"./tzm": "../../../../moment/locale/tzm.js",
	"./tzm-latn": "../../../../moment/locale/tzm-latn.js",
	"./tzm-latn.js": "../../../../moment/locale/tzm-latn.js",
	"./tzm.js": "../../../../moment/locale/tzm.js",
	"./uk": "../../../../moment/locale/uk.js",
	"./uk.js": "../../../../moment/locale/uk.js",
	"./ur": "../../../../moment/locale/ur.js",
	"./ur.js": "../../../../moment/locale/ur.js",
	"./uz": "../../../../moment/locale/uz.js",
	"./uz-latn": "../../../../moment/locale/uz-latn.js",
	"./uz-latn.js": "../../../../moment/locale/uz-latn.js",
	"./uz.js": "../../../../moment/locale/uz.js",
	"./vi": "../../../../moment/locale/vi.js",
	"./vi.js": "../../../../moment/locale/vi.js",
	"./x-pseudo": "../../../../moment/locale/x-pseudo.js",
	"./x-pseudo.js": "../../../../moment/locale/x-pseudo.js",
	"./yo": "../../../../moment/locale/yo.js",
	"./yo.js": "../../../../moment/locale/yo.js",
	"./zh-cn": "../../../../moment/locale/zh-cn.js",
	"./zh-cn.js": "../../../../moment/locale/zh-cn.js",
	"./zh-hk": "../../../../moment/locale/zh-hk.js",
	"./zh-hk.js": "../../../../moment/locale/zh-hk.js",
	"./zh-tw": "../../../../moment/locale/zh-tw.js",
	"./zh-tw.js": "../../../../moment/locale/zh-tw.js"
};
function webpackContext(req) {
	return __webpack_require__(webpackContextResolve(req));
};
function webpackContextResolve(req) {
	var id = map[req];
	if(!(id + 1)) // check for number or string
		throw new Error("Cannot find module '" + req + "'.");
	return id;
};
webpackContext.keys = function webpackContextKeys() {
	return Object.keys(map);
};
webpackContext.resolve = webpackContextResolve;
module.exports = webpackContext;
webpackContext.id = "../../../../moment/locale recursive ^\\.\\/.*$";

/***/ }),

/***/ 0:
/***/ (function(module, exports, __webpack_require__) {

module.exports = __webpack_require__("../../../../../src/main.ts");


/***/ })

},[0]);
//# sourceMappingURL=main.bundle.js.map