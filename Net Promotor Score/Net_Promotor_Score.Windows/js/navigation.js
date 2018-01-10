/// <reference path="//Microsoft.WinJS.1.0.RC/js/base.js" />
(function () {
    "use strict";
    WinJS.strictProcessing();
    WinJS.Namespace.define("AppNavigator", {
        homepage: "/pages/login.html",
        historyURL: [],
        historyOptions: [],
        currentURL: '',
        currentOptions: null,
        reloadReport: function () {
            this.navigate(this.currentURL, this.currentOptions);
        },
        purgeHistory: function () {
            this.historyURL = [];
            this.historyOptions = [];
            WinJS.Utilities.query("button.win-backbutton").addClass('hideElement');
        },
        navigate: function (url, options) {
            
            if (this.historyURL.indexOf(url) < 0) {
                this.historyOptions.push(options);
                this.historyURL.push(url);
            }
            if (this.homepage.toLowerCase() === url.toLowerCase()) {
                this.purgeHistory();
            }
            var oContentHost = document.getElementById("sAppContent");
            WinJS.Utilities.empty(oContentHost);
           
            if (this.historyURL.length) {
                WinJS.Utilities.query("button.win-backbutton").removeClass('hideElement');
            }
           
            WinJS.UI.Pages.render(url, oContentHost, options);
            this.currentURL = url;
            this.currentOptions = options;
            oContentHost.style.opacity = "0";
            setTimeout(function () {
                WinJS.UI.Animation.enterContent(oContentHost, null).done(
                    null,
                     function () {
                         oContentHost.style.opacity = "1";
                     }
                );
            }, 100);
        },
        navback: function (options) {            
            this.historyURL.length = this.historyURL.length - 1;
            this.historyOptions.length = this.historyOptions.length - 1;
            var sURL = this.historyURL[this.historyURL.length - 1];
            if (!sURL) {
                sURL = this.homepage;
            }
            if (this.homepage.toLowerCase() === sURL.toLowerCase()) {
                this.purgeHistory();
            }
            
            var oContentHost = document.getElementById("sAppContent");
            WinJS.Utilities.empty(oContentHost);
            options = this.historyOptions[this.historyURL.length - 1];
          
            WinJS.UI.Pages.render(sURL, oContentHost, options);
            this.currentURL = sURL;
            this.currentOptions = options;
            if (this.historyURL.length === 0) {
                WinJS.Utilities.query("button.win-backbutton").addClass('hideElement');
            }
            oContentHost.style.opacity = "0";
            setTimeout(function () {
                WinJS.UI.Animation.enterContent(oContentHost, null).done(
                    null,
                     function () {
                         oContentHost.style.opacity = "1";
                     }
                );
            }, 100);
        },
        bgtask: function (sURL) {
            try {
                if (sURL.toLowerCase().indexOf('licensehealth.html') > -1) {
                    WinJS.Utilities.query('#LicenseTypeSelect').removeClass('hideElement');
                } else {
                    WinJS.Utilities.query('#LicenseTypeSelect').addClass('hideElement');
                    LicenseTypeSelect.selectedIndex = 0;
                }
                WinJS.Utilities.addClass(oAsyncLoader, "hideElement");
                SalesInsightsAppFilter.ClearExtraFilters();
                NavBar.winControl.hide();
                HideFilterSummaryFlyout();
            } catch (exp) {
                console.log(exp.message);
                WriteToCrashLog(exp);
            }
        }
    });
})();