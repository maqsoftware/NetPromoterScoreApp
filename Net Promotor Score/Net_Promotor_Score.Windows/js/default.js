/// <reference path="navigation.js" />
var npsCurrentCookies = null;
var npsCanary = "";
(function () {
    "use strict";
    WinJS.Binding.optimizeBindingReferences = true;
    var app = WinJS.Application;
    var activation = Windows.ApplicationModel.Activation;

    app.onactivated = function (args) {
        WinJS.Application.addEventListener("consentCheckCompleted", function (e) {
            if (e.state && e.action) {
                // User has clicked continue on first launch
                if (e.action == "allow" && e.state == "firstLaunch") {
                    AppConsentCheckCompleted(args);
                } else if (e.action == "cancel" && e.state == "firstLaunch") { // User has clicked exit on first launch
                } else if (e.action == "continue" && e.state == "upgrade") { // User has clicked continue on upgrade scenario                    
                    AppConsentCheckCompleted(args);
                } else if (e.action == "cancel" && e.state == "upgrade") { // User has clicked exit on upgrade scenario
                } else if (e.action == "pass" && e.state == "normal") { // Normal launch                    
                    AppConsentCheckCompleted(args);
                }

            }
        });
        args.setPromise(WinJS.UI.processAll().then(function () {
            AppNavigator.navigate('/pages/login.html', {});
        }));

        // LOBCompliance.load();
    };
    app.oncheckpoint = function (args) {
        // TODO: This application is about to be suspended. Save any state
        // that needs to persist across suspensions here. You might use the
        // WinJS.Application.sessionState object, which is automatically
        // saved and restored across suspension. If you need to complete an
        // asynchronous operation before your application is suspended, call
        // args.setPromise(). 
    };
    document.addEventListener('msvisibilitychange', function () {
        console.log('visibility changed');
        console.log(document.visibilityState); // 'hidden' or 'visible'
        Survey.isCameraOpen = 0;
    });    
    app.start();
    app.onerror = function errorHandler(event) {
        //WriteToCrashLog(event, true);
    }
    function AppConsentCheckCompleted(args) {
        if (args.detail.kind === activation.ActivationKind.launch) {
            if (args.detail.previousExecutionState !== activation.ApplicationExecutionState.terminated) {
                //ShowSplashScreen(args);

            } else {
                //var c = SharePointDataAccess.SharePointAccess.getSurveyList();
            }
            args.setPromise(WinJS.UI.processAll().then(function () {
                AppNavigator.navigate('/pages/login.html', {});
            }));

        }
    }
    function InitAppBarIcon() {
        WinJS.Utilities.query('#async-data-load').addClass('hideElement');
        var oAppBar = AppBar.winControl;
        if (cmdHome) {
            cmdHome.addEventListener('click', function () {
                AppNavigator.navigate('/pages/hub.html', { pagetitle: '' });
                AppNavigator.purgeHistory();
            }, false);
        }

        if (cmdEndSurvey) {
            cmdEndSurvey.addEventListener('click', function () {
                Survey.endSurvey();
            }, false);
        }

        if (cmdSettings) {
            cmdSettings.addEventListener('click', function () {
                AppNavigator.navigate('/pages/settings.html');
            }, false);
        }
        if (cmdClearData) {
            cmdClearData.addEventListener('click', function () {
                var fileName = 'Survey_' + encodeURIComponent(localStorage.getItem('npsCurrentUserName') + localStorage.getItem('npsCurrentSPUrl'));
                FileOps.removeFile(fileName, null);

            }, false);
        }
    }
    WinJS.UI.Pages.define('default.html', {
        ready: function (element, options) {
            InitAppBarIcon();
        }
    });
    WinJS.UI.Pages.define('/pages/hub.html', {
        ready: function (element, options) {
            SetAppBarState('hub');
        }
    });
    WinJS.Namespace.define('Hub', {

        navigate: function (page) {
            switch (page) {
                case 'survey': AppNavigator.navigate('/pages/survey.html'); break;
                case 'admin':
                    {
                        AppNavigator.navigate('/pages/admin.html');
                        SetAppBarState('admin');
                        break;
                    }
                case 'setting':
                    {
                        AppNavigator.navigate('/pages/settings.html');
                        SetAppBarState('setting');
                        break;
                    }
                case 'addsurvey': {
                    AppNavigator.navigate('/pages/addsurvey.html');
                    SetAppBarState('admin');
                    break;
                }
            }
        }
    });
})();
function SetAppBarState(sCurrentPage) {
    if (typeof (Survey) != 'undefined' && Survey.MessageInterval != -1) {
        clearInterval(Survey.MessageInterval)
    }
    var oAppBar = AppBar.winControl;
    oAppBar.sticky = false;
    oAppBar.disabled = false;
    oAppBar.hide();
    switch (sCurrentPage) {
        case 'survey':
            oAppBar.showOnlyCommands([cmdHome, cmdEndSurvey, cmdSettings], true)
            break;
        case 'setting':
            oAppBar.showOnlyCommands([cmdHome, cmdSettings, cmdClearData], true);
            break;
        case 'admin':
        case 'addsurvey':
            oAppBar.showOnlyCommands([cmdHome], true);
            break;
        case 'hub':
        case 'login':
            oAppBar.disabled = true;
            break;
    }
}