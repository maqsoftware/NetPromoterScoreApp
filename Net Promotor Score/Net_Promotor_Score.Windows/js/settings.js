/// <reference path="fileoperations.js" />
/// <reference path="//Microsoft.WinJS.1.0/js/base.js" />
(function () {
    WinJS.UI.Pages.define('/pages/settings.html', {
        ready: function (element, options) {
            SetAppBarState('setting');            
            VaultManager.asyncVaultLoad();            
            var oCreds = VaultManager.readCredential();
            Settings.load(oCreds);
        }
    });


    WinJS.Namespace.define('Settings', {
        save: function () {
            var fileName = 'Config_' + encodeURIComponent(localStorage.getItem('npsCurrentUserName') + localStorage.getItem('npsCurrentSPUrl'))
            var Config = {};
            Config.maxSurveys = txtMaxSurveys.value;
            Config.QuestionList = txtQuestions.value;
            Config.ResultsList = txtResults.value;
            Config.SurveyList = txtSurveyList.value;
            FileOps.saveToFile(JSON.stringify(Config), fileName, localStorage.getItem('npsCurrentUserName'), Settings.saveCallback);
            VaultManager.addCredential(txtAWSAccessKey.value, txtAWSSecretKey.value);
            AppNavigator.navigate('/pages/hub.html', {});
            AppNavigator.purgeHistory();
        },
        load: function (oCreds) {
            if (oCreds) {
                document.getElementById("txtAWSAccessKey").value = oCreds.AWSAccessKey;
                document.getElementById("txtAWSSecretKey").value = oCreds.AWSSecretKey;
            }
            var fileName = 'Config_' + encodeURIComponent(localStorage.getItem('npsCurrentUserName') + localStorage.getItem('npsCurrentSPUrl'));
            FileOps.retrieveFromFile(fileName, localStorage.getItem('npsCurrentUserName'), Settings.loadCallback);           
        },
        saveCallback: function (data) {
            if (data) {

            }
        },
        loadCallback: function (data) {
            if (data) {

            }
        }
    });

})();