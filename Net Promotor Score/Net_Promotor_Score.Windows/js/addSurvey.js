/// <reference path="//Microsoft.WinJS.1.0/js/base.js" />
(function () {
    WinJS.UI.Pages.define('/pages/addsurvey.html', {
        ready: function (element, options) {
            AddSurvey.initControls();

        }
    });
    WinJS.Namespace.define('AddSurvey', {
        oSurveyFields: {},
        NewSurveyEntry: null,
        initControls: function () {
            var datePickerContainer = document.getElementById("btnTrainingDate");
            var initialDate = new Date(1990, 8, 1, 0, 0, 0, 0);
            var control = new WinJS.UI.DatePicker(datePickerContainer, { current: initialDate });
            control.addEventListener("change", function () {
                day = control.current.getDate();
                month = control.current.getMonth() + 1;
                year = control.current.getFullYear();
                AddSurvey.oSurveyFields.TrainingDate = month + "/" + day + "/" + year;
            });
        },
        save: function () {
            var isValidForm = AddSurvey.validateForm();
            if (isValidForm) {
                // Create entry and add to list                
                var entry = SharePointDataAccess.Data.NewSurvey();
                entry.addEntry(JSON.stringify(AddSurvey.oSurveyFields));
                AddSurvey.NewSurveyEntry = entry;
                AddSurvey._sendUploadData();
            }
        },
        validateForm: function () {
            var isValidForm = 1;
            AddSurvey.oSurveyFields.SurveyTitle = txtSurveyTitle.value;
            AddSurvey.oSurveyFields.Presenter = txtPresenter.value + "@maqsoftware.com";            
            AddSurvey.oSurveyFields.TrainingDuration = txtTrainingDuration.value;
            AddSurvey.oSurveyFields.SurveyStatus = statusMenuOptions.options[statusMenuOptions.selectedIndex].text;
            AddSurvey.oSurveyFields.SurveyQuestion = txtSurveyQuestion.value;
            for (var prop in AddSurvey.oSurveyFields) {
                if ("undefined" === typeof AddSurvey.oSurveyFields || "" === AddSurvey.oSurveyFields[prop]) {
                    isValidForm = 0;
                }                
            }
            if (isValidForm) {
                return true;
            } else {
                return false;
            }
        }, _sendUploadData: function () {
            var requestUri = localStorage.getItem('npsCurrentSPUrl');
            var entry = AddSurvey.NewSurveyEntry;
            var returnValue = SharePointDataAccess.SharePointAccess.createNewSurvey(requestUri, npsCurrentCookies, WinJS.Resources.getString('ReadListName').value, entry);
            if (returnValue === 'success') {
                var fileName = 'Survey_' + encodeURIComponent(localStorage.getItem('npsCurrentUserName') + requestUri);
                var msgBox = new Windows.UI.Popups.MessageDialog("New Survey created successfully!");
                msgBox.showAsync();
                AddSurvey.resetFields();
            }
            else {
                AppNavigator.navigate('/pages/hub.html', {});
            }
        }, resetFields: function () {
            txtSurveyTitle.value = "";
            txtPresenter.value = "";                        
        }
    });
})();