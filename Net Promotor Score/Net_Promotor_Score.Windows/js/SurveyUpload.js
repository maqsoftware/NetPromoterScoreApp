(function () {
    WinJS.Namespace.define("SurveyUpload", {
        FileData: null,
        uploadEntries: null,
        redirectpage: null,
        messagepopup: null,
        uploadSurveyData: function (toPage) {
            SurveyUpload.uploadEntries = null;
            SurveyUpload.FileData = null;
            SurveyUpload.redirectpage = toPage;
            SurveyUpload.AvailabilityCheck = false;
            var entries = SurveyUpload.getUploadEntries(WinJS.Resources.getString('WriteListName').value);
        },
        clearSurveyData: function (fileName) {
            FileOps.removeFile(fileName, function (data) {
                Windows.Storage.ApplicationData.current.localFolder.getFolderAsync("NPSData").then(function (folder) {
                    return folder.deleteAsync();
                }).done(function () {
                    AppNavigator.navigate('/pages/' + SurveyUpload.redirectpage, {});
                }, function (err) {
                    AppNavigator.navigate('/pages/' + SurveyUpload.redirectpage, {});
                });
            }, function () {
                AppNavigator.navigate('/pages/' + SurveyUpload.redirectpage, {});
            });
        },
        _sendUploadData: function () {
            var requestUri = localStorage.getItem('npsCurrentSPUrl');
            var entries = SurveyUpload.uploadEntries;
            var returnValue = SharePointDataAccess.SharePointAccess.addListItemsInBatch(requestUri, npsCurrentCookies, WinJS.Resources.getString('WriteListName').value, entries);
            if (returnValue === 'success') {
                var fileName = 'Survey_' + encodeURIComponent(localStorage.getItem('npsCurrentUserName') + localStorage.getItem('npsCurrentSPUrl'));
                SurveyUpload.clearSurveyData(fileName);
            }
            else {
                AppNavigator.navigate('/pages/' + SurveyUpload.redirectpage, {});
            }
        },
        getFileData: function (fileName) {
            FileOps.retrieveFromFile(fileName, localStorage.getItem('npsCurrentUserName'), SurveyUpload.returnFileData);

        },
        returnFileData: function (data, availabilityRequest) {
            SurveyUpload.FileData = data;
            if (data == "" || data==null) {
                AppNavigator.navigate('/pages/' + SurveyUpload.redirectpage, {})
            }
            else {
                SurveyUpload.createUploadEntries();
            }
        },
        getUploadEntries: function (listName) {
            var fileName = 'Survey_' + encodeURIComponent(localStorage.getItem('npsCurrentUserName') + localStorage.getItem('npsCurrentSPUrl'));
            SurveyUpload.getFileData(fileName);
        },
        createUploadEntries: function () {            
            var processeddata = FileOps.processSurveyData(SurveyUpload.FileData);
            var entries = SharePointDataAccess.Data.SurveyEntries();
            var listName = WinJS.Resources.getString('WriteListName').value;
            for (var npsEntry in processeddata.surveys) {
                entries.addEntry(localStorage.getItem("npsCurrentUserName") + " --> " + new Date().toGMTString(), processeddata.surveys[npsEntry].Passive, processeddata.surveys[npsEntry].Promoter, processeddata.surveys[npsEntry].Detractor, JSON.stringify(processeddata.surveys[npsEntry].Emotion), processeddata.surveys[npsEntry].surveyid, "SP.Data." + listName + "ListItem");                
            }
            SurveyUpload.uploadEntries = entries;
            SurveyUpload._sendUploadData();
        }
    })
})();