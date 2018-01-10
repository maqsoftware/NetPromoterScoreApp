/// <reference path="//Microsoft.WinJS.1.0/js/base.js" />
/// <reference path="navigation.js" />
/// <reference path="fileoperations.js" />
/// <reference path="login.js" />

(function () {
    WinJS.UI.Pages.define("/pages/survey.html", {
        ready: function (element, options) {
            var SPUrl = localStorage.getItem('npsCurrentSPUrl');
            SetAppBarState('survey');
            WinJS.Utilities.query('#async-data-load').removeClass('hideElement');
            WinJS.Utilities.query('#SubmitContainer').addClass('hideElement');

            var data = Survey.getSurveyData(WinJS.Resources.getString('ReadListName').value, WinJS.Resources.getString('ReadListStatus').value, '2', SPUrl, Survey.renderSelectSurveyDropDown);
            document.addEventListener('keyup', function (event) {
                if (event.keyCode == 13) {
                    Survey.SubmitFeedback();
                }
            });

            Survey.initializeDialog();
        },
    });
    WinJS.Namespace.define("Survey", {
        oDefaultEmotion: {
            "NetScore": 0.0,
            "HAPPY": {
                "Weight": 1,
                "Count": 0
            },
            "SURPRISED": {
                "Weight": 0.8,
                "Count": 0
            },
            "CALM": {
                "Weight": 0.7,
                "Count": 0
            },
            "CONFUSED": {
                "Weight": 0.5,
                "Count": 0
            },
            "UNKNOWN": {
                "Weight": 0,
                "Count": 0
            },
            "DISGUSTED": {
                "Weight": 0.4,
                "Count": 0
            },
            "SAD": {
                "Weight": 0.3,
                "Count": 0
            },
            "ANGRY": {
                "Weight": 0.2,
                "Count": 0
            }
        },
        isExpanded: false,
        dialog: {},
        oNewSurveyList: [],
        fileName: 'Survey_' + encodeURIComponent(localStorage.getItem('npsCurrentUserName') + localStorage.getItem('npsCurrentSPUrl')),
        feedbackString: "",
        sDialogChoice: "Yes",
        isCameraOpen: 1,
        openCamera: function () {
            Windows.Devices.Enumeration.DeviceInformation.findAllAsync(Windows.Devices.Enumeration.DeviceClass.videoCapture).done(function (devices) {
                if (devices.length > 0) {
                    mediaCapture = new Windows.Media.Capture.MediaCapture();
                    var settings = new Windows.Media.Capture.MediaCaptureInitializationSettings();
                    if ('undefined' !== typeof devices[1]) {
                        settings.videoDeviceId = devices[1].id;
                    } else {
                        settings.videoDeviceId = devices[0].id;
                    }
                    mediaCapture.initializeAsync(settings).done(function () {
                        livePreview = document.getElementById("live-preview");
                        livePreview.src = URL.createObjectURL(mediaCapture);
                        livePreview.play();
                        Survey.isCameraOpen = 1;
                    }, function () {
                        livePreview = document.getElementById("application");
                        livePreview.innerHTML = "Unable to load video";
                        Survey.isCameraOpen = 0;
                    });
                } else {
                    livePreview = document.getElementById("application");
                    livePreview.innerHTML = "Unable to load video";
                    Survey.isCameraOpen = 0;
                }
            });
        },
        closeCamera: function () {            
                //mediaCapture.close();
                mediaCapture = null;
                Survey.isCameraOpen = 0;                       
            //document.getElementById("live-preview").src = "";
        },
        initializeDialog: function () {
            Survey.dialog = Windows.UI.Popups.MessageDialog("Do you want to continue without AWS Emotion Recognition", "Confirmation");
            Survey.dialog.commands.append(new Windows.UI.Popups.UICommand("Yes", Survey.dialogYesHandler));
            Survey.dialog.commands.append(new Windows.UI.Popups.UICommand("No", Survey.dialogNoHandler));
        },
        takePhoto: function (feedbackString) {
            if (Survey.isCameraOpen) {
                var oResponse = "",
            folder = Windows.Storage.ApplicationData.current.localFolder,
            photoProperties = Windows.Media.MediaProperties.ImageEncodingProperties.createJpeg(),
            reader = new FileReader(),
            oPictureLib = Windows.Storage.KnownFolders.picturesLibrary,
            oCreds = VaultManager.readCredential(),
            isAWSCredsAvailable = 0;
                Survey.feedbackString = feedbackString;
                if (oCreds && oCreds.hasOwnProperty("AWSAccessKey") && oCreds.hasOwnProperty("AWSSecretKey")) {
                    if (oCreds.AWSAccessKey !== "" && oCreds.AWSSecretKey !== "") {
                        isAWSCredsAvailable = 1;
                    }
                }
                oPictureLib.createFileAsync("photo.jpg", Windows.Storage.CreationCollisionOption.generateUniqueName)
                  .then(function (file) {
                      try{
                          mediaCapture.capturePhotoToStorageFileAsync(photoProperties, file).then(function () {
                              //document.getElementById("imgCapture").src = URL.createObjectURL(file);
                              reader.readAsDataURL(file);
                              reader.onload = function () {
                                  folder.createFolderAsync("NPSData", Windows.Storage.CreationCollisionOption.openIfExists).then(function (newFolder) {
                                      console.log(newFolder.path);
                                      newFolder.createFileAsync(file.name.replace(/\.[^/.]+$/, "") + ".txt", Windows.Storage.CreationCollisionOption.generateUniqueName).then(function (txtfile) {
                                          return Windows.Storage.FileIO.writeTextAsync(txtfile, reader.result);
                                      }).then(function () {
                                          file.deleteAsync().done(function () {
                                          });
                                          if (isAWSCredsAvailable) {
                                              EmotionRecognitionAPI.EmotionDetectionHandler.getFaceDetails(reader.result, oCreds.AWSAccessKey, oCreds.AWSSecretKey).then(function (data) {
                                                  Survey.manipulateImageResponse(feedbackString, data);
                                              }, function () {
                                                  // If timeout or some error occurred while getting data from Amazon API then write without Emotion data
                                                  FileOps.saveToFile(Survey.feedbackString, Survey.fileName, localStorage.getItem('npsCurrentUserName'), Survey.returnFileOperation);
                                              });
                                          } else {
                                              if ("Yes" === Survey.sDialogChoice) {
                                                  Survey.dialog.defaultCommandIndex = 0;
                                                  Survey.dialog.showAsync();
                                              } else {
                                                  Survey.dialogYesHandler();
                                              }
                                          }
                                      });
                                  });
                              };
                          }, function () {
                              // If Camera is closed in between the survey, still continue taking feedback in old fashion
                              FileOps.saveToFile(Survey.feedbackString, Survey.fileName, localStorage.getItem('npsCurrentUserName'), Survey.returnFileOperation);
                          });
                      }catch(error){
                          // If camera is not initialized then continue in old fashion
                          FileOps.saveToFile(Survey.feedbackString, Survey.fileName, localStorage.getItem('npsCurrentUserName'), Survey.returnFileOperation);
                      }
                  });
            } else {
                // If camera is not initialized then continue in old fashion
                FileOps.saveToFile(feedbackString, Survey.fileName, localStorage.getItem('npsCurrentUserName'), Survey.returnFileOperation);
            }
        },
        dialogYesHandler: function () {
            FileOps.saveToFile(Survey.feedbackString, Survey.fileName, localStorage.getItem('npsCurrentUserName'), Survey.returnFileOperation);
            Survey.sDialogChoice = "No";
        },
        dialogNoHandler: function () {
            Hub.navigate('setting');
            Survey.sDialogChoice = "Yes";
        },
        manipulateImageResponse: function (feedbackString, oResponse) {
            var oJSONResponse = {};
            var oEmotion = JSON.parse(JSON.stringify(Survey.oDefaultEmotion));//FileOps.clone(Survey.oDefaultEmotion);
            var fileName = 'Survey_' + encodeURIComponent(localStorage.getItem('npsCurrentUserName') + localStorage.getItem('npsCurrentSPUrl'));
            var iNetEmotionScore = 0, sFeedbackWithEmotionString = "";
            var oFeedbackString = {};
            if (oResponse) {
                try {
                    oJSONResponse = JSON.parse(oResponse);
                    var trimmeddata = feedbackString.substr(0, feedbackString.length - 1);
                    var CoveredJsonString = '{"raw":[' + trimmeddata + ']}';
                    oFeedbackString = JSON.parse(CoveredJsonString);
                    if (oJSONResponse && oJSONResponse.length > 0 && oFeedbackString && oFeedbackString.raw.length) {
                        oEmotion[oJSONResponse[0].Emotions[0].Type.Value].Count += 1;
                        //iNetEmotionScore += (oJSONResponse[0].Emotions[0].Confidence / 100) * oEmotion[oJSONResponse[0].Emotions[0].Type.Value].Weight;
                        //oEmotion.NetScore = iNetEmotionScore.toFixed(2);
                        for (var iCount = 0; iCount < oFeedbackString.raw.length; iCount++) {
                            oFeedbackString.raw[iCount].Emotion = oEmotion;
                        }
                        sFeedbackWithEmotionString = JSON.stringify(oFeedbackString);
                        sFeedbackWithEmotionString = sFeedbackWithEmotionString.substring(8, sFeedbackWithEmotionString.length - 2) + ",";
                        FileOps.saveToFile(sFeedbackWithEmotionString, fileName, localStorage.getItem('npsCurrentUserName'), function () { });
                    } else {
                        // If response is 0 (No face detected) then write without Emotion data
                        FileOps.saveToFile(Survey.feedbackString, Survey.fileName, localStorage.getItem('npsCurrentUserName'), Survey.returnFileOperation);
                    }
                } catch (e) {
                    console.log(e.message);
                    // If something broke while parsing Amazon data then write without Emotion data
                    FileOps.saveToFile(Survey.feedbackString, Survey.fileName, localStorage.getItem('npsCurrentUserName'), Survey.returnFileOperation);
                }
            } else {
                // If response is empty in Amazon data then write without Emotion data
                FileOps.saveToFile(Survey.feedbackString, Survey.fileName, localStorage.getItem('npsCurrentUserName'), Survey.returnFileOperation);
            }
        },
        getEmotionScores: function () {
            var options = new Windows.Storage.Search.QueryOptions(Windows.Storage.Search.CommonFileQuery.defaultQuery, ['*']);
            var fileNames = new Array();
            Windows.Storage.ApplicationData.current.localFolder.createFolderAsync("NPSData", Windows.Storage.CreationCollisionOption.openIfExists).then(function (newFolder) {
                options.folderDepth = Windows.Storage.Search.FolderDepth.deep;
                newFolder.createFileQueryWithOptions(options).getFilesAsync().then(function (files) {
                    files.forEach(function (x) {
                        fileNames.push(x.name);
                    });
                    fileNames.sort();
                    for (var iCount = 0; iCount < fileNames.length; iCount++) {
                        console.log(fileNames[iCount]);
                        newFolder.createFileAsync(fileNames[iCount], Windows.Storage.CreationCollisionOption.openIfExists).then(function (txtfile) {
                            txtfile.openAsync(Windows.Storage.FileAccessMode.read).then(function (stream) {
                                var size = stream.size;
                                if (size == 0) {
                                    // Data not found
                                    console.log("image file is empty");
                                }
                                else {
                                    var inputStream = stream.getInputStreamAt(0);
                                    var reader = new Windows.Storage.Streams.DataReader(inputStream);

                                    reader.loadAsync(size).then(function () {
                                        var contents = reader.readString(size);
                                        console.log(contents);
                                        EmotionRecognitionAPI.EmotionDetectionHandler.getFaceDetails(contents).then(function (data) {
                                            Survey.manipulateImageResponse(txtfile.name, data);
                                        });
                                    });
                                }
                            });
                        });
                    }
                });
            });
        },
        clearMessage: function () {
            if (lblMessage)
                lblMessage.innerHTML = ''
        },
        returnFileOperation: function (result) {
            Survey.writeMessage("Survey Submitted", true);
            WinJS.Utilities.query('input[type="radio"]').forEach(function (value, index, object) {
                value.checked = false;
            });
            var element = WinJS.Utilities.query('.SurveyDataContainer');
            element.forEach(function (currobj, index) {
                currobj.setAttribute('data-value', "-1")
            })

        },
        showCheckboxes: function () {
            var checkboxes = document.getElementById("surveyOptions");
            if (!Survey.isExpanded) {
                checkboxes.style.display = "block";
                Survey.isExpanded = true;
            } else {
                checkboxes.style.display = "none";
                Survey.isExpanded = false;
            }
        },
        renderUpdatedSurveyList: function () {
            var checkboxes = document.getElementsByName("checkbox");
            var checkboxesChecked = [];
            var newSurveyData = {
                d: {
                    results: []
                }
            };
            var oResponse = "";
            if (Survey.oNewSurveyList.length) {
                // loop over them all
                for (var iCount = 0; iCount < checkboxes.length; iCount++) {
                    // And stick the checked ones onto an array...
                    if (checkboxes[iCount].checked) {
                        checkboxesChecked.push(checkboxes[iCount].value);
                    }
                }
                for (iItem in checkboxesChecked) {
                    for (entry in Survey.oNewSurveyList) {
                        if (Survey.oNewSurveyList[entry].ID == checkboxesChecked[iItem]) {
                            newSurveyData.d.results.push(Survey.oNewSurveyList[entry]);
                        }
                    }
                }
                oResponse = JSON.stringify(newSurveyData);
                document.getElementById("selectMultipleSurvey").className += " hide";
                document.getElementById("btnCreateNewSurveys").className = "hide";
                Survey.createSurveyTabs(oResponse);
            }
        },
        renderSelectSurveyDropDown: function (sResult) {
            data = JSON.parse(sResult);
            var results = data.d.results;
            Survey.oNewSurveyList = results;
            if (results.length > 1) {
                document.getElementById("selectMultipleSurvey").className = "multiselect";
                document.getElementById("btnCreateNewSurveys").className = "";
                WinJS.Utilities.query('#async-data-load').addClass('hideElement');
                MSApp.execUnsafeLocalFunction(function () {
                    document.getElementById("surveyOptions").innerHTML = "";
                    for (entry in results) {
                        document.getElementById("surveyOptions").innerHTML += '<label><input type="checkbox" name="checkbox" value="' + results[entry].ID + '" /><span>' + results[entry].Title + '</span></label>';
                    }
                });
            } else {
                document.getElementById("selectMultipleSurvey").className += " hide";
                document.getElementById("btnCreateNewSurveys").className = "hide";
                Survey.createSurveyTabs(sResult);
            }
        },
        createSurveyTabs: function (data) {
            data = JSON.parse(data);
            var template = WinJS.Utilities.query('#surveyEntryContainerTemplate');
            if (template.length > 0) {
                template = template[0].innerHTML;
            }
            var results = data.d.results;
            var resultSet = [];
            if (results.length) {
                for (var counter = 0; counter < results.length; counter++) {
                    var currentResult = {}
                    currentResult.counter = results.length - counter;
                    currentResult.question = results[counter].Question;
                    currentResult.survey = results[counter].Title;
                    currentResult.classCounter = 'SurveyRdGroup_' + counter;
                    currentResult.rdSurveyName = 'rdSurvey_' + results[counter].Id;
                    currentResult.SurveyId = results[counter].Id + '$!$' + results[counter].Title;
                    currentResult.Presenter = results[counter].Presenter.Title;

                    resultSet.push(currentResult);
                }
                var templateElement = document.getElementById("surveyEntryContainerTemplate");
                var renderElement = document.getElementById("surveyEntryContainerCover");
                renderElement.innerHTML = "";
                var templateControl = templateElement.winControl;
                var selected = resultSet.length - 1;
                while (selected >= 0) {
                    templateElement.winControl.render(resultSet[selected--], renderElement);
                }
                WinJS.Utilities.query('input[type="radio"]').forEach(function (value, index, object) {
                    value.addEventListener('click', function () {
                        Survey.surveySelected(this);
                        lblMessage.innerHTML = '';
                    }, false);
                });
                WinJS.Utilities.query('#SubmitContainer').removeClass('hideElement');
                Survey.openCamera();
            }
            WinJS.Utilities.query('#async-data-load').addClass('hideElement');
        },
        SubmitFeedback: function () {
            var feedbackString = "";
            var Feedback = { survey: -1, value: 0 };
            WinJS.Utilities.query('.SurveyDataContainer').forEach(
                function (obj, val, arr) {
                    var Feedback = { survey: -1, value: 0 };
                    Feedback.value = obj.getAttribute('data-value');
                    Feedback.Presenter = obj.title;
                    var name = obj["name"] ? obj["name"] : obj.getAttribute('name');
                    var names = name.split('$!$');
                    Feedback.survey = names[0];
                    Feedback.surveyName = names[1];
                    if (Feedback.value >= 0 && Feedback.value <= 10) {
                        feedbackString += JSON.stringify(Feedback) + ',';
                    }
                });
            // Take picture on click of submit button
            Survey.takePhoto(feedbackString);
            Survey.returnFileOperation();
        },
        surveySelected: function (obj) {
            if (obj.parentElement && obj.parentElement.parentElement && obj.parentElement.parentElement.parentElement) {
                obj.parentElement.parentElement.parentElement.setAttribute('data-value', obj.value);

            }
        },
        getSurveyData: function (listname, state, maxcount, spUrl, callBackFn) {
            var listExtension = "/_api/lists/GetByTitle('" + listname + "')/items?$select=Id,Title,Question,Presenter/Title&$Expand=Presenter&$filter=SurveyStatus eq '" + state + "'&$top=" + maxcount + '';
            SharePointDataAccess.SharePointAccess.makeRestCall(npsCurrentCookies, spUrl + listExtension).then(function (data) {
                callBackFn(data);
            });

        },
        endSurvey: function () {
            Survey.closeCamera();
            SurveyUpload.uploadSurveyData('hub.html');
            
        },
        writeMessage: function (message, successmessage) {
            lblMessage.innerHTML = message;
            if (successmessage) {
                lblMessage.classList.remove('ErrorMessage');
            }
            else {
                lblMessage.classList.add('ErrorMessage');
            }
            if (Survey.MessageInterval)
                clearInterval(Survey.MessageInterval);
            Survey.MessageInterval = setInterval(Survey.clearMessage, 3000);
        },
        MessageInterval: -1
    });


})();
