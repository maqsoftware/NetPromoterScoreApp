/// <reference path="//Microsoft.WinJS.1.0/js/base.js" />
(function () {
    var storageFolder = Windows.Storage.ApplicationData.current.localFolder;
    WinJS.Namespace.define('FileOps', {
        saveToFile: function (data, fileName, password, callBackFn) {
            storageFolder.createFileAsync(fileName, Windows.Storage.CreationCollisionOption.openIfExists).then(
                function (file) {
                    Windows.Storage.FileIO.appendTextAsync(file, data).then(
                        function (successdata) {
                            callBackFn('success');
                        },
                        function (errordata) {
                            callBackFn(null);
                        });
                },
            function (filerror) {
                callBackFn(null);
            });
        },
        retrieveFromFile: function (fileName, password, callBackFn) {
            storageFolder.getFileAsync(fileName).then(
                function (file) {
                    Windows.Storage.FileIO.readTextAsync(file).then(
                        function (data) {
                            callBackFn(data);
                        },
                        function () {
                            callBackFn(null);
                        });
                },
            function () {
                callBackFn(null);
            });
        },
        removeFile: function (fileName, callBackFn) {
            storageFolder.getFileAsync(fileName).then(
                function (file) {
                    file.deleteAsync(Windows.Storage.StorageDeleteOption.default).then(
                        function (status) {
                            if (callBackFn)
                                callBackFn('success');
                        },
                        function (error) {
                            if (callBackFn)
                                callBackFn(null);
                        });
                },
            function (error) {
                if (callBackFn)
                    callBackFn(null);
            });
        },
        processSurveyData: function (data) {
            if (data && data.length > 0) {
                var trimmeddata = data.substr(0, data.length - 1);
                var CoveredJsonString = '{"raw":[' + trimmeddata + ']}';
                var dataJson = JSON.parse(CoveredJsonString);
                dataJson.surveys = FileOps.processEachSurveyData(dataJson);
                FileOps.processEmotionData(dataJson);
                dataJson.nps = FileOps.processNpsScore(dataJson);
            }
            else {
            }
            return dataJson;
        },
        processEmotionData: function (surveyData) {
            var iTotalWeights = 0,
            CurrentEmotion = {};
            if (surveyData && "undefined" != typeof surveyData.raw) {
                var data = surveyData.raw;
                for (list in surveyData.surveys) {
                    CurrentEmotion = JSON.parse(JSON.stringify(Survey.oDefaultEmotion));
                    for (surveyentry in data) {
                        if (data[surveyentry].survey === surveyData.surveys[list].surveyid) {
                            for (var property in data[surveyentry].Emotion) {
                                if (property !== "NetScore") {
                                    CurrentEmotion[property].Count += data[surveyentry].Emotion[property].Count;
                                }
                            }
                        }
                    }
                    surveyData.surveys[list].Emotion = FileOps.calculateAvgEmotionScore(CurrentEmotion);
                }
            }
        },
        calculateAvgEmotionScore: function (CurrentEmotion) {
            var promoters = 0, detractors = 0, passives = 0, percentPromoters = 0, percentDetractors = 0, EmotionNPS = 0, iTotal = 0;
            for (var property in CurrentEmotion) {
                switch (property) {
                    case "HAPPY":
                    case "SURPRISED":
                        promoters += CurrentEmotion[property].Count;
                        break;
                    case "CALM":
                    case "CONFUSED":
                    case "UNKNOWN":
                        passives += CurrentEmotion[property].Count;
                        break;
                    case "DISGUSTED":
                    case "SAD":
                    case "ANGRY":
                        detractors += CurrentEmotion[property].Count;
                        break;
                }
            }
            iTotal = promoters + passives + detractors;
            percentPromoters = (promoters / iTotal) * 100;
            percentDetractors = (detractors / iTotal) * 100;
            EmotionNPS = percentPromoters - percentDetractors;
            CurrentEmotion["NetScore"] = (isNaN(EmotionNPS) ? null : EmotionNPS);
            return CurrentEmotion;
        },
        processEachSurveyData: function (data) {
            var processedSurveys = {};
            if (data.raw) {
                for (entry in data.raw) {
                    var CurrentSurvey;
                    var surveyentry = data.raw[entry];
                    if (processedSurveys.hasOwnProperty(surveyentry.survey)) {
                        CurrentSurvey = processedSurveys[surveyentry.survey];
                    }
                    else {
                        CurrentSurvey = {};
                        CurrentSurvey.surveyid = surveyentry.survey;
                        CurrentSurvey.Detractor = 0;
                        CurrentSurvey.Passive = 0;
                        CurrentSurvey.Promoter = 0;
                        CurrentSurvey.surveyName = surveyentry.surveyName;
                        CurrentSurvey.Presenter = surveyentry.Presenter;
                        CurrentSurvey.Emotion = surveyentry.Emotion;
                    }
                    if (surveyentry.value >= 0 && surveyentry.value <= 6) {
                        CurrentSurvey.Detractor += 1;
                    }
                    else if (surveyentry.value >= 7 && surveyentry.value <= 8) {
                        CurrentSurvey.Passive += 1;
                    }
                    else if (surveyentry.value >= 9 && surveyentry.value <= 10) {
                        CurrentSurvey.Promoter += 1;
                    }
                    processedSurveys[CurrentSurvey.surveyid] = CurrentSurvey;
                }
            }

            return processedSurveys;
        },
        processNpsScore: function (data) {
            var processedNpsSurveys = {};
            if (data.surveys) {
                for (entry in data.surveys) {
                    var currentSurvey = data.surveys[entry];
                    var CurrentSurvey;
                    if (processedNpsSurveys.hasOwnProperty(entry)) {
                        CurrentSurvey = processedNpsSurveys[entry];
                    }
                    else {
                        CurrentSurvey = {};
                        CurrentSurvey.surveyid = currentSurvey.surveyid;
                        CurrentSurvey.score = -1;
                    }
                    var percentPromoters = (currentSurvey.Promoter / (currentSurvey.Detractor + currentSurvey.Passive + currentSurvey.Promoter)) * 100;
                    var percentDetractors = (currentSurvey.Detractor / (currentSurvey.Detractor + currentSurvey.Passive + currentSurvey.Promoter)) * 100;
                    var nps = percentPromoters - percentDetractors;
                    CurrentSurvey.score = nps;
                    processedNpsSurveys[CurrentSurvey.surveyid] = CurrentSurvey;

                }
            }
            return processedNpsSurveys;
        }
    });
})();