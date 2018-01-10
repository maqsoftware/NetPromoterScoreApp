/// <reference path="fileoperations.js" />

(function () {
    WinJS.UI.Pages.define('/pages/admin.html', {

        ready: function (element, options) {
            var SPUrl = localStorage.getItem("npsCurrentSPUrl", SPUrl);
            var data = Results.getAllSurveys(WinJS.Resources.getString('ReadListName').value, SPUrl, Results.showSurveys);
            document.addEventListener('keyup', function (event) {
                if (event.keyCode == 13) {
                    Survey.SubmitFeedback();
                }
            });
            SetAppBarState('admin');
        }
    });
    WinJS.Namespace.define('Results', {
        resultData: {},
        iflagShowAllData: 0,
        oInProgressSurveysData: [],
        isExpanded: false,
        getAllSurveys: function (listname, spUrl, callBackFn) {
            var listExtension = "/_api/lists/GetByTitle('" + listname + "')/items?$select=Id,Title,SurveyStatus,Promoter,Passive,Detractor,EmotionScore,NPSScore,Presenter/Title&$Expand=Presenter&$orderby=TrainingDate desc'";
            SharePointDataAccess.SharePointAccess.makeRestCall(npsCurrentCookies, spUrl + listExtension).then(function (data) {
                Results.iflagShowAllData = 1
                callBackFn(data, Results.iflagShowAllData);
            });
        },
        getInProgressSurveys: function (listname, spUrl, callBackFn) {
            var iCount = 0;
            var listExtension = "/_api/web/lists/GetByTitle('" + listname + "')/items?$select=Title,Promoter,Passive,Detractor,EmotionScore,Survey,Survey/Title&$expand=Survey&$filter=Survey/Title eq '";
            for (survey in Results.resultData.d.results) {
                if (Results.resultData.d.results[survey].SurveyStatus == 'In Progress') {
                    iCount++;
                    if (1 === iCount || 0 === iCount) {
                        listExtension += Results.resultData.d.results[survey].Title + "'";
                    } else {
                        listExtension += " or Survey/Title eq '" + Results.resultData.d.results[survey].Title + "'";
                    }
                }
            }
            if (0 === iCount) {
                listExtension += "'";
            }
            SharePointDataAccess.SharePointAccess.makeRestCall(npsCurrentCookies, spUrl + listExtension).then(function (data) {
                callBackFn(data);
            });
        },
        renderDropDown: function (iflagShowAllData) {
            if (document.getElementById('NpsShowAll').checked) {
                Results.iflagShowAllData = 1;
            }
            else {
                Results.iflagShowAllData = 0;
            }
            Results.bindData(Results.iflagShowAllData);
        },
        showInProgressSurveys: function (data) {
            Results.oInProgressSurveysData = [];
            Results.iflagShowAllData = 0;
            if (data && data.length > 0) {
                var SurveyData = JSON.parse(data);
                if ("undefined" !== typeof SurveyData.d.results) {
                    var lookup = {};
                    var items = SurveyData.d.results;
                    var result = [];
                    for (var item, iIterate = 0; item = items[iIterate++];) {
                        var name = item.Survey.Title;
                        if (!(name in lookup)) {
                            lookup[name] = 1;
                            result.push(name);
                        }
                    }
                    for (item in result) {
                        var percentPromoters, percentDetractors, npsScore;
                        var oJSONResult = {
                            SurveyEntries: []
                        };
                        var oEmotionResult = JSON.parse(JSON.stringify(Survey.oDefaultEmotion));
                        var oTempObject = Results.getObjectByPropertyValue(Results.resultData.d.results, result[item]);
                        for (entry in SurveyData.d.results) {
                            if (result[item] === SurveyData.d.results[entry].Survey.Title) {
                                for (prop in SurveyData.d.results[entry]) {
                                    if ("Detractor" === prop || "Passive" === prop || "Promoter" === prop) {
                                        if ("undefined" !== typeof oJSONResult[prop]) {
                                            oJSONResult[prop] += SurveyData.d.results[entry][prop];
                                        } else {
                                            oJSONResult[prop] = 0;
                                            oJSONResult[prop] += SurveyData.d.results[entry][prop];
                                        }
                                    } else if ("EmotionScore" === prop) {
                                        var oEmotionObject = {};
                                        oEmotionObject = JSON.parse(SurveyData.d.results[entry][prop]);
                                        for (var property in oEmotionObject) {
                                            if (property !== "NetScore") {
                                                oEmotionResult[property].Count += parseInt(oEmotionObject[property].Count);
                                            }
                                        }
                                    }
                                }
                                oJSONResult.SurveyEntries.push(SurveyData.d.results[entry]);
                            }
                        }
                        oJSONResult.EmotionScore = JSON.stringify(FileOps.calculateAvgEmotionScore(oEmotionResult));
                        oJSONResult.Title = result[item];
                        oJSONResult.SurveyStatus = oTempObject.SurveyStatus;
                        oJSONResult.ID = oTempObject.ID;
                        oJSONResult.Presenter = oTempObject.Presenter;
                        percentPromoters = (oJSONResult.Promoter / (oJSONResult.Detractor + oJSONResult.Passive + oJSONResult.Promoter)) * 100;
                        percentDetractors = (oJSONResult.Detractor / (oJSONResult.Detractor + oJSONResult.Passive + oJSONResult.Promoter)) * 100;
                        oJSONResult.NPSScore = percentPromoters - percentDetractors;
                        Results.oInProgressSurveysData.push(oJSONResult);
                    }
                }
            }
            Results.bindData(Results.iflagShowAllData);
        },
        getObjectByPropertyValue: function (oObject, sPropertyValue) {
            for (var item in oObject) {
                if (oObject[item].Title === sPropertyValue) {
                    return oObject[item];
                }
            }
        },
        bindData: function (iflagShowAllData) {
            var surveySelector = WinJS.Utilities.query('.SurveySelector');
            var firstId = null;
            document.getElementById("surveyContainer").innerHTML = "";
            if (Results.resultData || Results.oInProgressSurveysData.length > 0) {

                if (iflagShowAllData == 0) {
                    if (Results.oInProgressSurveysData.length > 0) {
                        for (survey in Results.oInProgressSurveysData) {
                            var option = document.createElement('option');
                            option.value = survey;
                            option.innerHTML = Results.oInProgressSurveysData[survey].Title;
                            surveySelector[0].appendChild(option);
                            if (!firstId) {
                                firstId = survey;
                            }
                        }
                    } else {
                        Results.showDefaultReportData();
                    }
                }
                else {
                    for (survey in Results.resultData.d.results) {
                        var option = document.createElement('option');
                        option.value = survey;
                        option.innerHTML = Results.resultData.d.results[survey].Title;
                        surveySelector[0].appendChild(option);
                        if (!firstId) {
                            firstId = survey;
                        }
                    }
                }
                if (firstId)
                    Results.getResults(firstId)
            }
        },
        showSurveys: function (data, iflagShowAllData) {
            var iHasInProgress = 0;
            var SPUrl = localStorage.getItem("npsCurrentSPUrl", SPUrl);
            if (data && data.length > 0) {
                var SurveyData = JSON.parse(data);
            }
            if (!SurveyData) {
                var msgpopup = new Windows.UI.Popups.MessageDialog("There are no survey records.");
                msgpopup.commands.append(new Windows.UI.Popups.UICommand("Ok", function () {
                }));
                msgpopup.showAsync().then(
                    function () {
                        AppNavigator.navigate('/pages/hub.html', {});
                    },
                    function () { }
                    );
            }
            else {
                Results.resultData = SurveyData;
                for (survey in Results.resultData.d.results) {
                    if (Results.resultData.d.results[survey].SurveyStatus == 'In Progress') {
                        iHasInProgress = 1;
                        break;
                    }
                }
                if (iHasInProgress) {
                    Results.getInProgressSurveys(WinJS.Resources.getString('WriteListName').value, SPUrl, Results.showInProgressSurveys);
                } else {
                    Results.bindData(iflagShowAllData);
                    document.getElementById('NpsShowAll').checked = true;
                    var msgpopup = new Windows.UI.Popups.MessageDialog("There are no in progress survey records.");
                    msgpopup.commands.append(new Windows.UI.Popups.UICommand("Ok", function () {
                    }));
                    msgpopup.showAsync().then(
                    function () {
                    },
                    function () { }
                    );
                }
            }
        },       
        getEmotionsCount: function (oEmotions) {
            document.getElementById("EmotionCounter").innerHTML = "";
            document.getElementById("spanEmotionScore").innerHTML = "";
            var promoters = 0, passives = 0, detractors = 0;
            for (var property in oEmotions) {
                switch (property) {
                    case "HAPPY":
                    case "SURPRISED":
                        promoters += oEmotions[property].Count;
                        break;
                    case "CALM":
                    case "CONFUSED":
                    case "UNKNOWN":
                        passives += oEmotions[property].Count;
                        break;
                    case "DISGUSTED":
                    case "SAD":
                    case "ANGRY":
                        detractors += oEmotions[property].Count;
                        break;
                }
            }
            for (var property in oEmotions) {
                if (property !== "NetScore") {
                    document.getElementById("EmotionCounter").innerHTML += "<div class='emotionContainer'>" + Results.getEmoticon(property) + "<div class='emotioncount'>" + (("HAPPY" ===property ||  "SURPRISED" === property)?promoters: ("CALM" ===property ||  "CONFUSED" === property ||  "UNKNOWN" === property)?passives:detractors) + "</div><div class='emotion'>" + property + " (" + oEmotions[property].Count + ")</div></div>";
                } else {
                    document.getElementById("spanEmotionScore").innerHTML = (isNaN(parseFloat(oEmotions["NetScore"])) ? '0.00' : parseFloat(oEmotions["NetScore"]).toFixed(2));
                }
            }
            document.getElementById("gradientScale").innerHTML = "";
            Results.repeat(Results.renderGradient, 100, (isNaN(parseFloat(oEmotions["NetScore"])) ? 0 : parseFloat(oEmotions["NetScore"]).toFixed(2)), "gradientScale");
        },
        getEmoticon: function (sEmotion) {
            var emoticon = "";
            switch (sEmotion) {
                case "HAPPY":
                    emoticon = "<img src='/images/happy.png' class='emoticon' />";
                    break;
                case "SAD":
                    emoticon = "<img src='/images/sad.png' class='emoticon' />";
                    break;
                case "ANGRY":
                    emoticon = "<img src='/images/angry.png' class='emoticon' />";
                    break;
                case "CONFUSED":
                    emoticon = "<img src='/images/confused.png' class='emoticon' />";
                    break;
                case "DISGUSTED":
                    emoticon = "<img src='/images/disgusted.png' class='emoticon' />";
                    break;
                case "SURPRISED":
                    emoticon = "<img src='/images/surprised.png' class='emoticon' />";
                    break;
                case "CALM":
                    emoticon = "<img src='/images/calm.png' class='emoticon' />";
                    break;
                case "UNKNOWN":
                    emoticon = "<img src='/images/unknown.png' class='emoticon' />";
                    break;
            }
            return emoticon;
        },
        closeSurvey: function () {
            var iSelectedIndex = document.getElementById("surveyContainer").value,
            oIPList,
            oCurrentSurvey = {},
            listName = WinJS.Resources.getString('ReadListName').value;
            if (Results.iflagShowAllData) {
                oCurrentSurvey = Results.resultData.d.results[iSelectedIndex];
            } else {
                oCurrentSurvey = Results.oInProgressSurveysData[iSelectedIndex];
            }
            if ("undefined" !== typeof oCurrentSurvey && "In Progress" === oCurrentSurvey.SurveyStatus) {
                if (Results.oInProgressSurveysData && Results.oInProgressSurveysData.length > 0) {
                    oIPList = Results.getObjectByPropertyValue(Results.oInProgressSurveysData, oCurrentSurvey.Title);
                    var entries = SharePointDataAccess.Data.SurveyEntries();
                    entries.addEntry(oIPList.Title, oIPList.Passive, oIPList.Promoter, oIPList.Detractor, oIPList.EmotionScore, oIPList.ID, "SP.Data." + listName + "ListItem");
                    var returnValue = SharePointDataAccess.SharePointAccess.updateAndCloseSurveyList(localStorage.getItem('npsCurrentSPUrl'), npsCurrentCookies, listName, entries);
                    if (returnValue === "success") {
                        Results.oInProgressSurveysData.pop(oIPList);
                        var msgpopup = new Windows.UI.Popups.MessageDialog("Email sent successfully.");
                        msgpopup.commands.append(new Windows.UI.Popups.UICommand("Ok", function () {
                        }));
                        msgpopup.showAsync().then(
                            function () {
                                WinJS.Promise.timeout(500).then(function () {
                                    AppNavigator.navigate('/pages/hub.html', {});
                                });
                            });
                    } else {
                        var msgBox = new Windows.UI.Popups.MessageDialog("Something went wrong :( \nTry again later.");
                        msgBox.showAsync();
                    }
                }
            }
            else {
                var msgBox = new Windows.UI.Popups.MessageDialog("This survey is not in 'In Progress' state.");
                msgBox.showAsync();
            }
        },
        showDefaultReportData: function () {
            var oDefaultEmotion = JSON.parse(JSON.stringify(Survey.oDefaultEmotion));
            Results.getEmotionsCount(oDefaultEmotion);
            spanNpsScore.innerText = "0.00";
            spnNormalText.innerText = 0;
            spnHappyText.innerText = 0;
            spnSadText.innerText = 0;
        },
        renderInProgressSurvey: function (oIPList) {
            var percentPromoters = 0, percentDetractors = 0;
            if (0 === oIPList.Promoter === oIPList.Passive === oIPList.Detractor) {
                spanNpsScore.innerText = 0;
            }
            else {
                percentPromoters = (oIPList.Promoter / (oIPList.Detractor + oIPList.Passive + oIPList.Promoter)) * 100;
                percentDetractors = (oIPList.Detractor / (oIPList.Detractor + oIPList.Passive + oIPList.Promoter)) * 100;
                oIPList.NPSScore = percentPromoters - percentDetractors;
                spanNpsScore.innerText = parseFloat(oIPList.NPSScore).toFixed(2);
                spnNormalText.innerText = oIPList.Passive;
                spnHappyText.innerText = oIPList.Promoter;
                spnSadText.innerText = oIPList.Detractor;
                // Show indicator
                document.getElementById("NPSgradientScale").innerHTML = "";
                Results.repeat(Results.renderGradient, 100, parseFloat(spanNpsScore.innerText), "NPSgradientScale");
                if (null !== oIPList.EmotionScore && oIPList.EmotionScore.length > 0) {
                    var oEmotionData = JSON.parse(oIPList.EmotionScore);
                    Results.getEmotionsCount(oEmotionData);
                } else {
                    var oDefaultEmotion = JSON.parse(JSON.stringify(Survey.oDefaultEmotion));
                    Results.getEmotionsCount(oDefaultEmotion);
                }
            }
        },
        getSelectedSurveyOptions: function () {
            var checkboxes = document.getElementsByName("checkbox");
            var checkboxesChecked = [];
            // loop over them all
            for (var i = 0; i < checkboxes.length; i++) {
                // And stick the checked ones onto an array...
                if (checkboxes[i].checked) {
                    checkboxesChecked.push(checkboxes[i].value);
                }
            }
            // Return the array if it is non-empty, or null
            return checkboxesChecked.length > 0 ? checkboxesChecked : null;
        },
        renderSingleSurvey: function () {
            Results.showCheckboxes();
            var surveyid = document.getElementById("surveyContainer").value;
            var selectedResult = Results.getSelectedSurveyOptions();
            var surveyEntryID = "";
            if (null !== selectedResult) {
                surveyEntryID = selectedResult[0];
            } else {
                surveyEntryID = "Aggregated";
            }
            var NPS = Results.resultData.d.results[surveyid];
            if ("Aggregated" === surveyEntryID) {
                Results.getResults(surveyid);
            } else {
                if (Results.iflagShowAllData) {
                    var oSurvey = Results.getObjectByPropertyValue(Results.oInProgressSurveysData, NPS.Title);
                } else {
                    var oSurvey = Results.oInProgressSurveysData[surveyid];
                }
                if (selectedResult.length > 1) {
                    var oJSONResult = {};
                    var oEmotionResult = JSON.parse(JSON.stringify(Survey.oDefaultEmotion));
                    for (entry in selectedResult) {
                        var item = parseInt(selectedResult[entry]);
                        for (prop in oSurvey.SurveyEntries[item]) {
                            if ("Detractor" === prop || "Passive" === prop || "Promoter" === prop) {
                                if ("undefined" !== typeof oJSONResult[prop]) {
                                    oJSONResult[prop] += oSurvey.SurveyEntries[item][prop];
                                } else {
                                    oJSONResult[prop] = 0;
                                    oJSONResult[prop] += oSurvey.SurveyEntries[item][prop];
                                }
                            } else if ("EmotionScore" === prop) {
                                var oEmotionObject = {};
                                oEmotionObject = JSON.parse(oSurvey.SurveyEntries[item][prop]);
                                for (var property in oEmotionObject) {
                                    if (property !== "NetScore") {
                                        oEmotionResult[property].Count += parseInt(oEmotionObject[property].Count);
                                    }
                                }
                            }
                        }
                    }
                    oJSONResult.EmotionScore = JSON.stringify(FileOps.calculateAvgEmotionScore(oEmotionResult));
                    Results.renderInProgressSurvey(oJSONResult);
                } else {
                    var oIPList = oSurvey.SurveyEntries[surveyEntryID];
                    Results.renderInProgressSurvey(oIPList);
                }
            }
        },
        showCheckboxes: function () {
            var checkboxes = document.getElementById("checkboxes");
            if (!Results.isExpanded) {
                checkboxes.style.display = "block";
                Results.isExpanded = true;
            } else {
                checkboxes.style.display = "none";
                Results.isExpanded = false;
            }
        },
        getResults: function (id) {
            if (!id) {
                id = document.getElementById("surveyContainer").value;
            }
            var Selector = WinJS.Utilities.query('.SurveySelector');
            var surveyid;
            if (Selector.length > 0) {
                surveyid = Selector[0].value;
            }
            if (id)
                surveyid = id;
            var NPS = Results.resultData.d.results[surveyid];
            var oCheckboxContainer = document.getElementById("checkboxes");
            if (!Results.iflagShowAllData || "In Progress" === NPS.SurveyStatus) {
                if ((Results.oInProgressSurveysData && Results.oInProgressSurveysData.length > 0)) {
                    if (Results.iflagShowAllData) {
                        var oIPList = Results.getObjectByPropertyValue(Results.oInProgressSurveysData, NPS.Title);
                    } else {
                        var oIPList = Results.oInProgressSurveysData[surveyid];
                    }
                    if ("undefined" === typeof oIPList) {
                        Results.showDefaultReportData();
                    } else {
                        MSApp.execUnsafeLocalFunction(function () {
                            oCheckboxContainer.innerHTML = '<label><input type="checkbox" checked value="Aggregated" name="checkbox"/><span>Aggregated</span></label>';
                            for (entry in oIPList.SurveyEntries) {
                                oCheckboxContainer.innerHTML += '<label><input type="checkbox" name="checkbox" value="' + entry + '" /><span>' + oIPList.SurveyEntries[entry].Title + '</span></label>';
                            }
                        });
                        document.getElementById("multiSelectDropDown").className = "";
                        document.getElementById("btnSingleSurvey").className = "";
                        Results.renderInProgressSurvey(oIPList);
                    }
                } else {
                    Results.showDefaultReportData();
                }
            } else {
                document.getElementById("multiSelectDropDown").className = "hide";
                document.getElementById("btnSingleSurvey").className = "hide";
                Results.renderInProgressSurvey(NPS);
            }
        },
        renderGradient: function (iCount, netscore, renderID) {
            var item = "";
            if ((Math.round(netscore / 10) * 10) == iCount.toFixed(2)) {
                item = "<li class='gradientBox avgEmotionHighlighter' style='background-color:" + Results.percentToRGB(iCount) + "'  title='" + netscore + "'>" + iCount + "</li>";
            } else {
                item = "<li class='gradientBox' style='background-color:" + Results.percentToRGB(iCount) + "' title='" + netscore + "'>" + iCount + "</li>";
            }
            document.getElementById(renderID).innerHTML += item;
        },
        percentToRGB: function (percent) {
            if (percent === 100) {
                percent = 99
            }
            var r, g, b;
            if (percent > 50) {
                // yellow to green
                r = Math.floor(255 * ((50 - percent % 50) / 50));
                g = 255;
            } else {
                // red to yellow
                r = 255;
                g = Math.floor(255 * (percent / 50));
            }
            b = 0;
            return "rgb(" + r + "," + g + "," + b + ")";
        },
        repeat: function (fn, times, netscore, renderID) {
            for (var iCount = -100; iCount <= times; iCount += 10) fn(iCount, netscore, renderID);
        }
    });
})();

