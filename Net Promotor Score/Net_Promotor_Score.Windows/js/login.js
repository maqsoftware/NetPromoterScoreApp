/// <reference path="//Microsoft.WinJS.1.0/js/base.js" />
/// <reference path="navigation.js" />
/// <reference path="SurveyUpload.js" />

(function () {
    WinJS.Namespace.define("Login", {
        authenticate: function (uname, password) {
            var uName = txtUserName.value;
            var Password = txtPassword.value;
            var SPUrl = txtUrl.value;
            lblError.classList.add('Information');
            lblError.innerHTML = "Authenticating...";
            Login.setFormDisable(true);
            
            SharePointDataAccess.SharePointAccess.authenticate(uName, Password, SPUrl).then(function (data) {
                if (data.hasKey('token')) {
                    localStorage.setItem("npsCurrentToken", JSON.stringify(data));
                    SharePointDataAccess.SharePointAccess.getCookies(SPUrl, data["token"]).then(function (cookies) {
                        if (cookies.size > 0 && cookies.hasKey('FedAuth')) {
                            txtPassword.value = '';
                            npsCurrentCookies = cookies
                            localStorage.setItem("npsCurrentSPUrl", SPUrl);
                            localStorage.setItem("npsCurrentUserName", uName);
                            npsCanary = SharePointDataAccess.SharePointAccess.getSPCanary(SPUrl, npsCurrentCookies);
                            SurveyUpload.uploadSurveyData('hub.html');
                        }
                        else {
                            Login.setFormDisable(false);
                            lblError.classList.remove('Information');
                            lblError.innerHTML = "Authentication Failure";
                        }
                    }, function (d) {
                        Login.setFormDisable(false);
                        lblError.classList.remove('Information');
                        lblError.innerHTML = "Authentication Failure";
                    });
                }
                else {
                    Login.setFormDisable(false);
                    lblError.classList.remove('Information');
                    lblError.innerHTML = data["fault"];
                }
            }, function (d) {
                Login.setFormDisable(false);
                lblError.classList.remove('Information');
                lblError.innerHTML = "Authentication Failure";
            });
        },
        setFormDisable: function (state) {
            txtUserName.disabled = state;
            txtPassword.disabled = state;
            txtUrl.disabled = state;
            btnClear.disabled = state;
            btnLogin.disabled = state;
        },
        clear: function () {
            txtUserName.value = '';
            txtPassword.value = '';
            txtUrl.value = '';

        }

    });

    WinJS.UI.Pages.define("/pages/login.html", {
        ready: function (element, options) {
            SetAppBarState('login');
        }
    });

})();


