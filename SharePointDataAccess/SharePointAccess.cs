namespace SharePointDataAccess
{
    using Microsoft.SharePoint;
    using Microsoft.SharePoint.Client;
    using SharePointDataAccess.Data;
    using SP = Microsoft.SharePoint.Client;
    #region References
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Net.Http;
    using System.Text;
    using System.Threading.Tasks;
    using System.Xml.Linq;
    using Windows.Foundation;
    using Newtonsoft.Json;
    using Microsoft.SharePoint.Client.UserProfiles;
    using System.Globalization;
    
    #endregion
    public sealed class SharePointAccess
    {
        #region Constants
        const string msoAuthUrl = "https://login.microsoftonline.com/extSTS.srf";
        const string spLoginUrl = "/_forms/default.aspx?wa=wsignin1.0";
        const string spCanaryUrl = "/_api/contextinfo";
        const string spAppRedirectUrl = "/_layouts/15/appredirect.aspx?instance_id={0}";
        const string nssoap = "http://www.w3.org/2003/05/soap-envelope";
        const string nsdataservice = "http://schemas.microsoft.com/ado/2007/08/dataservices";
        const string nswssecurity = "http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd";
        const string nspssoapf = "http://schemas.microsoft.com/Passport/SoapServices/SOAPFault";
        static string FullCookieUrl = "";
        static string CanaryURL = "";
        static string MakeRequestSPsiteUrl = "";
        static string MakeRequestheaders = "";
        static string MakeRequestcanary = "";
        static string MakeRequestData = "";
        static string MakeRequestMethod = "";
        static string MakeRequestaddtionalUrlPart = "";
        static IDictionary<string, string> MakeRequestcookies = new Dictionary<string, string>();

        #endregion

        #region Private Constructor
        private SharePointAccess()
        {
        }
        #endregion

        #region Async public Methods

        public static IAsyncOperation<IDictionary<string, string>> Authenticate(string username, string password, string url)
        {
            //return new Dictionary<string, string>();
            return GetAccessToken(username, password, url);

        }
        public static IAsyncOperation<IDictionary<string, string>> GetCookies(string url, string token)
        {
            return Task.Run<IDictionary<string, string>>(() =>
            {
                Uri spSiteUrl = new Uri(url, UriKind.Absolute);
                var cookies = GetCookieContainer(spSiteUrl, token);
                IDictionary<string, string> cookieArray = new Dictionary<string, string>();
                CookieCollection collection = cookies.GetCookies(spSiteUrl);
                IEnumerator enumr = collection.GetEnumerator();
                while (enumr.MoveNext())
                {
                    var Name = ((Cookie)enumr.Current).Name;
                    var value = ((Cookie)enumr.Current).Value;
                    cookieArray.Add(Name, value);
                }
                return cookieArray;
            }).AsAsyncOperation();
        }
        public static IAsyncOperation<string> MakeRestCall(IDictionary<string, string> cookies, string url)
        {
            MakeRequestcookies = cookies;
            MakeRequestSPsiteUrl = url;
            return Task.Run<string>(() =>
            {
                Uri spSiteUrl = new Uri(MakeRequestSPsiteUrl);
                Uri requesturi = new Uri(MakeRequestSPsiteUrl);
                string ResponseData;
                // Populates the list of videos from the sharepoint site .
                var httpRequestMessage = new HttpRequestMessage(System.Net.Http.HttpMethod.Get, requesturi);
                string contentType = "application/json;odata=verbose;charset=utf-8";
                var httpClientHandler = new HttpClientHandler();
                HttpClient requestNew = new HttpClient(httpClientHandler);
                httpRequestMessage.Headers.Add("Accept", contentType); // set the content type of the request
                if (httpClientHandler.CookieContainer == null)
                    httpClientHandler.CookieContainer = new CookieContainer();

                // get the auth cookies  after authenticating with Microsoft Online Services 

                foreach (KeyValuePair<string, string> kvp in cookies)
                {
                    Cookie ck = new Cookie();
                    ck.Name = kvp.Key;
                    ck.Value = kvp.Value;
                    // append  auth cookies to the request
                    httpClientHandler.CookieContainer.Add(spSiteUrl, ck);
                }
                var resultData = Task.Run<HttpResponseMessage>(() =>
                {
                    return requestNew.SendAsync(httpRequestMessage);
                }).Result;
                ResponseData = Task.Run<string>(() =>
                {
                    return resultData.Content.ReadAsStringAsync();
                }).Result;

                MakeRequestSPsiteUrl = "";
                MakeRequestcookies = new Dictionary<string, string>();
                return ResponseData;

            }).AsAsyncOperation();

        }
        /// <summary>
        /// Makes the REST call to Sharepoint server with Request Digest in header
        /// </summary>
        /// <param name="spSiteUrl">Url of Sharepoint</param>
        /// <param name="cookies">Cookies to be added to the request</param>
        /// <param name="canary">Request Digest</param>
        /// <param name="Data">Additional data to be sent in POST/PUT/DELETE request</param>
        /// <returns></returns>
        //public static IAsyncOperation<string> MakeRestWriteCall(string spSiteUrl, IDictionary<string, string> cookies, string headers, string canary, string Data, string Method, string addtionalUrlPart)
        //{
        //    MakeRequestaddtionalUrlPart = addtionalUrlPart;
        //    MakeRequestcanary = canary;
        //    MakeRequestcookies = cookies;
        //    MakeRequestData = Data;
        //    MakeRequestheaders = headers;
        //    MakeRequestMethod = Method;
        //    MakeRequestSPsiteUrl = spSiteUrl;
        //    return Task.Run<string>(() =>
        //    {
        //        Dictionary<string, string> headersList = ConvertJsonToDictionary(MakeRequestheaders);
        //        Uri SpUri = new Uri(MakeRequestSPsiteUrl);
        //Uri spCanaryUri = new Uri(SpUri.AbsoluteUri + (addtionalUrlPart ?? ""));

        //HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(spCanaryUri.AbsoluteUri);
        //request.Method = MakeRequestMethod;
        //request.ContentType = "application/json;odata=verbose;charset=utf-8";
        //request.Accept = "application/json;odata=verbose;charset=utf-8";
        //CookieContainer cookieJar = new CookieContainer();

        //// get the auth cookies  after authenticating with Microsoft Online Services 
        //foreach (KeyValuePair<string, string> kvp in MakeRequestcookies)
        //{
        //    cookieJar.Add(SpUri, new Cookie(kvp.Key, kvp.Value));
        //}
        //request.CookieContainer = cookieJar;

        //Stream RequestStream = Task.Run(() =>
        //{
        //    return request.GetRequestStreamAsync();
        //}).Result;
        //using (StreamWriter requestWriter = new StreamWriter(RequestStream))
        //{
        //    requestWriter.Write(MakeRequestData);
        //}
        //foreach (KeyValuePair<string, string> headerKvp in headersList)
        //{
        //    request.Headers[headerKvp.Key] = headerKvp.Value;
        //}
        //request.Headers["X-RequestDigest"] = MakeRequestcanary;

        //HttpWebResponse resultData = (HttpWebResponse)Task.Run<WebResponse>(() =>
        //{
        //    return request.GetResponseAsync();
        //}).Result;
        //Stream responseStream = resultData.GetResponseStream();
        //using (StreamReader ResponseReader = new StreamReader(responseStream))
        //{
        //    return ResponseReader.ReadToEndAsync();
        //}
        //return "";

        //    }).AsAsyncOperation();
        //}

        //#endregion

        //#region Private Methods
        private static Dictionary<string, string> ConvertJsonToDictionary(string headers)
        {
            return new Dictionary<string, string>();
        }
        private static CookieContainer GetCookieContainer(Uri spSiteUrl, string token)
        {
            StringBuilder UBuilder = new StringBuilder(spSiteUrl.Scheme);
            UBuilder.Append("://");
            UBuilder.Append(spSiteUrl.Authority);
            UBuilder.Append(spLoginUrl);
            FullCookieUrl = UBuilder.ToString();
            return Task.Run(() =>
            {
                HttpWebRequest request = HttpWebRequest.CreateHttp(FullCookieUrl);
                request.CookieContainer = new CookieContainer();
                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded";
                request.Accept = "*/*";

                try
                {
                    var requestStream = Task.Run(() =>
                    {
                        return request.GetRequestStreamAsync();
                    }).Result;

                    using (var writer = new StreamWriter(requestStream))
                    {
                        writer.WriteAsync(token);
                    }

                    WebResponse response = Task.Run(() =>
                    {
                        return request.GetResponseAsync();
                    }).Result; ;

                    FullCookieUrl = "";
                    response.Dispose();
                }
                catch
                {
                    return null;
                }


                return request.CookieContainer;
            }).Result;
        }
        private static IAsyncOperation<IDictionary<string, string>> GetAccessToken(string username, string password, string url)
        {
            IDictionary<string, string> result = new Dictionary<string, string>();
            return Task.Run<IDictionary<string, string>>(() =>
            {
                using (var httpClient = new HttpClient())
                {
                    //Uri spSiteUrl = new Uri(url);
                    //string token = string.Empty;
                    using (var httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, "https://login.microsoftonline.com/extSTS.srf"))
                    {
                        string token = "";
                        Uri spSiteUrl = new Uri(url);
                        XDocument loadedData = XDocument.Load("Resources/SAML.xml");
                        string samlXML = loadedData.ToString();
                        StringBuilder tokenRequestXml = new StringBuilder(samlXML);// samlXML.Replace("{0}", username);//, password, spSiteUrl.AbsoluteUri);
                        tokenRequestXml.Replace("{0}", username);
                        tokenRequestXml.Replace("{1}", password);
                        tokenRequestXml.Replace("{2}", spSiteUrl.AbsoluteUri);
                        httpRequestMessage.Content = new StringContent(tokenRequestXml.ToString());
                        var httpResponseMessage = Task.Run(() =>
                            {
                                return httpClient.SendAsync(httpRequestMessage);
                            }).Result;
                        var stringResponse = Task.Run(() =>
                        {
                            return httpResponseMessage.Content.ReadAsStringAsync();
                        }).Result;
                        var xDoc = XDocument.Parse(stringResponse);
                        var body = xDoc.Descendants(XName.Get("Body", nssoap)).FirstOrDefault();
                        if (body != null)
                        {
                            var fault = body.Descendants(XName.Get("Fault", nssoap)).FirstOrDefault();
                            if (fault != null)
                            {
                                var error = fault.Descendants(XName.Get("text", nspssoapf)).FirstOrDefault().Value;
                                if (error != null)
                                    result.Add("fault", Convert.ToString(error));
                            }
                            else
                            {
                                var binaryToken = body.Descendants(XName.Get("BinarySecurityToken", nswssecurity)).FirstOrDefault();
                                if (binaryToken != null)
                                    result.Add("token", binaryToken.Value);

                            }
                        }
                    }
                }
                return result;
            }).AsAsyncOperation();
        }

        //#endregion

        //#region Non async Public methods
        ///// <summary>
        ///// Gets the App Context token from the Sharepoint for getting the App permission context
        ///// </summary>
        ///// <param name="spSiteUrl">Url of Sharepoint</param>
        ///// <param name="cookies">Cookies to be added to the request</param>
        ///// <param name="instanceId">InstanceId of the App</param>
        ///// <returns>AP App Context Token</returns>
        //public static string GetAppContextToken(string spSiteUrl, IDictionary<string, string> cookies, string instanceId)
        //{
        //    return Task.Run<string>(async () =>
        //    {
        //        Uri SpUri = new Uri(spSiteUrl);
        //        string spApptokenUrl = SpUri.AbsoluteUri + string.Format(spAppRedirectUrl, instanceId);
        //        var httpRequestMessage = new HttpRequestMessage(System.Net.Http.HttpMethod.Post, spApptokenUrl);
        //        var httpClientHandler = new HttpClientHandler();
        //        HttpClient requestNew = new HttpClient(httpClientHandler);
        //        CookieContainer cookieJar = new CookieContainer();

        //        foreach (KeyValuePair<string, string> kvp in cookies)
        //        {
        //            cookieJar.Add(SpUri, new Cookie(kvp.Key, kvp.Value));
        //        }
        //        httpClientHandler.CookieContainer = cookieJar;

        //        // get the auth cookies  after authenticating with Microsoft Online Services 
        //        var resultData = Task.Run<HttpResponseMessage>(async () =>
        //        {
        //            return await requestNew.SendAsync(httpRequestMessage);
        //        }).Result;
        //        return await Task.Run<string>(async () =>
        //        {
        //            var ResponseData = await resultData.Content.ReadAsStringAsync();
        //            string KeyName = "\"SPAppToken\"";
        //            int initialIndex = ResponseData.IndexOf(KeyName);
        //            int finalIndexValueFirst = ResponseData.IndexOf("\"", initialIndex + KeyName.Length + 1);
        //            int finalIndexValueLast = ResponseData.IndexOf("\"", finalIndexValueFirst + 1);
        //            string SPAppToken = ResponseData.Substring(finalIndexValueFirst + 1, (finalIndexValueLast - (finalIndexValueFirst)));
        //            return SPAppToken;
        //        });


        //    }).Result;
        //}
        ///// <summary>
        ///// Gets the Request Digest from the SharePoint used for write operations to Sharepoint
        ///// </summary>
        ///// <param name="spSiteUrl">Url of Sharepoint</param>
        ///// <param name="cookies">Cookies to be added to the request</param>
        ///// <returns>Request Digest</returns>
        public static string GetSPCanary(string spSiteUrl, IDictionary<string, string> cookies)
        {
            Uri SpUri = new Uri(spSiteUrl);
            StringBuilder canaryBuilder = new StringBuilder(SpUri.AbsoluteUri);
            canaryBuilder.Append(spCanaryUrl);
            CanaryURL = canaryBuilder.ToString();
            return Task.Run<string>(() =>
            {
                var httpRequestMessage = new HttpRequestMessage(System.Net.Http.HttpMethod.Post, CanaryURL);
                var httpClientHandler = new HttpClientHandler();
                HttpClient requestNew = new HttpClient(httpClientHandler);
                CookieContainer cookieJar = new CookieContainer();

                foreach (KeyValuePair<string, string> kvp in cookies)
                {
                    cookieJar.Add(SpUri, new Cookie(kvp.Key, kvp.Value));
                }
                httpClientHandler.CookieContainer = cookieJar;

                // get the auth cookies  after authenticating with Microsoft Online Services 
                var resultData = Task.Run<HttpResponseMessage>(() =>
                {
                    return requestNew.SendAsync(httpRequestMessage);
                }).Result;
                try
                {
                    var ResponseData = Task.Run(() =>
                    {
                        return resultData.Content.ReadAsStringAsync();
                    }).Result;
                    XDocument document = XDocument.Parse(ResponseData);
                    var Node = document.Descendants(XName.Get("FormDigestValue", nsdataservice)).FirstOrDefault();
                    return Node.Value;
                }
                catch
                {
                    return "";
                }


            }).Result;
        }
        public static string AddListItemsInBatch(string spSiteUrl, IDictionary<string, string> cookies, string ListName, SurveyEntries entries)
        {

            try
            {
                using (ClientContext context = new ClientContext(spSiteUrl))
                {
                    context.ExecutingWebRequest += delegate (object sender, WebRequestEventArgs e)
                    {
                        e.WebRequestExecutor.WebRequest.CookieContainer = new CookieContainer();
                        foreach (KeyValuePair<string, string> KvpCookie in cookies)
                        {
                            Cookie cookie = new Cookie(KvpCookie.Key, KvpCookie.Value);
                            e.WebRequestExecutor.WebRequest.CookieContainer.Add(new Uri(spSiteUrl), cookie);
                        }
                    };
                    List list = context.Web.Lists.GetByTitle(ListName);
                    foreach (Entry entry in entries.GetEntries())
                    {
                        ListItemCreationInformation info = new ListItemCreationInformation();
                        ListItem listItem = list.AddItem(info);
                        listItem["Title"] = entry.Title;
                        listItem["Promoter"] = entry.PromoterString;
                        listItem["Detractor"] = entry.DetractorString;
                        listItem["Passive"] = entry.PassiveString;
                        var Survey = new FieldLookupValue();
                        Survey.LookupId = Convert.ToInt32(entry.SurveyLookupIdString);
                        listItem["Survey"] = Survey;
                        listItem["EmotionScore"] = entry.emotionScore;
                        listItem.Update();
                    }
                    context.ExecuteQuery();

                }

                return "success";
            }
            catch (Exception e)
            {
                return "failed";
            }

        }
        public static string CreateNewSurvey(string spSiteUrl, IDictionary<string, string> cookies, string ListName, NewSurvey entries)
        {

            try
            {
                using (ClientContext context = new ClientContext(spSiteUrl))
                {
                    context.ExecutingWebRequest += delegate (object sender, WebRequestEventArgs e)
                    {
                        e.WebRequestExecutor.WebRequest.CookieContainer = new CookieContainer();
                        foreach (KeyValuePair<string, string> KvpCookie in cookies)
                        {
                            Cookie cookie = new Cookie(KvpCookie.Key, KvpCookie.Value);
                            e.WebRequestExecutor.WebRequest.CookieContainer.Add(new Uri(spSiteUrl), cookie);
                        }
                    };

                    Entry entry = entries.GetEntry();
                    User newUser = context.Web.EnsureUser(entry.Presenter);
                    context.Load(newUser);
                    context.ExecuteQuery();
                    FieldUserValue userValue = new FieldUserValue();
                    userValue.LookupId = newUser.Id;
                    List list = context.Web.Lists.GetByTitle(ListName);
                    ListItemCreationInformation info = new ListItemCreationInformation();
                    ListItem listItem = list.AddItem(info);
                    listItem["Title"] = entry.Title;
                    listItem["Presenter"] = userValue;
                    listItem["TrainingDate"] = entry.TrainingDate;
                    listItem["TrainingDuration"] = entry.TrainingDuration;
                    listItem["SurveyStatus"] = entry.SurveyStatus;
                    listItem["SurveyStatus_old"] = entry.SurveyStatus;
                    listItem["Question"] = entry.SurveyQuestion;
                    listItem.Update();
                    context.ExecuteQuery();
                }
                return "success";
            }
            catch (Exception e)
            {
                return "failed";
            }

        }

        public static string updateAndCloseSurveyList(string spSiteUrl, IDictionary<string, string> cookies, string ListName, SurveyEntries entries)
        {
            try
            {
                using (ClientContext context = new ClientContext(spSiteUrl))
                {
                    context.ExecutingWebRequest += delegate (object sender, WebRequestEventArgs e)
                    {
                        e.WebRequestExecutor.WebRequest.CookieContainer = new CookieContainer();
                        foreach (KeyValuePair<string, string> KvpCookie in cookies)
                        {
                            Cookie cookie = new Cookie(KvpCookie.Key, KvpCookie.Value);
                            e.WebRequestExecutor.WebRequest.CookieContainer.Add(new Uri(spSiteUrl), cookie);
                        }
                    };
                    string sDetractorString = "", sPassiveString = "", sPromoterString = "", sEmotionScore = "", sSurveyName = "";
                    foreach (Entry entry in entries.GetEntries())
                    {
                        sDetractorString = entry.DetractorString;
                        sPassiveString = entry.PassiveString;
                        sPromoterString = entry.PromoterString;
                        sEmotionScore = entry.emotionScore;
                        sSurveyName = entry.Title;
                    }
                    var web = context.Web;
                    context.Load(web);
                    context.ExecuteQuery();
                    List oList = context.Web.Lists.GetByTitle(ListName);
                    context.Load(oList);
                    context.ExecuteQuery();
                    CamlQuery query = new CamlQuery();
                    query.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>" + sSurveyName + "</Value></Eq></Where><OrderBy><FieldRef Name='SurveyStatus' Ascending='False' /></OrderBy></Query><ViewFields><FieldRef Name='Promoter' /><FieldRef Name='Passive' /><FieldRef Name='Detractor' /><FieldRef Name='EmotionScore' /><FieldRef Name='NPSScore' /><FieldRef Name='Title' /><FieldRef Name='SurveyStatus' /></ViewFields><QueryOptions /></View>";
                    ListItemCollection oListItem = oList.GetItems(query);
                    context.Load(oListItem);
                    context.ExecuteQuery();
                    ListItem oItem = oListItem[0];
                    if (oListItem.Count > 0)
                    {
                        RootEmotion jsonEmotionObject = JsonConvert.DeserializeObject<RootEmotion>(sEmotionScore);
                        oItem["Detractor"] = sDetractorString;
                        oItem["Passive"] = sPassiveString;
                        oItem["Promoter"] = sPromoterString;
                        oItem["EmotionScore"] = sEmotionScore;
                        oItem["NetEmotionScore"] = (null == jsonEmotionObject.NetScore ? float.Parse("0.00") : float.Parse((string)jsonEmotionObject.NetScore, CultureInfo.InvariantCulture.NumberFormat));
                        oItem["SurveyStatus"] = "Closed";
                        oItem.Update();
                        context.ExecuteQuery();
                        return "success";
                    }
                    else
                    {
                        return "failed";
                    }
                }
            }
            catch (Exception e)
            {
                return "failed";
            }
        }
        #endregion
    }
    public sealed class Emotion
    {
        public float Weight { get; set; }
        public float Count { get; set; }
    }
    public sealed class RootEmotion
    {
        public string NetScore { get; set; }
        public Emotion HAPPY { get; set; }
        public Emotion SURPRISED { get; set; }
        public Emotion CALM { get; set; }
        public Emotion CONFUSED { get; set; }
        public Emotion UNKNOWN { get; set; }
        public Emotion DISGUSTED { get; set; }
        public Emotion SAD { get; set; }
        public Emotion ANGRY { get; set; }
    }

}
