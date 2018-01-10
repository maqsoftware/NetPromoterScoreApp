
using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Xml.Linq;
using System.Threading.Tasks;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace SharePointDataAccess.Data
{
    // Authenticates the user on Office 365 and gets the cookie container with the cookies FedAuth and rtFa.
    public sealed class SurveyEntries
    {
        private List<Entry> _entries = new List<Entry>();
        public void AddEntry(string Title, int passive, int promoter, int detractor, string emotionScore, int surveyLookupId, string type)
        {
            Entry entry = Entry.CreateEntry(Title, passive, promoter, detractor, emotionScore, surveyLookupId, type);
            _entries.Add(entry);
        }
        public void AddEntry(Entry entry)
        {
            _entries.Add(entry);
        }
        public IList<Entry> GetEntries()
        {
            return _entries;
        }
    }

    public sealed class NewSurvey {
        private Entry _ent;
        public void AddEntry(string SurveyEntry) {        
        dynamic jsonObject = JsonConvert.DeserializeObject(SurveyEntry);
        Entry entry = Entry.NewSurveyEntry(jsonObject);
        _ent = entry;
        }
        public Entry GetEntry() {
            return _ent;
        }
    }

    public sealed class Entry
    {
        private string _title, _type, _emotionScore;
        private int _passive, _promoter, _detractor, _surveyLookupId;
        private string _presenter;
        private string _surveyquestion;
        private string _trainingduration;
        private string _trainingdate;
        private string _surveystatus;

        public string emotionScore
        {
            get
            {
                return _emotionScore;
            }
        }
        public string Title
        {
            get
            {
                return _title;
            }
        }
        public string Type
        {
            get
            {
                return _type;
            }
        }
        public int Passive
        {
            get
            {
                return _passive;
            }
        }
        public string PassiveString
        {
            get
            {
                return Convert.ToString(_passive);
            }
        }
        public int Promoter
        {
            get
            {
                return _promoter;
            }
        }
        public string PromoterString
        {
            get
            {
                return Convert.ToString(_promoter);
            }
        }
        public string DetractorString
        {
            get
            {
                return Convert.ToString(_detractor);
            }
        }
        public int Detractor
        {
            get
            {
                return _detractor;
            }
        }
        public int SurveyLookupId
        {
            get
            {
                return _surveyLookupId;
            }
        }
        public string SurveyLookupIdString
        {
            get
            {
                return Convert.ToString(_surveyLookupId);
            }
        }
        public string Presenter
        {
            get
            {
                return Convert.ToString(_presenter);
            }
        }
        public string TrainingDate
        {
            get
            {
                return Convert.ToString(_trainingdate);
            }
        }
        public string TrainingDuration {
            get
            {
                return Convert.ToString(_trainingduration);
            }
        }
        public string SurveyStatus
        {
            get
            {
                return Convert.ToString(_surveystatus);
            }
        }
        public string SurveyQuestion
        {
            get
            {
                return Convert.ToString(_surveyquestion);
            }
        }

        private Entry() { }

        public static Entry CreateEntry(string Title, int passive, int promoter, int detractor, string emotionScore, int surveyLookupId, string type)
        {
            Entry ent = new Entry();
            ent._title = Title;
            ent._passive = passive;
            ent._promoter = promoter;
            ent._detractor = detractor;
            ent._emotionScore = emotionScore;
            ent._surveyLookupId = surveyLookupId;
            ent._type = type;
            return ent;
        }

       public static Entry NewSurveyEntry(dynamic jsonObject)
        {
            Entry ent = new Entry();
            ent._title = jsonObject.SurveyTitle;
            ent._presenter = jsonObject.Presenter;
            ent._trainingdate = jsonObject.TrainingDate;
            ent._trainingduration = jsonObject.TrainingDuration;
            ent._surveystatus = jsonObject.SurveyStatus;
            ent._surveyquestion = jsonObject.SurveyQuestion;
            return ent;
        }
    }


}

