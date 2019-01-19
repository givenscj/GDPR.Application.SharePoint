using GDPR.Common;
using GDPR.Common.Classes;
using Microsoft.Office.Server.Search.Query;
using Microsoft.Office.Server.UserProfiles;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Newtonsoft.Json;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Data;

namespace GDPR.Applications
{
    public class SharePoint : BaseGDPROAuthApplication
    {
        HttpHelper hh = new HttpHelper();

        public SharePoint()
        {
            this._version = "1.0.0.0";

            this._supportsEmailSearch = true;
            this._supportsPersonalSearch = true;
            this._supportsPhoneSearch = true;
        }

        public override void Init()
        {
            base.Init();
        }

        public override string GetAuthorizationUrl()
        {
            return "";
        }

        public override List<Record> GetAllRecordTypes()
        {
            List<Record> items = new List<Record>();

            items.Add(new Record { Type = "SPListItem" });
            items.Add(new Record { Type = "SPUser" });
            items.Add(new Record { Type = "SPUserProfile" });

            return items;
        }

        public override void Install(bool reInit)
        {
            base.Install(reInit);

            CreateSecurityProperties(reInit);

            core.SaveEntityProperties(this.ApplicationId, this.Properties, false);
        }

        public override void AnonymizeRecord(Record r)
        {
            base.AnonymizeRecord(r);
        }

        public dynamic ExecuteApiRequest(string url, string method, string data)
        {
            string domain = GetProperty("TenantDomain");
            string finalUrl = string.Format("https://{0}.sharepoint.com/{1}", domain, url);
            hh.contentTypeOverride = "application/json";
            hh.acceptOverride = "application/json";
            //hh.headers.Add("Authorization", "Bearer " + this.AccessToken);
            string html = "";

            if (method == "GET")
                html = hh.DoGet(finalUrl, "");

            if (method == "POST")
                html = hh.DoPost(finalUrl, data, "");

            if (method == "DELETE")
                html = hh.DoDelete(finalUrl, data, "");

            dynamic json = JsonConvert.DeserializeObject(html);
            return json;
        }

        public List<Record> DoSharePointSearch(GDPRSubject search)
        {
            List<Record> records = new List<Record>();

            try
            {
                using (SPSite siteCollection = GetFirstSPSite())
                {
                    KeywordQuery keywordQuery = new KeywordQuery(siteCollection);

                    foreach (GDPRSubjectEmail se in search.EmailAddresses)
                    {
                        keywordQuery.QueryText = se.EmailAddress;
                        SearchExecutor searchExecutor = new SearchExecutor();
                        ResultTableCollection resultTableCollection = searchExecutor.ExecuteQuery(keywordQuery);
                        ResultTable resultTable = resultTableCollection.Filter("TableType", KnownTableTypes.RelevantResults).FirstOrDefault();
                        DataTable dataTable = resultTable.Table;
                    }
                }
            }
            catch (Exception ex)
            {
                //GDPRCore.Current.Log(ex, LogLevel.Error);
            }

            return records;
        }

        public List<Record> DoProfileSearch(GDPRSubject search)
        {
            List<Record> records = new List<Record>();

            //do userprofile
            foreach (GDPRSubjectEmail se in search.EmailAddresses)
            {
                Record r = new Record();

                string loginName = se.EmailAddress;

                //testing 
                loginName = "administrator@sanspug.org";
                //loginName = Utility.UrlEncode(loginName);

                UserProfileManager mgr = GetUserProfileManager();
                UserProfile up = mgr.GetUserProfile(Guid.Parse("82712b6f-bfe0-4353-905e-0539d7dcc027"));
                //UserProfile up = mgr.GetUserProfile(loginName);
                string data = Utility.SerializeObject(up.Properties, 1);
                //r.AdminLinkUrl = up.PersonalUrl.ToString();
                //r.LinkUrl = up.PersonalUrl.ToString();
                r.ApplicationId = Guid.Parse(GDPR.Common.Configuration.ApplicationId);
                r.Object = data;
                r.Type = "UserProfile";
                r.RecordId = up.ID.ToString();
                r.RecordDate = up.PersonalSiteLastCreationTime;

                ValidateRecord(r);

                records.Add(r);
            }

            return records;
        }

        public override List<Record> GetAllRecords(GDPRSubject search)
        {
            List<Record> records = new List<Record>();

            SPSecurity.RunWithElevatedPrivileges(delegate ()
            {
                //do search
                records.AddRange(DoSharePointSearch(search));

                //do profile
                records.AddRange(DoProfileSearch(search));
            });

            return records;
        }

        public SPSite GetFirstSPSite()
        {
            SPFarm farm = SPFarm.Local;

            SPWebService service = farm.Services.GetValue<SPWebService>("");

            foreach (SPWebApplication webApp in service.WebApplications)
            {
                foreach (SPSite site in webApp.Sites)
                {
                    return site;
                }
            }

            return null;
        }

        public UserProfileManager GetUserProfileManager()
        {
            SPFarm farm = SPFarm.Local;

            SPWebService service = farm.Services.GetValue<SPWebService>("");

            foreach (SPWebApplication webApp in service.WebApplications)
            {
                foreach (SPSite site in webApp.Sites)
                {
                    //get user profiles...
                    SPServiceContext serviceContext = SPServiceContext.GetContext(site);

                    try
                    {
                        UserProfileManager userProfileMgr = new UserProfileManager(serviceContext);
                        return userProfileMgr;
                    }
                    catch (System.Exception e)
                    {
                        Console.WriteLine(e.GetType().ToString() + ": " + e.Message);
                        Console.Read();
                    }
                }
            }

            return null;
        }

        public override List<GDPRSubject> GetAllSubjects(int skip, int count, DateTime? changeDate)
        {
            List<GDPRSubject> subjects = new List<GDPRSubject>();

            SPFarm farm = SPFarm.Local;

            bool done = false; 

            SPWebService service = farm.Services.GetValue<SPWebService>("");

            foreach (SPWebApplication webApp in service.WebApplications)
            {
                foreach (SPSite site in webApp.Sites)
                {
                    if (done)
                        return subjects;

                    //get user profiles...
                    SPServiceContext serviceContext = SPServiceContext.GetContext(site);

                    try
                    {
                        UserProfileManager userProfileMgr = new UserProfileManager(serviceContext);

                        foreach (UserProfile profile in userProfileMgr)
                        {
                            GDPRSubject s = new GDPRSubject();
                            s.Attributes = new System.Collections.Hashtable();
                            s.ApplicationSubjectId = profile.ID.ToString();
                            
                            foreach(var prop in profile.Properties)
                            {
                                UserProfileValueCollection val = profile[prop.Name];

                                if (val == null || val.Value == null)
                                    continue;

                                s.Attributes.Add(prop.Name, val.Value);

                                switch(prop.Name)
                                {
                                    case "FirstName":
                                        s.FirstName = val.Value.ToString();
                                        break;
                                    case "LastName":
                                        s.LastName = val.Value.ToString();
                                        break;
                                    case "AboutMe":
                                        break;
                                    case "WorkPhone":
                                    case "MobilePhone":
                                    case "HomePhone":
                                        BasePhone p = BasePhone.Parse(val.Value.ToString());

                                        if (p != null)
                                        {
                                            GDPRSubjectPhone sp = new GDPRSubjectPhone();
                                            sp.Raw = p.ToString();
                                            s.Phones.Add(sp);
                                        }
                                        break;
                                    case "WorkEmail":
                                        s.EmailAddresses.Add(new GDPRSubjectEmail { EmailAddress = val.Value.ToString() });
                                        break;
                                    case "OfficeLocation":
                                        string address = val.Value.ToString();
                                        BaseAddress a = core.GeocodeAddress(null, address);
                                        s.Addresses.Add(new GDPRSubjectAddress() { Raw = a.ToString() });
                                        break;
                                    case "Birthday":
                                        s.BirthDate = DateTime.Parse(val.Value.ToString());
                                        break;
                                }
                            }

                            subjects.Add(s);
                        }

                        done = true;
                    }
                    catch (System.Exception e)
                    {
                        Console.WriteLine(e.GetType().ToString() + ": " + e.Message);
                        Console.Read();
                    }
                }
            }

            return subjects;
        }
    }
}
