using GDPR.Common;
using GDPR.Common.Classes;
using Microsoft.Office.Server.UserProfiles;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using System;
using System.Collections.Generic;

namespace GDPR.Applications
{
    public class SharePoint : BaseGDPROAuthApplication
    {
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
