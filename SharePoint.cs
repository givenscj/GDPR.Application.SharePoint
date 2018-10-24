using GDPR.Applications;
using GDPR.Common.Classes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GDPR.Util;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.Office.Server.UserProfiles;

namespace GDPR.Applications
{
    public class SharePoint : BaseGDPROAuthApplication
    {
        public SharePoint()
        {
            this._version = "1.0.0.0";

            this._supportsEmailSearch = true;
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

            using (SPSite site = new SPSite("http://servername"))
            {
                SPServiceContext serviceContext = SPServiceContext.GetContext(site);

                try
                {
                    UserProfileManager userProfileMgr = new UserProfileManager(serviceContext);

                    foreach (UserProfile profile in userProfileMgr)
                    {
                        GDPRSubject s = new GDPRSubject();

                        subjects.Add(s);
                    }
                }

                catch (System.Exception e)
                {
                    Console.WriteLine(e.GetType().ToString() + ": " + e.Message);
                    Console.Read();
                }
            }

            return subjects;
        }

    }
}
