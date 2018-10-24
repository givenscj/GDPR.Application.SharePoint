using GDPR.Util;
using GDPR.Util.Classes;
using GDPR.Util.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        
    }
}
