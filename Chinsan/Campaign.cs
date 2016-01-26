using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Chinsan
{
    class Campaign
    {
        public String campaign;
        public ReportEntry yahooBranding;
        public ReportEntry yahooGeneral;
        public ReportEntry googleBranding;
        public ReportEntry googleGeneral;
        public DateTime reportStartDate;
        public DateTime reportEndDate;
        public Campaign(String campaign) {
            this.campaign = campaign;
        this.yahooBranding = new ReportEntry("Yahoo","社名");
        this.yahooGeneral = new ReportEntry("Yahoo", "一般");
        this.googleBranding = new ReportEntry("Google", "社名");
        this.googleGeneral = new ReportEntry("Google", "一般");
        }
    }
}
