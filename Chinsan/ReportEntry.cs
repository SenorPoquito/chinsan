using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Chinsan
{
    class ReportEntry
    {
        public String media;
        public String campaignType;
        public Double weightedRank;
        public Double impressions;
        public Double clickThroughs;
        public Double cost;
        public Double conversionUnique;
        public Double conversionBrochure;
        public Double conversionBooking;
        
        public ReportEntry(String media, String campaignType) {
            this.media = media;
            this.campaignType = campaignType;
            this.weightedRank = 0.0;
            this.impressions = 0.0;
        
            this.clickThroughs = 0.0;
            this.cost = 0.0;
            this.conversionBooking = 0.0;
            this.conversionUnique = 0.0;
            this.conversionBrochure = 0.0;
        
        }

    }
}
