using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualBasic.FileIO;
using System.IO;
using Aspose.Cells;

namespace Chinsan
{
    class Utilities
    {
        public static DateTime dateStart = new DateTime();
        public static DateTime dateEnd = new DateTime();
        private const int YCR_CAMPAIGN_NAME = 0;
        private const int YCR_CAMPAIGN_TYPE = 1;
        private const int YCR_IMPRESSIONS = 2;
        private const int YCR_CLICKS = 3;
        private const int YCR_CTR = 4;
        private const int YCR_AVG_RANK = 5;
        private const int YCR_COST = 6;
        private const int YCR_AVG_CPC = 7;
        private const int YCR_COST_PER_UNIQUE_CV = 8;
        private const int YCR_UNIQUE_CV = 9;
        private const int YCR_UNIQUE_CVR = 10;

        private const int YCV_CV_NAME = 0;
        private const int YCV_CVs = 3;
        private const int YCV_CAMPAIGN_NAME = 1;

        private const int GCR_CAMPAIGN_NAME = 1;
        private const int GCR_AVG_RANK = 4;
        private const int GCR_IMPRESSIONS = 5;
        private const int GCR_CLICKS = 6;
        private const int GCR_COST = 9;
        private const int GCR_UNIQUE_CVs = 11; 

        private const int GCV_CV_NAME = 0;
        private const int GCV_CVs = 5;
        //private const int GCV_UNIQUE_CVs = 5;
        private const int GCV_CAMPAIGN_NAME = 2;

        private const String OP_DATE = "B7";
        private const String OP_AVG_RANK = "D";
        private const String OP_IMP = "E";
        private const String OP_CTs = "F";
        private const String OP_CTR = "G";
        private const String OP_CPC = "H";
        private const String OP_COST = "I";
        private const String OP_COST_FEE = "J";
        private const String OP_CV_UNIQUE = "K";
        private const String OP_CV_BOOKING = "L";
        private const String OP_CV_DOCUMENT = "M";
        private const String OP_CVR_UNIQUE = "N";
        private const String OP_CPA_UNIQUE = "O";

        private const String OP_YAHOO_BRANDING_ROW = "12";
        private const String OP_YAHOO_GENERAL_ROW = "13";
        private const String OP_GOOGLE_BRANDING_ROW = "14";
        private const String OP_GOOGLE_GENERAL_ROW = "15";
        private const String OP_TOTAL_BRANDING_ROW = "10";
        private const String OP_TOTAL_GENERAL_ROW = "11";

        public static void parseYahooCVReport(String path, Campaign campaign)
        {

            TextFieldParser parser = new TextFieldParser(path, Encoding.GetEncoding("shift_jis"));
            parser.TextFieldType = FieldType.Delimited;
            parser.SetDelimiters(",");

            Boolean ready = false;

            while (!parser.EndOfData)
            {
                //Process row
                int column = 0;
                Boolean branding = false;
                Boolean document = false;
                string[] fields = parser.ReadFields();
                foreach (string field in fields)
                {
                    if (field.Contains("売上/総コンバージョン数"))
                    {
                        ready = true;
                    }

                    if (ready)
                    {
                        switch (column)
                        {
                            case YCV_CV_NAME:
                                if (field.Contains("資料"))
                                {
                                    document = true;
                                }

                                break;
                                

                            case YCV_CAMPAIGN_NAME:
                                if (field.Contains("社名"))
                                {
                                    branding = true;
                                }
                                if (field.Contains("--"))
                                {
                                    return;
                                }
                                break;

                            case YCV_CVs:

                                if (branding)
                                {
                                    if (document)
                                    {
                                        campaign.yahooBranding.conversionBrochure += Convert.ToDouble(field);
                                    }
                                    else 
                                    {
                                        campaign.yahooBranding.conversionBooking += Convert.ToDouble(field);
                                    }
                                }
                                else 
                                {
                                    if (document)
                                    {
                                        campaign.yahooGeneral.conversionBrochure += Convert.ToDouble(field);
                                    }
                                    else 
                                    {
                                        campaign.yahooGeneral.conversionBooking += Convert.ToDouble(field);
                                    }
                                }

                                break;
                        }
                    }
                    //TODO: Process field
                    column++;
                }

            }
            parser.Close();

            return;
        }


        public static void parseYahooCampaignReport(String path, Campaign campaign) 
        {

            TextFieldParser parser = new TextFieldParser(path, Encoding.GetEncoding("shift_jis"));
            parser.TextFieldType = FieldType.Delimited;
            parser.SetDelimiters(",");

            Boolean ready = false;

            while (!parser.EndOfData)
            {
                //Process row
                int column = 0;
                Boolean branding = false;
                Double impression = 0.0;
                string[] fields = parser.ReadFields();
                    foreach (string field in fields)
                    {
                        if(field.Contains("ユニークコンバージョン率")){
                            ready=true;
                        }

                        if (ready)
                        {
                            switch(column){
                                case YCR_CAMPAIGN_NAME:
                                    if (field.Contains("社名")) {
                                        branding = true;
                                    }
                                    if (field.Contains("--")) {
                                        return;
                                    }
                                    break;
                                case YCR_AVG_RANK:
                                      if (branding)
                                    {
                                        campaign.yahooBranding.weightedRank += Convert.ToDouble(field) * impression;
                                    }
                                    else
                                    {
                                        campaign.yahooGeneral.weightedRank += Convert.ToDouble(field) * impression;
                                    }
                                    break;
                                    
                                    break;
                               
                                case YCR_IMPRESSIONS:
                                    impression = Convert.ToDouble(field);
                                    if (branding)
                                    {
                                        
                                        campaign.yahooBranding.impressions += Convert.ToDouble(field);
                                    }
                                    else {
                                        
                                        campaign.yahooGeneral.impressions += Convert.ToDouble(field);
                                    }
                                    break;
                                case YCR_CLICKS:
                                    if (branding)
                                    {
                                        campaign.yahooBranding.clickThroughs += Convert.ToDouble(field);
                                    }
                                    else
                                    {
                                        campaign.yahooGeneral.clickThroughs += Convert.ToDouble(field);
                                    }
                                    break;
                                case YCR_COST:
                                    if (branding)
                                    {
                                        campaign.yahooBranding.cost += Convert.ToDouble(field);
                                    }
                                    else
                                    {
                                        campaign.yahooGeneral.cost += Convert.ToDouble(field);
                                    }
                                    break;

                                case YCR_UNIQUE_CV:
                                    if (branding)
                                    {
                                        campaign.yahooBranding.conversionUnique += Convert.ToDouble(field);
                                    }
                                    else
                                    {
                                        campaign.yahooGeneral.conversionUnique += Convert.ToDouble(field);
                                    }
                                    break;

                            }
                        }
                        //TODO: Process field
                        column++;
                    }
                
            }
            parser.Close();

            return;
        }


        public static void parseGoogleCampaignReport(String path, Campaign campaign)
        {

            TextFieldParser parser = new TextFieldParser(path, Encoding.GetEncoding("shift_jis"));
            parser.TextFieldType = FieldType.Delimited;
            parser.SetDelimiters("\t");

            Boolean ready = false;

            while (!parser.EndOfData)
            {
                //Process row
                int column = 0;
                Boolean branding = false;
                Double avgRank = 0.0;
                string[] fields = parser.ReadFields();
                foreach (string field in fields)
                {
                    if (field.Contains("キャンペーン レポート")) {
                        String [] dateSegment = field.Split(' ')[2].Split('(')[1].Split(')')[0].Split('-');
                        dateStart = Convert.ToDateTime(dateSegment[0]);
                        dateEnd = Convert.ToDateTime(dateSegment[1]);

                        Console.WriteLine(dateSegment);


                    }
                    if (field.Contains("すべてのコンバージョン"))
                    {
                        ready = true;
                    }

                    if (ready)
                    {
                        if (field.Contains("--"))
                        {
                            return;
                        }
                        switch (column)
                        {
                            case GCR_CAMPAIGN_NAME:
                                if (field.Contains("社名"))
                                {
                                    branding = true;
                                }
                              
                                break;
                            case GCR_AVG_RANK:
                                avgRank = Convert.ToDouble(field);
                                break;
                            case GCR_IMPRESSIONS:
                                if (branding)
                                {
                                    campaign.googleBranding.weightedRank += avgRank * Convert.ToDouble(field);
                                    campaign.googleBranding.impressions += Convert.ToDouble(field);
                                }
                                else
                                {
                                    campaign.googleGeneral.weightedRank += avgRank * Convert.ToDouble(field);
                                    campaign.googleGeneral.impressions += Convert.ToDouble(field);
                                }
                                break;
                            case GCR_CLICKS:
                                if (branding)
                                {
                                    campaign.googleBranding.clickThroughs += Convert.ToDouble(field);
                                }
                                else
                                {
                                    campaign.googleGeneral.clickThroughs += Convert.ToDouble(field);
                                }
                                break;
                            case GCR_COST:
                                if (branding)
                                {
                                    campaign.googleBranding.cost += Convert.ToDouble(field);
                                }
                                else
                                {
                                    campaign.googleGeneral.cost += Convert.ToDouble(field);
                                }
                                break;

                            case GCR_UNIQUE_CVs:
                                if (branding)
                                {
                                    campaign.googleBranding.conversionUnique += Convert.ToDouble(field);
                                }
                                else
                                {
                                    campaign.googleGeneral.conversionUnique += Convert.ToDouble(field);
                                }
                                break;

                        }
                    }
                    //TODO: Process field
                    column++;
                }

            }
            parser.Close();

            return;
        }


        public static void parseGoogleCVReport(String path, Campaign campaign)
        {

            TextFieldParser parser = new TextFieldParser(path, Encoding.GetEncoding("shift_jis"));
            parser.TextFieldType = FieldType.Delimited;
            parser.SetDelimiters("\t");

            Boolean ready = false;

            while (!parser.EndOfData)
            {
                //Process row
                int column = 0;
                Boolean branding = false;
                Boolean document = false;
                string[] fields = parser.ReadFields();
                foreach (string field in fields)
                {
                    if (field.Contains("すべてのコンバージョン"))
                    {
                        ready = true;
                    }

                    if (ready)
                    {
                        switch (column)
                        {
                            case GCV_CV_NAME:
                                if (field.Contains("資料"))
                                {
                                    document = true;
                                }

                                break;


                            case GCV_CAMPAIGN_NAME:
                                if (field.Contains("社名"))
                                {
                                    branding = true;
                                }
                                if (field.Contains("--"))
                                {
                                    return;
                                }
                                break;

                            case GCV_CVs:

                                if (branding)
                                {
                                    if (document)
                                    {
                                        campaign.googleBranding.conversionBrochure += Convert.ToDouble(field);
                                    }
                                    else
                                    {
                                        campaign.googleBranding.conversionBooking += Convert.ToDouble(field);
                                    }
                                }
                                else
                                {
                                    if (document)
                                    {
                                        campaign.googleGeneral.conversionBrochure += Convert.ToDouble(field);
                                    }
                                    else
                                    {
                                        campaign.googleGeneral.conversionBooking += Convert.ToDouble(field);
                                    }
                                }

                                break;

                        }
                    }
                    //TODO: Process field
                    column++;
                }

            }
            parser.Close();

            return;
        }


            

        public static void writeReport(List<Campaign> campaigns) {
            DateTime reportDate = dateEnd.AddDays(1);
            String reportDateString = reportDate.ToString("yyyMMdd");
            String reportMonth = reportDate.Year+"年"+reportDate.Month+"月";
            String filename = reportDateString+"_株式会社一蔵御中_"+reportMonth+"配信レポート.xlsx";
            Console.WriteLine(filename);
           
            Workbook workbook = new Workbook("..\\..\\input\\template.xlsx");

            foreach (Campaign campaign in campaigns)
            {
                Worksheet activeSheet = workbook.Worksheets[campaign.campaign];
                
                String selectedRow = OP_GOOGLE_BRANDING_ROW;
                //GOOGLE BRANDING
                //Average Rank
                Cell cell = activeSheet.Cells[OP_AVG_RANK + selectedRow];
                cell.Value = campaign.googleBranding.weightedRank / campaign.googleBranding.impressions;
                //Impressions
                cell = activeSheet.Cells[OP_IMP+selectedRow];
                cell.Value = campaign.googleBranding.impressions;
                //Clicks
                cell = activeSheet.Cells[ OP_CTs+selectedRow];
                cell.Value = campaign.googleBranding.clickThroughs;
                //CTR
                cell = activeSheet.Cells[ OP_CTR+selectedRow];
                cell.Value = campaign.googleBranding.clickThroughs/campaign.googleBranding.impressions;
                //CPC
                cell = activeSheet.Cells[ OP_CPC+selectedRow];
                cell.Value = campaign.googleBranding.cost / campaign.googleBranding.clickThroughs;
                //Cost
                cell = activeSheet.Cells[ OP_COST+selectedRow];
                cell.Value = campaign.googleBranding.cost;
                //Cost+20%
                cell = activeSheet.Cells[ OP_COST_FEE+selectedRow];
                cell.Value = campaign.googleBranding.cost*1.20;
                //CV Unique
                cell = activeSheet.Cells[OP_CV_UNIQUE + selectedRow];
                cell.Value = campaign.googleBranding.conversionUnique;
                //CV Brochure
                cell = activeSheet.Cells[OP_CV_DOCUMENT + selectedRow];
                cell.Value = campaign.googleBranding.conversionBrochure;
                //CV Booking
                cell = activeSheet.Cells[OP_CV_BOOKING + selectedRow];
                cell.Value = campaign.googleBranding.conversionBooking;
                //CVR Unique
                cell = activeSheet.Cells[OP_CVR_UNIQUE + selectedRow];
                cell.Value = campaign.googleBranding.conversionUnique/campaign.googleBranding.clickThroughs;
                //CPA Unique
                cell = activeSheet.Cells[OP_CPA_UNIQUE + selectedRow];
                cell.Value = campaign.googleBranding.cost / campaign.googleBranding.conversionUnique;

                selectedRow = OP_GOOGLE_GENERAL_ROW;
                //GOOGLE GENERAL
                //Average Rank
                cell = activeSheet.Cells[OP_AVG_RANK + selectedRow];
                cell.Value = campaign.googleGeneral.weightedRank / campaign.googleGeneral.impressions;
                //Impressions
                cell = activeSheet.Cells[OP_IMP + selectedRow];
                cell.Value = campaign.googleGeneral.impressions;
                //Clicks
                cell = activeSheet.Cells[OP_CTs + selectedRow];
                cell.Value = campaign.googleGeneral.clickThroughs;
                //CTR
                cell = activeSheet.Cells[OP_CTR + selectedRow];
                cell.Value = campaign.googleGeneral.clickThroughs / campaign.googleGeneral.impressions;
                //CPC
                cell = activeSheet.Cells[OP_CPC + selectedRow];
                cell.Value = campaign.googleGeneral.cost / campaign.googleGeneral.clickThroughs;
                //Cost
                cell = activeSheet.Cells[OP_COST + selectedRow];
                cell.Value = campaign.googleGeneral.cost;
                //Cost+20%
                cell = activeSheet.Cells[OP_COST_FEE + selectedRow];
                cell.Value = campaign.googleGeneral.cost * 1.20;
                //CV Unique
                cell = activeSheet.Cells[OP_CV_UNIQUE + selectedRow];
                cell.Value = campaign.googleGeneral.conversionUnique;
                //CV Brochure
                cell = activeSheet.Cells[OP_CV_DOCUMENT + selectedRow];
                cell.Value = campaign.googleGeneral.conversionBrochure;
                //CV Booking
                cell = activeSheet.Cells[OP_CV_BOOKING + selectedRow];
                cell.Value = campaign.googleGeneral.conversionBooking;
                //CVR Unique
                cell = activeSheet.Cells[OP_CVR_UNIQUE + selectedRow];
                cell.Value = campaign.googleGeneral.conversionUnique / campaign.googleGeneral.clickThroughs;
                //CPA Unique
                cell = activeSheet.Cells[OP_CPA_UNIQUE + selectedRow];
                cell.Value = campaign.googleGeneral.cost / campaign.googleGeneral.conversionUnique;

                selectedRow = OP_YAHOO_GENERAL_ROW;
                //YAHOO GENERAL
                //Average Rank
                cell = activeSheet.Cells[OP_AVG_RANK + selectedRow];
                cell.Value = campaign.yahooGeneral.weightedRank / campaign.yahooGeneral.impressions;
                //Impressions
                cell = activeSheet.Cells[OP_IMP + selectedRow];
                cell.Value = campaign.yahooGeneral.impressions;
                //Clicks
                cell = activeSheet.Cells[OP_CTs + selectedRow];
                cell.Value = campaign.yahooGeneral.clickThroughs;
                //CTR
                cell = activeSheet.Cells[OP_CTR + selectedRow];
                cell.Value = campaign.yahooGeneral.clickThroughs / campaign.yahooGeneral.impressions;
                //CPC
                cell = activeSheet.Cells[OP_CPC + selectedRow];
                cell.Value = campaign.yahooGeneral.cost / campaign.yahooGeneral.clickThroughs;
                //Cost
                cell = activeSheet.Cells[OP_COST + selectedRow];
                cell.Value = campaign.yahooGeneral.cost;
                //Cost+20%
                cell = activeSheet.Cells[OP_COST_FEE + selectedRow];
                cell.Value = campaign.yahooGeneral.cost * 1.20;
                //CV Unique
                cell = activeSheet.Cells[OP_CV_UNIQUE + selectedRow];
                cell.Value = campaign.yahooGeneral.conversionUnique;
                //CV Brochure
                cell = activeSheet.Cells[OP_CV_DOCUMENT + selectedRow];
                cell.Value = campaign.yahooGeneral.conversionBrochure;
                //CV Booking
                cell = activeSheet.Cells[OP_CV_BOOKING + selectedRow];
                cell.Value = campaign.yahooGeneral.conversionBooking;
                //CVR Unique
                cell = activeSheet.Cells[OP_CVR_UNIQUE + selectedRow];
                cell.Value = campaign.yahooGeneral.conversionUnique / campaign.yahooGeneral.clickThroughs;
                //CPA Unique
                cell = activeSheet.Cells[OP_CPA_UNIQUE + selectedRow];
                cell.Value = campaign.yahooGeneral.cost / campaign.yahooGeneral.conversionUnique;

                selectedRow = OP_YAHOO_BRANDING_ROW;
                //YAHOO BRANDING
                //Average Rank
                cell = activeSheet.Cells[OP_AVG_RANK + selectedRow];
                cell.Value = campaign.yahooBranding.weightedRank / campaign.yahooBranding.impressions;
                //Impressions
                cell = activeSheet.Cells[OP_IMP + selectedRow];
                cell.Value = campaign.yahooBranding.impressions;
                //Clicks
                cell = activeSheet.Cells[OP_CTs + selectedRow];
                cell.Value = campaign.yahooBranding.clickThroughs;
                //CTR
                cell = activeSheet.Cells[OP_CTR + selectedRow];
                cell.Value = campaign.yahooBranding.clickThroughs / campaign.yahooBranding.impressions;
                //CPC
                cell = activeSheet.Cells[OP_CPC + selectedRow];
                cell.Value = campaign.yahooBranding.cost / campaign.yahooBranding.clickThroughs;
                //Cost
                cell = activeSheet.Cells[OP_COST + selectedRow];
                cell.Value = campaign.yahooBranding.cost;
                //Cost+20%
                cell = activeSheet.Cells[OP_COST_FEE + selectedRow];
                cell.Value = campaign.yahooBranding.cost * 1.20;
                //CV Unique
                cell = activeSheet.Cells[OP_CV_UNIQUE + selectedRow];
                cell.Value = campaign.yahooBranding.conversionUnique;
                //CV Brochure
                cell = activeSheet.Cells[OP_CV_DOCUMENT + selectedRow];
                cell.Value = campaign.yahooBranding.conversionBrochure;
                //CV Booking
                cell = activeSheet.Cells[OP_CV_BOOKING + selectedRow];
                cell.Value = campaign.yahooBranding.conversionBooking;
                //CVR Unique
                cell = activeSheet.Cells[OP_CVR_UNIQUE + selectedRow];
                cell.Value = campaign.yahooBranding.conversionUnique / campaign.yahooBranding.clickThroughs;
                //CPA Unique
                cell = activeSheet.Cells[OP_CPA_UNIQUE + selectedRow];
                cell.Value = campaign.yahooBranding.cost / campaign.yahooBranding.conversionUnique;


                //DATE

                cell = activeSheet.Cells[OP_DATE];
                cell.Value = "◆Date："+ dateStart.ToString("yyyy/MM/dd")+"-"+ dateEnd.AddDays(1).ToString("yyyy/MM/dd");

            }

            workbook.Save(filename);
        }

    }
}