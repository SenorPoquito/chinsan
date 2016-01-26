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

        private const String YCR_CAMPAIGN_NAME = "キャンペーン名";
        private const String YCR_IMPRESSIONS = "インプレッション数";
        private const String YCR_CLICKS = "クリック数";
        private const String YCR_AVG_RANK = "平均掲載順位";
        private const String YCR_COST = "コスト";
        private const String YCR_UNIQUE_CV = "ユニークコンバージョン数";

        private const String YCV_CV_NAME = "コンバージョン名";
        private const String YCV_CVs = "ユニークコンバージョン数";
        private const String YCV_CAMPAIGN_NAME = "キャンペーン名";

        private const String GCR_CAMPAIGN_NAME = "キャンペーン";
        private const String GCR_AVG_RANK = "平均掲載順位";
        private const String GCR_IMPRESSIONS = "表示回数";
        private const String GCR_CLICKS = "クリック数";
        private const String GCR_COST = "費用";
        private const String GCR_UNIQUE_CVs = "コンバージョンに至ったクリック";

        private const String GCV_CV_NAME = "コンバージョン名";
        private const String GCV_CVs = "コンバージョン";
        private const String GCV_CAMPAIGN_NAME = "キャンペーン";

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


        private static int cprColumnCampaign = 0;
        private static int cprColumnImpressions = 0;
        private static int cprColumnClicks = 0;
        private static int cprColumnAvgRank = 0;
        private static int cprColumnCost = 0;
        private static int cprColumnUniqueCV = 0;
        private static int cvrColumnCV = 0;
        private static int cvrColumnCampaign = 0;
        private static int cvrColumnConverstionType = 0;


        public static void parseCVReport(String path, Campaign campaign)
        {

            TextFieldParser parser = new TextFieldParser(path, Encoding.GetEncoding("shift_jis"));
            parser.TextFieldType = FieldType.Delimited;
            parser.SetDelimiters(new String[] { ",", "\t" });

            Boolean google = false;
            Boolean ready = false;
            int row = 0;

            if (path.Contains("Yahoo"))
            {
                google = false;
            }
            else {
                google = true;
            }
           
            while (!parser.EndOfData)
            {
                //Process row
                int column = 0;
                Boolean branding = false;
                Boolean document = false;
                string[] fields = parser.ReadFields();

                

                foreach (string field in fields)
                {
                    if (!ready)
                    {
                        if (field.Equals(YCV_CAMPAIGN_NAME) || field.Equals(GCV_CAMPAIGN_NAME))
                        {
                            cvrColumnCampaign = column;
                        }
                        if (field.Equals(YCV_CVs) || field.Equals(GCV_CVs))
                        {
                            cvrColumnCV = column;
                        }
                        if (field.Equals(YCV_CV_NAME) || field.Equals(GCV_CV_NAME))
                        {
                            cvrColumnConverstionType = column;
                        }

                        if (row > 5 || (google && row>2))
                        {
                            ready = true;
                        }
                    }

                    if (ready)
                    {

                        if (column == cvrColumnConverstionType && field.Contains("資料"))
                        {
                                document = true;
                        }

                        if (column == cvrColumnCampaign)
                        {
                            if (field.Contains("社名"))
                            {
                                branding = true;
                            }
                            if (field.Contains("--"))
                            {
                                return;
                            }
                        }

                        if (column == cvrColumnCV)
                        {
                            if (branding)
                            {
                                if (document)
                                {
                                    if (google)
                                    {
                                        campaign.googleBranding.conversionBrochure += Convert.ToDouble(field);
                                    }
                                    else {
                                        campaign.yahooBranding.conversionBrochure += Convert.ToDouble(field);
                                    }
                                }
                                else
                                {
                                    if (google)
                                    {
                                        campaign.googleBranding.conversionBooking += Convert.ToDouble(field);
                                    }
                                    else {
                                        campaign.yahooBranding.conversionBooking += Convert.ToDouble(field);
                                    }
                                }
                            }
                            else
                            {
                                if (document)
                                {
                                    if (google)
                                    {
                                        campaign.googleGeneral.conversionBrochure += Convert.ToDouble(field);
                                    }
                                    else {
                                        campaign.yahooGeneral.conversionBrochure += Convert.ToDouble(field);
                                    }
                                }
                                else
                                {
                                    if (google)
                                    {
                                        campaign.googleGeneral.conversionBooking += Convert.ToDouble(field);
                                    }
                                    else {
                                        campaign.yahooGeneral.conversionBooking += Convert.ToDouble(field);
                                    }
                                }
                            }
                        }

                    }
                    //TODO: Process field
                    column++;
                }
                row++;
            }
            parser.Close();

            return;
        }


        public static void parseCampaignReport(String path, Campaign campaign)
        {

            TextFieldParser parser = new TextFieldParser(path, Encoding.GetEncoding("shift_jis"));
            parser.TextFieldType = FieldType.Delimited;
            parser.SetDelimiters(new String[] { ",", "\t" });
       
            Boolean ready = false;
            int row = 0;
            Boolean google = false;

            if (path.Contains("Yahoo"))
            {
                google = false;
            }
            else {
                google = true;
            }

            while (!parser.EndOfData)
            {
                //Process row
                int column = 0;
                Boolean branding = false;
                Double impression = 0.0;
                Double averageRank = 0.0;
                string[] fields = parser.ReadFields();
                foreach (string field in fields)
                {
                  
                    if (field.Equals(YCR_AVG_RANK) || field.Equals(GCR_AVG_RANK))
                    {
                        cprColumnAvgRank = column;
                    }
                    if (field.Equals(YCR_CAMPAIGN_NAME) || field.Equals(GCR_CAMPAIGN_NAME))
                    {
                        cprColumnCampaign = column;
                    }
                    if (field.Equals(YCR_CLICKS) || field.Equals(GCR_CLICKS))
                    {
                        cprColumnClicks = column;
                    }
                    if (field.Equals(YCR_COST) || field.Equals(GCR_COST))
                    {
                        cprColumnCost = column;
                    }
                    if (field.Equals(YCR_IMPRESSIONS) || field.Equals(GCR_IMPRESSIONS))
                    {
                        cprColumnImpressions = column;
                    }
                    if (field.Equals(YCR_UNIQUE_CV) || field.Equals(GCR_UNIQUE_CVs))
                    {
                        cprColumnUniqueCV = column;
                    }

                    if (field.Equals("ユニークコンバージョン率") || field.Equals("すべてのコンバージョン"))
                    {
                        ready = true;
                    }

                    if (ready)
                    {
                        if (column == cprColumnAvgRank)
                        {
                            averageRank = Convert.ToDouble(field);
                        }

                        if (column == cprColumnImpressions)
                        {
                            impression = Convert.ToDouble(field);
                        }

                        if (column == cprColumnCampaign)
                        {
                            if (field.Contains("社名"))
                            {
                                branding = true;
                            }
                            if (field.Contains("--"))
                            {
                                return;
                            }
                        }
                        if (column == cprColumnClicks)
                        {
                            if (branding)
                            {
                                if (google)
                                {
                                    campaign.googleBranding.clickThroughs += Convert.ToDouble(field);
                                }
                                else {
                                    campaign.yahooBranding.clickThroughs += Convert.ToDouble(field);
                                }

                                
                            }
                            else
                            {
                                if (google)
                                {
                                    campaign.googleGeneral.clickThroughs += Convert.ToDouble(field);
                                }
                                else {
                                    campaign.yahooGeneral.clickThroughs += Convert.ToDouble(field);
                                }
                            }
                        }

                        if (column == cprColumnCost)
                        {
                            if (branding)
                            {
                                if (google)
                                {
                                    campaign.googleBranding.cost += Convert.ToDouble(field);
                                }
                                else {
                                    campaign.yahooBranding.cost += Convert.ToDouble(field);
                                }
                            }
                            else
                            {
                                if (google)
                                {
                                    campaign.googleGeneral.cost += Convert.ToDouble(field);
                                }
                                else {
                                    campaign.yahooGeneral.cost += Convert.ToDouble(field);
                                }
                            }
                        }

                        if (column == cprColumnUniqueCV)
                        {
                            if (branding)
                            {
                                if (google)
                                {
                                    campaign.googleBranding.conversionUnique += Convert.ToDouble(field);
                                }
                                else {
                                    campaign.yahooBranding.conversionUnique += Convert.ToDouble(field);
                                }
                            }
                            else
                            {
                                if (google)
                                {
                                    campaign.googleGeneral.conversionUnique += Convert.ToDouble(field);
                                }
                                else {
                                    campaign.yahooGeneral.conversionUnique += Convert.ToDouble(field);
                                }
                            }
                        }

                     
                    }
                    //TODO: Process field
                    column++;
                }
                Double weightedRank = averageRank * impression;
                if (branding)
                {
                    if (google)
                    {
                        campaign.googleBranding.weightedRank += weightedRank;
                    }
                    else {
                        campaign.yahooBranding.weightedRank+= weightedRank;
                    }
                }
                else
                {
                    if (google)
                    {
                        campaign.googleGeneral.weightedRank += weightedRank;
                    }
                    else {
                        campaign.yahooGeneral.weightedRank += weightedRank;
                    }
                }


                row++;

            }
            parser.Close();

            return;
        }








        public static void writeReport(List<Campaign> campaigns)
        {
            DateTime reportDate = dateEnd.AddDays(1);
            String reportDateString = reportDate.ToString("yyyMMdd");
            String reportMonth = reportDate.Year + "年" + reportDate.Month + "月";
            String filename = reportDateString + "_株式会社一蔵御中_" + reportMonth + "配信レポート.xlsx";
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
                cell = activeSheet.Cells[OP_IMP + selectedRow];
                cell.Value = campaign.googleBranding.impressions;
                //Clicks
                cell = activeSheet.Cells[OP_CTs + selectedRow];
                cell.Value = campaign.googleBranding.clickThroughs;
                //CTR
                cell = activeSheet.Cells[OP_CTR + selectedRow];
                cell.Value = campaign.googleBranding.clickThroughs / campaign.googleBranding.impressions;
                //CPC
                cell = activeSheet.Cells[OP_CPC + selectedRow];
                cell.Value = campaign.googleBranding.cost / campaign.googleBranding.clickThroughs;
                //Cost
                cell = activeSheet.Cells[OP_COST + selectedRow];
                cell.Value = campaign.googleBranding.cost;
                //Cost+20%
                cell = activeSheet.Cells[OP_COST_FEE + selectedRow];
                cell.Value = campaign.googleBranding.cost * 1.20;
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
                cell.Value = campaign.googleBranding.conversionUnique / campaign.googleBranding.clickThroughs;
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
                cell.Value = "◆Date：" + dateStart.ToString("yyyy/MM/dd") + "-" + dateEnd.AddDays(1).ToString("yyyy/MM/dd");

            }

            workbook.Save(filename);
        }

    }
}