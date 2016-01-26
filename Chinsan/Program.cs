using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Aspose.Cells;

namespace Chinsan
{
    class Program
    {
        static void Main(string[] args)
        {

            License asposeLicense = new License();
            asposeLicense.SetLicense("Aspose.Cells.lic");
        

            //TODO: Google Campaign Betsu Report
            //TODO: AVG Rank
            List<Campaign> campaigns = new List<Campaign>();
            //Console.WriteLine("Chinsan Tool");
            String brandingFilter = "社名";

            String inputFolder = "..\\..\\input";

            String [] directories = Directory.GetDirectories(inputFolder);
            foreach (String directory in directories)
            {
                String [] splitPath=Path.GetFullPath(directory).Split('\\');
                String campaignName = splitPath[splitPath.Length - 1];
                campaigns.Add(new Campaign(campaignName));
            }

            Console.WriteLine(String.Format("====Writing Report for {0} Campaigns====",campaigns.Count));

            foreach (Campaign campaign in campaigns)
            {
                Console.WriteLine(campaign.campaign);
            }

            Console.WriteLine("======================================\n\n");

            foreach (Campaign campaign in campaigns) {
                Campaign currentCampaign = campaign;
                Console.WriteLine(String.Format("Processing Campaign {0}",campaign,campaign));
                String fullPath = String.Concat(inputFolder, "\\", campaign.campaign);
                Console.WriteLine(String.Format("Checking Folder {0} for Input Files", fullPath));
                String [] inputFiles = Directory.GetFiles(fullPath);
                Console.WriteLine(String.Format("{0} Files Found", inputFiles.Count()));
                foreach (String file in inputFiles) {
                    Console.WriteLine(file);
                }

                foreach (String file in inputFiles)
                {

                    if (file.Contains("CV"))
                        {
                            Utilities.parseCVReport(file, currentCampaign);
                        }
                        else
                        {
                            Utilities.parseCampaignReport(file, currentCampaign);
                       }

                       
                }

            }

            Utilities.writeReport(campaigns);
            Console.ReadLine();
        }
    }
}
