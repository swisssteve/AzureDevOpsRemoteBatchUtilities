using System;
using System.IO;

namespace AzureWriter
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length > 0 && File.Exists(args[0]))
            {
                VerifyExcelFile verifyFile = new VerifyExcelFile();
                if (verifyFile.VerifyFile(args[0])){

                    // Process the excel data into a format ready to be executed against Azure DevOps Server
                    ProcessExcelDataSheet excelSheet = new ProcessExcelDataSheet();
                    AzureDevOpsDao workItemData = excelSheet.GetExcelData(args[0]);

                    // Process the formatted Excel Data with the Azure DevOps Server
                    AzureDevOpsApi workItems = new AzureDevOpsApi(workItemData.ServerPath, workItemData.Pat, workItemData.Project);
                    workItems.ProcessWorkItems(workItemData);
                }
            }
            else 
            { 
                Console.WriteLine("Unable to locate the Azure Excel file to process");
            }           
        }
    }
}

