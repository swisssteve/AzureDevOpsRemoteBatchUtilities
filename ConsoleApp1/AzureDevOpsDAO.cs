using System;
using System.Collections;

namespace AzureWriter
{
    public class AzureDevOpsDao
    {
        public string ServerPath { get; set; }
        public string Pat { get; set; }
        public string Project { get; set; }
        public String AreaPath { get; set; }
        public ArrayList AzureData { get; set; } = new();
    }
}
