
using System;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;

namespace AzureWriter
{
    public class AzureDevOpsApi
    {
        readonly WorkItemTrackingHttpClient _workItemTrackingHttpClient;
        readonly string _project;

        /// <summary>
        /// Constructor. Set Azure DevOps Access Parameters.   This is project specific 
        /// </summary>
        public AzureDevOpsApi(string uri, string personalAccessToken, string azureProject)
        {
            _project = azureProject;
            var connection = new VssConnection(new Uri(uri), new VssBasicCredential("", personalAccessToken));
            _workItemTrackingHttpClient = connection.GetClient<WorkItemTrackingHttpClient>();
        }

        public void ProcessWorkItems(AzureDevOpsDao dat)
        {
            foreach (DevOpsEntryDao doe in dat.AzureData)
            {
                switch (doe.Operation)
                {
                   case "Create":
                   {
                       try
                       {
                           WorkItem result = _workItemTrackingHttpClient.CreateWorkItemAsync(doe.PatchDocument, dat.Project, doe.WorkItemType).Result;                            
                           Console.WriteLine("Item Successfully Created: WorkItem #{0}", result.Id);
                       }
                       catch (AggregateException ex)
                       {
                           if (ex.InnerException != null)
                               Console.WriteLine("Error creating WorkItem: #{0}", ex.InnerException.Message);
                       }
                       break;
                   }

                   case "Update":
                   {
                       try
                       {
                           WorkItem result = _workItemTrackingHttpClient.UpdateWorkItemAsync(doe.PatchDocument, doe.WorkItemId).Result;
                           Console.WriteLine("Item Successfully Updated: Item #{0}", result.Id);
                       }
                       catch (AggregateException ex)
                       {
                           if (ex.InnerException != null)
                               Console.WriteLine("Error Updating Work item: #{0}", ex.InnerException.Message);
                       }
                       break;
                   }

                   case "Delete":
                   {
                       try
                       {
                           var result = _workItemTrackingHttpClient.DeleteWorkItemAsync(_project, doe.WorkItemId, false).Result;
                           Console.WriteLine("Item Successfully Deleted: WorkItem #{0}", result.Id);
                       }
                       catch (AggregateException ex)
                       {
                           if (ex.InnerException != null)
                               Console.WriteLine("Error Deleting Item: #{0}", ex.InnerException.Message);
                       }
                       break;
                   }

                   default:
                   {
                       break;
                   }
                }
            }

        }
    }
}
