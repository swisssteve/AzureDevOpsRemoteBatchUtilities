using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Formatting;
using System.Threading.Tasks;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;

namespace AzureWriter
{
    public class QueryExecutor
    {
        private readonly Uri _uri;
        private readonly string _personalAccessToken;
        private readonly string _project;
        public QueryExecutor(string serverLocation, string personalAccessToken, string workingProject)
        {
            _uri = new Uri(serverLocation);
            _personalAccessToken = personalAccessToken;
            _project = workingProject;
        }

        public async Task<IList<WorkItem>> Execute(string queryString, string [] fields)
        {
            var credentials = new VssBasicCredential(string.Empty, this._personalAccessToken);
            var wiql = new Wiql()
            {
              Query = queryString
            };

            using (var httpClient = new WorkItemTrackingHttpClient(this._uri, new VssCredentials(credentials)))
            {
                try
                {
                    var result = httpClient.QueryByWiqlAsync(wiql).Result;
                    var ids = result.WorkItems.Select(item => item.Id).ToArray();

                    if (ids.Length == 0)
                    {
                        return Array.Empty<WorkItem>();
                    }

                   var WI = httpClient.GetWorkItemsAsync(ids, fields, result.AsOf).Result;
                   
                    return WI;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error querying work items: " + ex.Message);
                    return Array.Empty<WorkItem>();
                }
            }
        }
    }
}


/* Example Queries
  
public async Task PrintOpenBugsAsync(string project)
    {
        var workItems = await this.QueryOpenBugs(project).ConfigureAwait(false);
        Console.WriteLine("Query Results: {0} items found", workItems.Count);

        foreach (var workItem in workItems)
        {
            Console.WriteLine(
                "{0}\t{1}\t{2}",
                workItem.Id,
                workItem.Fields["System.Title"],
                workItem.Fields["System.State"]);

        Query = "Select [Id] " +
                         "From WorkItems " +
                         "Where [Work Item Type] = 'Task' " +
                         "And [System.TeamProject] = '" + project + "' " +
                         "And [System.State] <> 'Closed' " +
                         "Order By [State] Asc, [Changed Date] Desc",

                  Query = "Select [Id] " +
                         "From WorkItems " +

  Query = $"SELECT [System.Id], [System.Title] FROM WorkItems WHERE [System.Tags] CONTAINS 'Tag2'"

        */
