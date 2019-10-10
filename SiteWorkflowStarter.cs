using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Security;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Workflow;
using Microsoft.SharePoint.Client.WorkflowServices;

namespace SiteWorkflowStarter
{
    public static class SiteWorkflowStarter
    {
        [FunctionName("SiteWorkflowStarter")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            try
            {
                log.Info("C# HTTP trigger function processed a request.");
                // Get request body
                bool status = false;
                Workflow data = await req.Content.ReadAsAsync<Workflow>();
                if (data.Key == Reverse(data.WorkflowName))
                {
                    status = TriggerWorkflow(data.SiteURL, data.WorkflowName);
                }
                if (status == true)
                    return req.CreateResponse(HttpStatusCode.OK, "The workflow is trigger successfully");
                else
                    return req.CreateResponse(HttpStatusCode.BadRequest, "Pass the correct parameters inorder to trigger a workflow");
            }
            catch (Exception ex)
            {
                string[] error = { ex.Message, ex.StackTrace, ex.Source };
                string errormsg = string.Format("The Workflow is failed to start, the error detail follows. Source :{0}, Message : {1}, Stacktrace:{2}", error[2], error[0], error[1]);
                log.Info(ex.Message);
                return req.CreateResponse(HttpStatusCode.BadRequest, errormsg);
                throw;
            }
        }


        public static bool TriggerWorkflow(string SiteURL, string WorkflowName)
        {
            bool status = false;
            try
            {
                string siteUrl = SiteURL;
                string userName = "EMAIlID OF THE USER ACCOUNT";
                string password = "PASSWORD OF THE ACCOUNT";

                //Name of the SharePoint 2010 Workflow to start.
                string workflowName = WorkflowName;

                using (ClientContext clientContext = new ClientContext(siteUrl))
                {
                    SecureString securePassword = new SecureString();

                    foreach (char c in password.ToCharArray()) securePassword.AppendChar(c);

                    clientContext.Credentials = new SharePointOnlineCredentials(userName, securePassword);

                    Web web = clientContext.Web;

                    //Workflow Services Manager which will handle all the workflow interaction.
                    WorkflowServicesManager wfServicesManager = new WorkflowServicesManager(clientContext, web);

                    //Will return all Workflow Associations which are running on the SharePoint 2010 Engine
                    WorkflowAssociationCollection wfAssociations = web.WorkflowAssociations;

                    //Get the required Workflow Association
                    WorkflowAssociation wfAssociation = wfAssociations.GetByName(workflowName);

                    clientContext.Load(wfAssociation);

                    clientContext.ExecuteQuery();

                    //Get the instance of the Interop Service which will be used to create an instance of the Workflow
                    InteropService workflowInteropService = wfServicesManager.GetWorkflowInteropService();

                    var initiationData = new Dictionary<string, object>();

                    //Start the Workflow
                    ClientResult<Guid> resultGuid = workflowInteropService.StartWorkflow(wfAssociation.Name, new Guid(), Guid.Empty, Guid.Empty, initiationData);

                    clientContext.ExecuteQuery();
                    status = true;

                }
            }
            catch (Exception ex)
            {
                status = false;
                throw;
            }
            return status;
        }

        public static string Reverse(string key)
        {
            char[] charArray = key.ToCharArray();
            Array.Reverse(charArray);
            return new string(charArray);
        }
    }
}
