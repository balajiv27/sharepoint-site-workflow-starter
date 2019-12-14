# SharePoint SiteWorkflow Starter
This project is built as an Azure Function and can be used to trigger SharePoint site workflows. It requires a valid windows account which has the required permission to the site. 

This Azure function can take any POST request and execute the Workflow trigger action based on the inputs. 
The POST request requires the required JSON body, 

{ 
   "SiteURL":"Give the Site Url",
   "WorkflowName":"Give workflow name",
   "Key":"This is can be workflow name in reverse"
}

The workflow name is reverse is like tset for test. 

This can be set in a windows services/MS Flow to put it in a schedule. 
