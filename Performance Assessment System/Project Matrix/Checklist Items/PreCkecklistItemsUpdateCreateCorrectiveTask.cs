using Inkey.MSCRM.Plugin_V9._0.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Performance_Assessment_System.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Services.Description;

namespace Performance_Assessment_System.Project_Matrix.Checklist_Items
{
    public class PreCkecklistItemsUpdateCreateCorrectiveTask :IPlugin
    {
        #region Variable Declaration
        private const string preImageAlias = "PreChecklistItemUpdateCreateCorrectiveTaskPreImage";
        #endregion
        #region Execute
        /// <summary>
        /// 
        /// </summary>
        public void Execute(IServiceProvider iServiceProvider)
        {
            // Obtain the execution context from the service provider.
            IPluginExecutionContext iPluginExecutionContext = (IPluginExecutionContext)iServiceProvider.GetService(typeof(IPluginExecutionContext));

            // Obtain the organization service factory reference.
            IOrganizationServiceFactory iOrganizationServiceFactory = (IOrganizationServiceFactory)iServiceProvider.GetService(typeof(IOrganizationServiceFactory));

            // Obtain the tracing service reference.
            ITracingService iTracingService = (ITracingService)iServiceProvider.GetService(typeof(ITracingService));

            // Obtain the organization service reference.
            IOrganizationService iOrganizationService = iOrganizationServiceFactory.CreateOrganizationService(iPluginExecutionContext.UserId);

           
            try
            {
                if (Plugin.ValidateTargetAsEntity(CommonEntities.CHECKLISTITEM, iPluginExecutionContext))
                {

                    Entity checklistItemEntity = (Entity)iPluginExecutionContext.InputParameters["Target"];
                    Entity checklistItemPreImage = Plugin.GetPreEntityImage(iPluginExecutionContext, preImageAlias);

                  

                    if (checklistItemEntity != null && checklistItemPreImage != null)
                    {
                        int rating = Plugin.GetAttributeValue<int>(checklistItemEntity, checklistItemPreImage, "ink_rating");
                        string questionText = "Unknown Question";
                        string projectName = "Unknown Project";

                     

                            #region Get  Existing Corrective Task from Task Entity if it exists for the Checklist Item  
                            // 1. Query the Task table to see if a task already exists for this Checklist Item
                            QueryExpression taskQueryExpression = new QueryExpression("task");
                            taskQueryExpression.TopCount = 1; // We only need to know if at least ONE exists
                            taskQueryExpression.ColumnSet = new ColumnSet("activityid", "statecode");

                            // Search for any task where "Regarding" equals this specific Checklist Item
                            taskQueryExpression.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, checklistItemEntity.Id);

                        // Optional: Only check for OPEN tasks. If you want it to create a new one if the old one is closed, keep this.
                            taskQueryExpression.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); 

                            EntityCollection existingTaskEntityCollection = iOrganizationService.RetrieveMultiple(taskQueryExpression);
                            #endregion
                            // 2. If the query returns 0 results, it is safe to create the task!

                        if (rating == 1 || rating == 2)
                        {
                            // Step 1: Retrieve the full Checklist Item to get its Name and the Audit Lookup
                            Entity fullChecklistItem = Plugin.FetchEntityRecord(checklistItemEntity.LogicalName, checklistItemEntity.Id, new ColumnSet("ink_name", "ink_audit"),iOrganizationService);

                            if (fullChecklistItem.Contains("ink_name"))
                            {
                                questionText = fullChecklistItem.GetAttributeValue<string>("ink_name");
                            }

                            // Step 2: Extract the Audit Lookup and retrieve the Audit to get the Project
                            if (fullChecklistItem.Contains("ink_audit"))
                            {
                                EntityReference auditEntityReference = Plugin.GetAttributeValue<EntityReference>(fullChecklistItem,"ink_audit");

                                // Retrieve the Parent Audit just to get the Project Lookup
                                Entity parentAuditEntity = Plugin.FetchEntityRecord(auditEntityReference.LogicalName, auditEntityReference.Id, new ColumnSet("ink_project"),iOrganizationService);

                                if (parentAudit.Contains("ink_project"))
                                {
                                    EntityReference projectEntityReference = Plugin.GetAttributeValue<EntityReference>(parentAuditEntity,"ink_project");

                                    // DATA-VERSE MAGIC TRICK: When you retrieve a Lookup, Dataverse automatically 
                                    // attaches the text name of the record to the '.Name' property!
                                    projectName = projectEntityReference.Name;
                                }
                            }
                            // ==================================================
                            #region Create Corrective Task from Task Entity if it  does not exists for the Checklist Item  
                            if (existingTaskEntityCollection.Entities.Count == 0)
                            {
                                Entity correctiveTaskEntity = new Entity("task");
                                correctiveTaskEntity["subject"] = $"Corrective Action Required: Low Audit Score ({rating})";
                                correctiveTaskEntity["description"] = $"A critical score ({rating}) was recorded during an audit.\n\n" +
                                                                        $"Project: {projectName}\n" +
                                                                        $"Question: {questionText}\n\n" +
                                                                        $"Please review this item and take immediate corrective action.";
                                correctiveTaskEntity["scheduledend"] = DateTime.UtcNow.AddDays(5);

                                // Link it using your preferred method!
                                correctiveTaskEntity["regardingobjectid"] = new EntityReference("ink_checklistitem", checklistItemEntity.Id);

                                // Create the task
                                iOrganizationService.Create(correctiveTaskEntity);
                            }
                            #endregion
                        }
                        #region Close Corrective Task from Task Entity if rating > 2 for the Checklist Item  
                        else if (rating > 2)
                        {
                            // If the query found any existing tasks, loop through and close them
                            if (existingTaskEntityCollection.Entities.Count > 0)
                            {
                                foreach (Entity existingTask in existingTaskEntityCollection.Entities)
                                {
                                    Entity taskToClose = new Entity("task");
                                    taskToClose.Id = existingTask.Id;

                                    // StateCode 1 = Completed
                                    taskToClose["statecode"] = new OptionSetValue(1);
                                    // StatusCode 5 = Completed
                                    taskToClose["statuscode"] = new OptionSetValue(5);

                                    iOrganizationService.Update(taskToClose);
                                }
                            }
                        }
                            #endregion
                        
                    }
                }
            }

            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        #endregion
    }
}
