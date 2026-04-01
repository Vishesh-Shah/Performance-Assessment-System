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

            iTracingService.Trace("PostCkecklistItemsUpdateCreateCorrectiveTask plugin execution started.");
            try
            {
                if (Plugin.ValidateTargetAsEntity(CommonEntities.CHECKLISTITEM, iPluginExecutionContext))
                {

                    Entity checklistItemEntity = (Entity)iPluginExecutionContext.InputParameters["Target"];
                    Entity checklistItemPreImage = Plugin.GetPreEntityImage(iPluginExecutionContext, preImageAlias);

                    iTracingService.Trace("Checklist Item Entity and PreImage retrieved successfully.");

                    if (checklistItemEntity != null && checklistItemPreImage != null)
                    {
                        int rating = Plugin.GetAttributeValue<int>(checklistItemEntity, checklistItemPreImage, "ink_rating");

                        iTracingService.Trace($"Rating retrieved successfully. Current Rating: {rating}");

                        if (rating == 1 || rating == 2)
                        {
                            #region Get  Existing Corrective Task from Task Entity if it exists for the Checklist Item  
                            // 1. Query the Task table to see if a task already exists for this Checklist Item
                            QueryExpression taskQueryExpression = new QueryExpression("task");
                            taskQueryExpression.TopCount = 1; // We only need to know if at least ONE exists
                            taskQueryExpression.ColumnSet = new ColumnSet("activityid", "statecode");

                            // Search for any task where "Regarding" equals this specific Checklist Item
                            taskQueryExpression.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, checklistItemEntity.Id);

                            // Optional: Only check for OPEN tasks. If you want it to create a new one if the old one is closed, keep this.
                            // taskQuery.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); 

                            EntityCollection existingTaskEntityCollection = iOrganizationService.RetrieveMultiple(taskQueryExpression);
                            #endregion
                            // 2. If the query returns 0 results, it is safe to create the task!

                            #region Create Corrective Task from Task Entity if it  does not exists for the Checklist Item  
                            if (existingTaskEntityCollection.Entities.Count == 0)
                            {
                                Entity correctiveTaskEntity = new Entity("task");
                                correctiveTaskEntity["subject"] = $"Corrective Action Required: Low Audit Score ({rating})";
                                correctiveTaskEntity["description"] = "An audit checklist item received a critical score. Please review and take corrective action.";
                                correctiveTaskEntity["scheduledend"] = DateTime.UtcNow.AddDays(5);

                                // Link it using your preferred method!
                                correctiveTaskEntity["regardingobjectid"] = new EntityReference("ink_auditchecklistitem", checklistItemEntity.Id);

                                // Create the task
                                iOrganizationService.Create(correctiveTaskEntity);
                            }
                            #endregion

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
                                        // StatusCode 4 = Completed
                                        taskToClose["statuscode"] = new OptionSetValue(4);

                                        iOrganizationService.Update(taskToClose);
                                    }
                                }
                            }
                            #endregion
                        }
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
