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

namespace Performance_Assessment_System.Project_Matrix.Audit
{
    public class PreAuditUpdateCheckAtSubmission : IPlugin
    {
        #region Variable Declaration
       
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
                if (Plugin.ValidateTargetAsEntity(CommonEntities.AUDIT, iPluginExecutionContext))
                {

                    Entity auditEntity = (Entity)iPluginExecutionContext.InputParameters["Target"];

                    if (auditEntity != null)
                    {
                        OptionSetValue auditStatus = Plugin.GetAttributeValue<OptionSetValue>(auditEntity, "ink_status");

                        if (auditStatus != null && auditStatus.Value == Status.SUBMITTED)
                        {
                            #region Checklist Item Rating Validation 
                            // Query all child checklist items linked to this specific audit
                            QueryExpression checklistQueryExpression = new QueryExpression(CommonEntities.CHECKLISTITEM);
                            checklistQueryExpression.ColumnSet = new ColumnSet("ink_rating", "ink_name");

                            // Link the child items to the parent audit we are trying to save 
                            checklistQueryExpression.Criteria.AddCondition("ink_audit", ConditionOperator.Equal, auditEntity.Id);
                           
                            EntityCollection checklistItemEntityCollection = iOrganizationService.RetrieveMultiple(checklistQueryExpression);

                            // Loop through every checklist item found
                            foreach (Entity checklistItem in checklistItemEntityCollection.Entities)
                            {
                                // Extract the rating. If it's totally blank, default our check to 0
                                int rating = checklistItem.Contains("ink_rating") ? Plugin.GetAttributeValue<int>(CommonEntities.CHECKLISTITEM,"ink_rating") : 0;

                                // If any item has a rating of 0 (or wasn't filled out), block the submission!
                                if (rating == 0)
                                {
                                    string itemName = checklistItem.Contains("ink_name") ? Plugin.GetAttributeValue<string>(CommonEntities.CHECKLISTITEM, "ink_name") : "an unknown item";

                                    // This throws a red error banner on the user's screen and cancels the database save
                                    throw new InvalidPluginExecutionException($"Cannot submit audit! The checklist item '{itemName}' has not been scored. Please grade all items before submitting.");
                                }
                            }
                            #endregion

                            #region Check Presence of the Corrective Task from the Checklist Item for this Audit   
                            QueryExpression openTaskQueryExpression = new QueryExpression(CommonEntities.TASK);
                            openTaskQueryExpression.TopCount = 1; // We only need to find ONE open task to block the save
                            openTaskQueryExpression.ColumnSet = new ColumnSet("subject");

                            // Filter for OPEN tasks (StateCode 0)
                            openTaskQueryExpression.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);

                            // The Magic: Join the Task table to the Checklist Item table
                            LinkEntity checklistLink = new LinkEntity(
                                "task",                             // Primary entity of the query
                                "ink_checklistitem",           // Entity we are joining to
                                "regardingobjectid",                // Lookup column on the Task
                                "ink_checklistitemid",         // Primary Key column on the Checklist Item
                                JoinOperator.Inner
                            );

                            // Filter the joined Checklist Items to only include ones for THIS Audit
                            checklistLink.LinkCriteria.AddCondition("ink_audit", ConditionOperator.Equal, auditEntity.Id);
                           
                            openTaskQueryExpression.LinkEntities.Add(checklistLink);

                            // Execute the single, highly-optimized query
                            EntityCollection openTasksEntityCollection = iOrganizationService.RetrieveMultiple(openTaskQueryExpression);

                            if (openTasksEntityCollection.Entities.Count > 0)
                            {
                                // If it found even 1 open task, slam the brakes!
                                throw new InvalidPluginExecutionException("Cannot submit audit! There are still Open Corrective Tasks associated with the checklist items. All tasks must be Completed before submission.");
                            }

                            #endregion

                            Plugin.AddAttribute<DateTime>(auditEntity, "ink_closeddate", DateTime.UtcNow);

                            #region Update Parent Project's Last Audit Date
                            Entity currentAuditEntity = Plugin.FetchEntityRecord(CommonEntities.AUDIT, auditEntity.Id, new ColumnSet("ink_project"),iOrganizationService);

                            if (currentAuditEntity.Contains("ink_project"))
                            {
                                // Extract the lookup reference to the Parent Project
                                EntityReference parentProjectEntityReference = Plugin.GetAttributeValue<EntityReference>(currentAuditEntity, "ink_project");

                                // 3. Create an object for the Parent Project and update the date
                                Entity projectToUpdate = new Entity(CommonEntities.PROJECT);
                                projectToUpdate.Id = parentProjectEntityReference.Id;

                                // REPLACE "ink_lastauditdate" with the actual logical name of the date column on your Project table
                                projectToUpdate["ink_lastauditdate"] = DateTime.UtcNow;

                                // 4. Command the database to update the Parent Project!
                                iOrganizationService.Update(projectToUpdate);
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

