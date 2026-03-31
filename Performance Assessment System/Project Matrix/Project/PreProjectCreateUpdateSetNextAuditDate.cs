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

namespace Performance_Assessment_System.Project_Matrix
{
    public class PreProjectCreateUpdateSetNextAuditDate :IPlugin
    {
        #region Variable Declaration

        private const string preImageAlias = "PreProjectCreateUpdateSetNextAuditDatePreImage";
        #endregion

        #region Public Methods

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
                if (Plugin.ValidateTargetAsEntity(CommonEntities.PROJECT, iPluginExecutionContext))
                {

                    Entity projectEntity = (Entity)iPluginExecutionContext.InputParameters["Target"];
                    Entity projectPreImage = Plugin.GetPreEntityImage(iPluginExecutionContext, preImageAlias);

                    if(projectEntity != null && projectPreImage != null)
                    {
                        OptionSetValue criticalityOptionSet = Plugin.GetAttributeValue<OptionSetValue>(projectEntity, projectPreImage, "ink_criticality");
                        OptionSetValue sizeOptionSet = Plugin.GetAttributeValue<OptionSetValue>(projectEntity, projectPreImage, "ink_size");

                        if(criticalityOptionSet != null && sizeOptionSet != null)
                        {
                            #region Get  Frequency Days from Audit Frequency Matrix
                            QueryExpression frequencyDayQuery = new QueryExpression(CommonEntities.AUDITFREQUENCYMATRIX)
                            {
                                // Only retrieve the specific column we need (Frequency Days) to keep the query fast
                                ColumnSet = new ColumnSet("ink_auditfrequencydays"),

                                // Set up the logical AND filter
                                Criteria = new FilterExpression(LogicalOperator.And)
                            };

                            // 3. Add the conditions
                            frequencyDayQuery.Criteria.AddCondition("ink_projectcriticality", ConditionOperator.Equal, criticalityOptionSet.Value);
                            frequencyDayQuery.Criteria.AddCondition("ink_projectsize", ConditionOperator.Equal, sizeOptionSet.Value);

                            // 4. Execute the query
                            EntityCollection frequencyDayResults = iOrganizationService.RetrieveMultiple(frequencyDayQuery);
                            #endregion
                            // 5. Process the result
                            if (frequencyDayResults.Entities.Count > 0)
                            {
                                // Grab the first matching record
                                Entity frequencyEntity = frequencyDayResults.Entities[0];

                                // Extract the Frequency Days integer
                                int frequencyDays = frequencyEntity.GetAttributeValue<int>("ink_auditfrequencydays");
                                #region Calculate Next Audit Date
                                DateTime nextAuditDate;
                                DateTime baseDate = DateTime.UtcNow;

                                // Check if there is an existing audit for this project
                                // If yes use last audit date as base
                                // If no use today as base
                                Guid projectId = iPluginExecutionContext.PrimaryEntityId;

                                if (projectId != Guid.Empty)
                                {
                                    QueryExpression auditQuery = new QueryExpression(CommonEntities.AUDIT);
                                    auditQuery.ColumnSet = new ColumnSet("ink_auditeddate");
                                    auditQuery.AddOrder("ink_auditeddate", OrderType.Descending);
                                    auditQuery.TopCount = 1;
                                    auditQuery.Criteria.AddCondition("_ink_projectid_value",
                                        ConditionOperator.Equal, projectId);

                                    EntityCollection auditCollection =
                                        iOrganizationService.RetrieveMultiple(auditQuery);

                                    if (auditCollection != null && auditCollection.Entities.Count > 0)
                                    {
                                        DateTime? lastAuditDate = Plugin.GetAttributeValue<DateTime?>(
                                            auditCollection.Entities[0], "ink_auditeddate");

                                        if (lastAuditDate.HasValue && lastAuditDate.Value != DateTime.MinValue)
                                        {
                                            baseDate = lastAuditDate.Value;
                                        }
                                    }
                                }

                                // Calculate the next audit gate date by adding the frequency days to the base date

                                nextAuditDate = baseDate.AddDays(frequencyDays);

                                if (nextAuditDate.DayOfWeek == DayOfWeek.Saturday)
                                {
                                    // Saturday → move to Monday
                                    nextAuditDate = nextAuditDate.AddDays(2);
                                }
                                else if (nextAuditDate.DayOfWeek == DayOfWeek.Sunday)
                                {
                                    // Sunday → move to Monday
                                    nextAuditDate = nextAuditDate.AddDays(1);
                                }
                                #endregion
                                // Update next audit date on project record using helper method
                                Entity projectUpdateEntity = new Entity("ink_project");
                                projectUpdateEntity.Id = projectEntity.Id;
                                Plugin.AddAttribute(projectUpdateEntity, "ink_nextauditdate", nextAuditDate);
                                iOrganizationService.Update(projectUpdateEntity);
                            }
                            else
                            {
                                // Optional: Throw an error if someone forgot to configure the  frequency matrix for this combination
                                throw new InvalidPluginExecutionException("No Audit Frequency  Matrix configuration found for this Criticality and Size combination.");
                            }
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

        #endregion
    }
}
}
