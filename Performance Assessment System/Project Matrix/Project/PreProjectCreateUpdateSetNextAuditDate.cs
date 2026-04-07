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
                    Entity projectEntityPreImage = Plugin.GetPreEntityImage(iPluginExecutionContext, preImageAlias);
                    
                    if (projectEntity != null)
                    {
                        OptionSetValue criticalityOptionSet = Plugin.GetAttributeValue<OptionSetValue>(projectEntity, projectEntityPreImage, "ink_criticality");
                        OptionSetValue sizeOptionSet = Plugin.GetAttributeValue<OptionSetValue>(projectEntity, projectEntityPreImage, "ink_size");
                        DateTime lastAuditDate = Plugin.GetAttributeValue<DateTime>(projectEntity, projectEntityPreImage, "ink_lastauditdate");
                       
                        if(criticalityOptionSet != null && sizeOptionSet != null)
                        {
                            #region Get  Frequency Days from Audit Frequency Matrix
                            // Only retrieve the specific column we need (Frequency Days) to keep the query fast
                            QueryExpression frequencyDayQueryExpression = new QueryExpression(CommonEntities.AUDITFREQUENCYMATRIX);
                            frequencyDayQueryExpression.ColumnSet = new ColumnSet("ink_auditfrequencydays");


                            // 3. Add the conditions
                            FilterExpression frequencyDayFilterExpression = new FilterExpression(LogicalOperator.And);
                            frequencyDayFilterExpression.AddCondition("ink_projectcriticality", ConditionOperator.Equal, criticalityOptionSet.Value);
                            frequencyDayFilterExpression.AddCondition("ink_projectsize", ConditionOperator.Equal, sizeOptionSet.Value);
                            frequencyDayQueryExpression.Criteria.AddFilter(frequencyDayFilterExpression);
                            // 4. Execute the query
                            EntityCollection frequencyDayEntityCollection = iOrganizationService.RetrieveMultiple(frequencyDayQueryExpression);
                            #endregion
                            // 5. Process the result
                            if (frequencyDayEntityCollection.Entities.Count > 0)
                            {
                                // Grab the first matching record
                                Entity frequencyEntity = frequencyDayEntityCollection.Entities[0];

                                // Extract the Frequency Days integer
                                int frequencyDays = Plugin.GetAttributeValue<int>(frequencyEntity,"ink_auditfrequencydays");

                                #region Calculate Next Audit Date
                                DateTime nextAuditDate;
                              
                                if (lastAuditDate == DateTime.MinValue) 
                                {
                                    nextAuditDate= DateTime.UtcNow;
                                }
                                else
                                {
                                    nextAuditDate = lastAuditDate;
                                }
                             

                                // Calculate the next audit gate date by adding the frequency days to the base date

                                nextAuditDate = nextAuditDate.AddDays(frequencyDays);

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
                                Plugin.AddAttribute<DateTime>(projectEntity, "ink_nextauditdate", nextAuditDate);
                                
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

