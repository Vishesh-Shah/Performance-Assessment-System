using Inkey.MSCRM.Plugin_V9._0.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Performance_Assessment_System.Common;
using System;
using System.Data.SqlTypes;

namespace Performance_Assessment_System.Performance_Hub.Resource
{
    public class PreResourceOnActionCreatePerformanceEvaluations : IPlugin
    {
        #region Variable Declaration


        #endregion

        #region Public Methods

        #region Execute
        /// <summary>
        /// Plugin Description: Populate totalAmount on Department on creation of MonthlyPayment based on Employee
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
                QueryExpression resourceQuery = new QueryExpression(CommonEntities.RESOURCE);
                resourceQuery.ColumnSet = new ColumnSet("ink_firstname", "ink_lastname", "ink_reportingmanager", "ink_designation");
                EntityCollection fetchedResourceEntityCollection = iOrganizationService.RetrieveMultiple(resourceQuery);

                if (fetchedResourceEntityCollection.Entities.Count > 0)
                {
                    foreach (Entity resourceEntity in fetchedResourceEntityCollection.Entities)
                    {
                        EntityReference assesseeEntityReference = new EntityReference(resourceEntity.LogicalName, resourceEntity.Id);
                        EntityReference reportingManagerEntityReference = Plugin.GetAttributeValue<EntityReference>(resourceEntity, "ink_reportingmanager");
                        string firstName = Plugin.GetAttributeValue<string>(resourceEntity, "ink_firstname");
                        string lastName = Plugin.GetAttributeValue<string>(resourceEntity, "ink_lastname");

                        if (reportingManagerEntityReference != null)
                        {

                            OptionSetValue designation = Plugin.GetAttributeValue<OptionSetValue>(resourceEntity, "ink_designation");

                            QueryExpression templateQuery = new QueryExpression(CommonEntities.PERFORMANCEEVALUATIONTEMPLATE);
                            templateQuery.ColumnSet = new ColumnSet("ink_performancereviewcycle");
                            templateQuery.Criteria.AddCondition("ink_designation", ConditionOperator.Equal, designation.Value);

                            EntityCollection fetchedTemplateEntityCollection = iOrganizationService.RetrieveMultiple(templateQuery);

                            if (fetchedTemplateEntityCollection.Entities.Count > 0)
                            {
                                Entity templateEntity = fetchedTemplateEntityCollection.Entities[0];
                                EntityReference performanceReviewCycle = Plugin.GetAttributeValue<EntityReference>(
                                    templateEntity, "ink_performancereviewcycle"
                                );

                                string cycleName = performanceReviewCycle != null ? performanceReviewCycle.Name : "";
                                string PEName = firstName + " " + lastName + " " + cycleName;

                                Entity performanceEvaluationEntity = new Entity(CommonEntities.PERFORMANCEEVALUATION);
                                Plugin.AddAttribute<string>(performanceEvaluationEntity, "ink_name", PEName);
                                Plugin.AddAttribute<EntityReference>(performanceEvaluationEntity, "ink_assessee", assesseeEntityReference);
                                Plugin.AddAttribute<EntityReference>(performanceEvaluationEntity, "ink_assessor", reportingManagerEntityReference);
                                Plugin.AddAttribute<EntityReference>(performanceEvaluationEntity, "ink_performancereviewcycle", performanceReviewCycle);

                                iOrganizationService.Create(performanceEvaluationEntity);
                                iTracingService.Trace("Completed......" + PEName);

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
