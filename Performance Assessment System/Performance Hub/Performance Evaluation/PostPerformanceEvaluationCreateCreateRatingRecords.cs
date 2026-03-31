using Inkey.MSCRM.Plugin_V9._0.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Performance_Assessment_System.Common;
using System;
using System.Collections.Generic;
using System.Web.Services.Description;

namespace Performance_Assessment_System.Performance_Hub.Performance_Evaluation
{
    /// <summary>
    /// PreValidation-Operation Plugin on Create of Resource.
    /// Prevents duplicate Resource records based on First Name + Last Name + Reporting Manager + Designation.
    /// Filter Attribute: N/A
    /// Pre-Image Alias: N/A
    /// </summary>
    public class PostPerformanceEvaluationCreateCreateRatingRecords : IPlugin
    {
        #region Public Methods

        #region Execute
        /// <summary>
        /// Validates whether another active Resource already exists with the same
        /// First Name, Last Name, Reporting Manager, and Designation combination.
        /// </summary>
        public void Execute(IServiceProvider iServiceProvider)
        {
            // Obtain the execution context from the service provider.
            IPluginExecutionContext iPluginExecutionContext = (IPluginExecutionContext)iServiceProvider
                .GetService(typeof(IPluginExecutionContext));

            // Obtain the organization service factory reference.
            IOrganizationServiceFactory iOrganizationServiceFactory = (IOrganizationServiceFactory)iServiceProvider
                .GetService(typeof(IOrganizationServiceFactory));

            // Obtain the tracing service reference.
            ITracingService iTracingService = (ITracingService)iServiceProvider
                .GetService(typeof(ITracingService));

            // Obtain the organization service reference.
            IOrganizationService iOrganizationService = iOrganizationServiceFactory
                .CreateOrganizationService(iPluginExecutionContext.UserId);

            try
            {
                if (Plugin.ValidateTargetAsEntity(CommonEntities.PERFORMANCEEVALUATION, iPluginExecutionContext))
                {
                    Entity performanceEvaluationEntity = (Entity)iPluginExecutionContext.InputParameters["Target"];

                    if (performanceEvaluationEntity != null)
                    {
                        EntityReference assesseeReference = Plugin.GetAttributeValue<EntityReference>(performanceEvaluationEntity, "ink_assessee");
                        Entity assesseeEntity = Plugin.FetchEntityRecord(CommonEntities.RESOURCE, assesseeReference.Id, new ColumnSet("ink_designation"), iOrganizationService);
                        OptionSetValue designation = Plugin.GetAttributeValue<OptionSetValue>(assesseeEntity, "ink_designation");

                        EntityReference reviewCycleReference = Plugin.GetAttributeValue<EntityReference>(performanceEvaluationEntity, "ink_performancereviewcycle");

                        if (reviewCycleReference != null && assesseeEntity != null && designation != null)
                        {

                            QueryExpression query = new QueryExpression(CommonEntities.PERFORMANCEEVALUATIONTEMPLATE);
                            query.ColumnSet.AddColumns("ink_name");
                            query.Criteria.AddCondition("ink_designation", ConditionOperator.Equal, designation.Value);
                            query.Criteria.AddCondition("ink_performancereviewcycle", ConditionOperator.Equal, reviewCycleReference.Id);
                            EntityCollection fetchedTemplate = iOrganizationService.RetrieveMultiple(query);

                            Entity templateEntity = null;
                            if (fetchedTemplate.Entities.Count > 0)
                            {
                                iTracingService.Trace(fetchedTemplate.Entities.Count.ToString());

                                foreach (Entity template in fetchedTemplate.Entities)
                                {
                                    templateEntity = template;
                                }

                                #region QueryExpression to get Objectives

                                QueryExpression templateQuery = new QueryExpression(CommonEntities.OBJECTIVE);
                                templateQuery.ColumnSet = new ColumnSet("ink_name", "ink_objectivesid");

                                // Link to intersect table
                                LinkEntity linkIntersect = new LinkEntity(
                                    CommonEntities.OBJECTIVE,
                                    "ink_objectives_ink_performanceevaluatio",
                                    "ink_objectivesid",
                                    "ink_objectivesid",
                                    JoinOperator.Inner
                                );

                                // Link intersect to template
                                LinkEntity linkTemplate = new LinkEntity(
                                    "ink_objectives_ink_performanceevaluatio",
                                    CommonEntities.PERFORMANCEEVALUATIONTEMPLATE,
                                    "ink_performanceevaluationtemplateid",
                                    "ink_performanceevaluationtemplateid",
                                    JoinOperator.Inner
                                );

                                // Filter using Entity Id
                                linkTemplate.LinkCriteria.AddCondition(
                                    "ink_performanceevaluationtemplateid",
                                    ConditionOperator.Equal,
                                    templateEntity.Id
                                );

                                // Chain
                                linkIntersect.LinkEntities.Add(linkTemplate);
                                templateQuery.LinkEntities.Add(linkIntersect);

                                EntityCollection fetchedObjectives = iOrganizationService.RetrieveMultiple(templateQuery);

                                #endregion

                                if (fetchedObjectives.Entities.Count > 0)
                                {
                                    foreach (Entity entity in fetchedObjectives.Entities)
                                    {
                                        string name = Plugin.GetAttributeValue<string>(entity, "ink_name");
                                        iTracingService.Trace(name);
                                    }
                                }
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