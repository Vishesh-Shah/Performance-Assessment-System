using Inkey.MSCRM.Plugin_V9._0.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Performance_Assessment_System.Common;
using System;

namespace Performance_Assessment_System.Performance_Hub.Performance_Evaluation
{
    /// <summary>
    /// PreValidation-Operation Plugin on Create of Performance Evaluation
    /// Creating all Core Expectation Rating Records for the created Performance Evaluation record
    /// Filter Attribute: N/A
    /// Pre-Image Alias: N/A
    /// </summary>
    public class PostPerformanceEvaluationCreateCreateCoreExpectationRatings : IPlugin
    {
        #region Public Methods

        #region Execute

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

                    if( performanceEvaluationEntity != null )
                    {
                        //fetching all core expectation records
                        QueryExpression coreExpectationQueryExpression = new QueryExpression(CommonEntities.COREEXPECTATION);
                        coreExpectationQueryExpression.ColumnSet.AddColumns("ink_name");
                        EntityCollection fetchedCoreExpectations = iOrganizationService.RetrieveMultiple(coreExpectationQueryExpression);

                        string PEname, CEname, name = null;

                        if(fetchedCoreExpectations.Entities.Count > 0)
                        {
                            foreach (Entity coreExpectationEntity in fetchedCoreExpectations.Entities)
                            {
                                #region Creating a Core Expectation Rating Record

                                PEname = Plugin.GetAttributeValue<string>(performanceEvaluationEntity, "ink_name");
                                CEname = Plugin.GetAttributeValue<string>(coreExpectationEntity, "ink_name");
                                name = PEname + " " + CEname;

                                EntityReference performanceEvaluation = new EntityReference(CommonEntities.PERFORMANCEEVALUATION, performanceEvaluationEntity.Id);
                                EntityReference coreExpectation = new EntityReference(CommonEntities.COREEXPECTATION, coreExpectationEntity.Id);

                                Entity coreExpectationRatingEntity = new Entity(CommonEntities.COREEXPECTATIONRATING);
                                Plugin.AddAttribute<string>(coreExpectationRatingEntity, "ink_name", name);
                                Plugin.AddAttribute<DateTime>(coreExpectationRatingEntity, "ink_date", DateTime.UtcNow);
                                Plugin.AddAttribute<EntityReference>(coreExpectationRatingEntity, "ink_performanceevaluations", performanceEvaluation);
                                Plugin.AddAttribute<EntityReference>(coreExpectationRatingEntity, "ink_coreexpectations", coreExpectation);

                                iOrganizationService.Create(coreExpectationRatingEntity);

                                #endregion
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