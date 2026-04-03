using Inkey.MSCRM.Plugin_V9._0.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Performance_Assessment_System.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Performance_Assessment_System.Performance_Hub.Performance_Evaluation
{
    public class PostPerformanceEvaluationUpdateCalculateOverallRating: IPlugin
    {
        #region Variable Declaration

        private const string preImageAlias = "PostPerformanceEvaluationUpdateCalculateOverallRatingPreImage";

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
                if (Plugin.ValidateTargetAsEntity(CommonEntities.PERFORMANCEEVALUATION, iPluginExecutionContext))
                {
                    Entity performanceEvaluationEntity = (Entity)iPluginExecutionContext.InputParameters["Target"];
                    Entity performanceEvaluationPreImage = Plugin.GetPreEntityImage(iPluginExecutionContext, preImageAlias);

                    if (performanceEvaluationPreImage != null)
                    {
                        OptionSetValue statusField = Plugin.GetAttributeValue<OptionSetValue>(performanceEvaluationEntity, performanceEvaluationPreImage, "ink_statusfield");
                        if (statusField != null)
                        {
                            int statusFieldValue = statusField.Value;
                            if (statusFieldValue == CommonEntities.PerformanceEvaluation.StatusField.Q1Submitted)
                            {
                                QueryExpression coreQuery = new QueryExpression(CommonEntities.COREEXPECTATIONRATING);
                                coreQuery.ColumnSet.AddColumns("ink_quarter1rating");
                                coreQuery.Criteria.AddCondition("ink_performanceevaluations", ConditionOperator.Equal, performanceEvaluationEntity.Id);
                                EntityCollection coreExpectationRatings = iOrganizationService.RetrieveMultiple(coreQuery);
                                
                                decimal overallRating = 0;
                                decimal totalRatings = 0;
                                decimal numberOfFrequency = 0;
                                if (coreExpectationRatings.Entities.Count > 0)
                                {
                                    foreach (Entity core in coreExpectationRatings.Entities) 
                                    { 
                                        OptionSetValue quarter1Rating = Plugin.GetAttributeValue<OptionSetValue>(core, "ink_quarter1rating");
                                        if (quarter1Rating != null)
                                        {

                                            totalRatings += (quarter1Rating.Value - 826460000);
                                            numberOfFrequency++;
                                        }
                                    }

                                    overallRating= Math.Round((totalRatings/numberOfFrequency),2);
                                    iTracingService.Trace(overallRating.ToString());
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
