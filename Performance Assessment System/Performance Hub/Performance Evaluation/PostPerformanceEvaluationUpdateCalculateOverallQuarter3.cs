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
    public class PostPerformanceEvaluationUpdateCalculateOverallQuarter3 : IPlugin
    {
        #region Variable Declaration

        private const string preImageAlias = "PostPerformanceEvaluationUpdateCalculateOverallQuarter3PreImage";


        #endregion

        #region Public Methods

        #region Execute
        /// <summary>
        /// calculating overall rating and percentage for quarter 3 (core expectations + key results)
        /// pre-image fields: ["ink_statusfield"]
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
                        // getting status field to check quarter condition
                        OptionSetValue statusField = Plugin.GetAttributeValue<OptionSetValue>(performanceEvaluationEntity, performanceEvaluationPreImage, "ink_statusfield");

                        if (statusField != null)
                        {
                            int statusFieldValue = statusField.Value;

                            // checking if Q3 is acknowledged
                            if (statusFieldValue == CommonEntities.PerformanceEvaluation.StatusField.Q3Acknowledged)
                            {
                                // fetching all core expectation ratings for Q3
                                QueryExpression coreQuery = new QueryExpression(CommonEntities.COREEXPECTATIONRATING);
                                coreQuery.ColumnSet.AddColumns("ink_quarter3rating");
                                coreQuery.Criteria.AddCondition("ink_performanceevaluations", ConditionOperator.Equal, performanceEvaluationEntity.Id);

                                EntityCollection coreExpectationRatings = iOrganizationService.RetrieveMultiple(coreQuery);


                                if (coreExpectationRatings.Entities.Count > 0)
                                {
                                    int totalRatings = 0;
                                    int numberOfFrequency = 0;

                                    // calculating total core expectation ratings
                                    foreach (Entity core in coreExpectationRatings.Entities)
                                    {
                                        OptionSetValue quarter3Rating = Plugin.GetAttributeValue<OptionSetValue>(core, "ink_quarter3rating");

                                        if (quarter3Rating != null)
                                        {
                                            totalRatings += (quarter3Rating.Value - 826460000);
                                            numberOfFrequency++;
                                        }
                                    }

                                    // calculating average and percentage
                                    decimal overall = Math.Round((decimal)totalRatings / numberOfFrequency, 2);
                                    decimal overallPer = Math.Round(((overall / 5.0m) * 100.0m), 2);

                                    string coreExpectationPercentage = overallPer.ToString();
                                    string quarter3CoreExpectationRating = overall.ToString();

                                    // updating performance evaluation with Q3 core expectation values
                                    Entity updatedPerformanceEvaluation = new Entity(CommonEntities.PERFORMANCEEVALUATION, performanceEvaluationEntity.Id);
                                    Plugin.AddAttribute<string>(updatedPerformanceEvaluation, "ink_quarter3coreexpectations", quarter3CoreExpectationRating);
                                    Plugin.AddAttribute<string>(updatedPerformanceEvaluation, "ink_quater3coreexpectations", coreExpectationPercentage);

                                    iOrganizationService.Update(updatedPerformanceEvaluation);
                                }
                            }

                            // fetching all key result ratings for Q3
                            QueryExpression keyResultRating = new QueryExpression(CommonEntities.KEYRESULTRATING);
                            keyResultRating.ColumnSet.AddColumns("ink_quarter3rating", "ink_objectivenumbering");
                            keyResultRating.Criteria.AddCondition("ink_performanceevaluations", ConditionOperator.Equal, performanceEvaluationEntity.Id);

                            EntityCollection keyResultRatings = iOrganizationService.RetrieveMultiple(keyResultRating);

                            if (keyResultRatings.Entities.Count > 0)
                            {
                                int totalRatings = 0;
                                int numberOfFrequency = 0;

                                // calculating total key result ratings only for objectives
                                foreach (Entity keyResult in keyResultRatings.Entities)
                                {
                                    OptionSetValue quarter3Rating = Plugin.GetAttributeValue<OptionSetValue>(keyResult, "ink_quarter3rating");
                                    int objectiveNumbering = Plugin.GetAttributeValue<int>(keyResult, "ink_objectivenumbering");

                                    if (objectiveNumbering == 0)
                                    {
                                        if (quarter3Rating != null)
                                        {
                                            totalRatings += (quarter3Rating.Value - 826460000);
                                            numberOfFrequency++;
                                        }
                                    }
                                }

                                // calculating average and percentage
                                decimal overall = Math.Round((decimal)totalRatings / numberOfFrequency, 2);
                                decimal overallPer = Math.Round(((overall / 5.0m) * 100.0m), 2);

                                string keyResultPercentage = overallPer.ToString();
                                string quarter3KeyResultRating = overall.ToString();

                                // updating performance evaluation with Q3 objective and key result values
                                Entity updatedPerformanceEvaluation = new Entity(CommonEntities.PERFORMANCEEVALUATION, performanceEvaluationEntity.Id);
                                Plugin.AddAttribute<string>(updatedPerformanceEvaluation, "ink_quater3objectivekeyresultrating", quarter3KeyResultRating);
                                Plugin.AddAttribute<string>(updatedPerformanceEvaluation, "ink_quater3objective", keyResultPercentage);

                                iOrganizationService.Update(updatedPerformanceEvaluation);
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