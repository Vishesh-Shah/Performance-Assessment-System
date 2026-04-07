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
    public class PostPerformanceEvaluationUpdateCalculateYearEndOverallRating : IPlugin
    {
        #region Variable Declaration

        private const string preImageAlias = "PostPerformanceEvaluationUpdateCalculateYearEndOverallRatingPreImage";


        #endregion

        #region Public Methods

        #region Execute
        /// <summary>
        /// calculating overall rating of the whole year and updating it in the performance evaluation record
        /// pre-image fields: ["ink_quarter4coreexpectations", "ink_quater4objectivekeyresultrating", "ink_quarter3coreexpectations", "ink_quater3objectivekeyresultrating"
        ///                     "ink_quarter2coreexpectations", "ink_quater2objectivekeyresultrating", "ink_quarter1coreexpectations", "ink_quater1objectivekeyresultrating"]
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

                    if (performanceEvaluationPreImage != null && performanceEvaluationEntity != null)
                    {
                        // getting all quarters overall ratings
                        string quarter4Core = Plugin.GetAttributeValue<string>(performanceEvaluationEntity, performanceEvaluationPreImage, "ink_quarter4coreexpectations");
                        string quarter4Key = Plugin.GetAttributeValue<string>(performanceEvaluationEntity, performanceEvaluationPreImage, "ink_quater4objectivekeyresultrating");
                        string quarter3Core = Plugin.GetAttributeValue<string>(performanceEvaluationEntity, performanceEvaluationPreImage, "ink_quarter3coreexpectations");
                        string quarter3Key = Plugin.GetAttributeValue<string>(performanceEvaluationEntity, performanceEvaluationPreImage, "ink_quater3objectivekeyresultrating");
                        string quarter2Core = Plugin.GetAttributeValue<string>(performanceEvaluationEntity, performanceEvaluationPreImage, "ink_quarter2coreexpectations");
                        string quarter2Key = Plugin.GetAttributeValue<string>(performanceEvaluationEntity, performanceEvaluationPreImage, "ink_quater2objectivekeyresultrating");
                        string quarter1Core = Plugin.GetAttributeValue<string>(performanceEvaluationEntity, performanceEvaluationPreImage, "ink_quarter1coreexpectations");
                        string quarter1Key = Plugin.GetAttributeValue<string>(performanceEvaluationEntity, performanceEvaluationPreImage, "ink_quater1objectivekeyresultrating");

                        // calculating total of both ratings 
                        decimal coreTotal = Convert.ToDecimal(quarter1Core) + Convert.ToDecimal(quarter2Core) + Convert.ToDecimal(quarter3Core) + Convert.ToDecimal(quarter4Core);
                        decimal keyTotal = Convert.ToDecimal(quarter1Key) + Convert.ToDecimal(quarter2Key) + Convert.ToDecimal(quarter3Key) + Convert.ToDecimal(quarter4Key);

                        //calculating total to percentage
                        string corePercentage = Math.Round((((coreTotal / 4)/5) *100), 0).ToString() + "%";
                        string keyPercentage = Math.Round((((keyTotal / 4)/5) * 100), 0).ToString() + "%";

                        //updating the record
                        Entity performanceEvaluation = new Entity(CommonEntities.PERFORMANCEEVALUATION, performanceEvaluationEntity.Id);
                        Plugin.AddAttribute<string>(performanceEvaluation, "ink_overallcoreexpectations", corePercentage);
                        Plugin.AddAttribute<string>(performanceEvaluation, "ink_overallobjective", keyPercentage);
                        iOrganizationService.Update(performanceEvaluation);
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
