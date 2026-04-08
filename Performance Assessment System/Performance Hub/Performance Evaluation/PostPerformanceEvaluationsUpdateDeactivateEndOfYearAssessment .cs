using Inkey.MSCRM.Plugin_V9._0.Common;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Performance_Assessment_System.Common;
using System;

namespace Performance_Assessment_System.Plugins
{
    /// <summary>
    /// Post-Operation Plugin on Update of ink_performanceevaluations.
    /// This plugin handles End of Year logic.
    /// When ink_statusfield becomes Q4 Acknowledged → Full assessment becomes inactive/read-only.
    /// It also deactivates related Core Expectation Ratings and Key Result Ratings.
    /// Filter Attribute: ink_statusfield
    /// Pre-Image Alias: PostPerformanceEvaluationsUpdateDeactivateEndOfYearAssessmentPreImage
    /// Pre-Image Attributes: ink_statusfield
    /// </summary>
    public class PostPerformanceEvaluationsUpdateDeactivateEndOfYearAssessment : IPlugin
    {
        #region Variable Declaration

        private const string preImageAlias = "PostPerformanceEvaluationsUpdateDeactivateEndOfYearAssessmentPreImage";
        private const string customStatusField = "ink_statusfield";
        private const string stateCodeField = "statecode";
        private const string performanceEvaluationLookupField = "ink_performanceevaluations";

        private const int activeState = 0;
        private const int inactiveState = 1;
        private const int inactiveStatus = 2;

        #endregion

        #region Execute Method

        /// <summary>
        /// This method triggers when Performance Evaluation is updated.
        /// It checks whether ink_statusfield is changed to Q4 Acknowledged.
        /// If yes, then all related records are deactivated using ExecuteMultiple.
        /// </summary>
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider
                .GetService(typeof(IPluginExecutionContext));

            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider
                .GetService(typeof(IOrganizationServiceFactory));

            IOrganizationService service = factory.CreateOrganizationService(context.UserId);

            ITracingService tracingService = (ITracingService)serviceProvider
                .GetService(typeof(ITracingService));

            try
            {
                tracingService.Trace("===== End Of Year Plugin Started =====");

                if (!Plugin.ValidateTargetAsEntity(CommonEntities.PERFORMANCEEVALUATION, context))
                {
                    tracingService.Trace("Target is not valid.");
                    return;
                }

                Entity performanceEvaluationEntity = (Entity)context.InputParameters["Target"];
                Entity performanceEvaluationPreImage = Plugin.GetPreEntityImage(context, preImageAlias);

                if (performanceEvaluationEntity == null)
                {
                    tracingService.Trace("Target is null.");
                    return;
                }

                if (performanceEvaluationPreImage == null)
                {
                    tracingService.Trace("Pre Image is null.");
                    return;
                }

                OptionSetValue currentStatusField = Plugin.GetAttributeValue<OptionSetValue>(
                    performanceEvaluationEntity,
                    performanceEvaluationPreImage,
                    customStatusField);

                OptionSetValue oldStatusField = Plugin.GetAttributeValue<OptionSetValue>(
                    performanceEvaluationPreImage,
                    customStatusField);

                tracingService.Trace("Current ink_statusfield: " + (currentStatusField != null ? currentStatusField.Value.ToString() : "null"));
                tracingService.Trace("Old ink_statusfield: " + (oldStatusField != null ? oldStatusField.Value.ToString() : "null"));

                if (currentStatusField == null)
                {
                    tracingService.Trace("Current custom status is null. Plugin stopped.");
                    return;
                }

                if (currentStatusField.Value != CommonEntities.PerformanceEvaluation.StatusField.Q4Acknowledged)
                {
                    tracingService.Trace("Current custom status is not Q4 Acknowledged. Plugin stopped.");
                    return;
                }

                if (oldStatusField != null &&
                    oldStatusField.Value == CommonEntities.PerformanceEvaluation.StatusField.Q4Acknowledged)
                {
                    tracingService.Trace("Old custom status already Q4 Acknowledged. Plugin stopped.");
                    return;
                }

                Guid performanceEvaluationId = performanceEvaluationEntity.Id != Guid.Empty
                    ? performanceEvaluationEntity.Id
                    : performanceEvaluationPreImage.Id;

                if (performanceEvaluationId == Guid.Empty)
                {
                    tracingService.Trace("Performance Evaluation Id is empty.");
                    return;
                }

                ExecuteMultipleRequest batchRequest = new ExecuteMultipleRequest()
                {
                    Settings = new ExecuteMultipleSettings()
                    {
                        ContinueOnError = false,
                        ReturnResponses = true
                    },
                    Requests = new OrganizationRequestCollection()
                };

                AddCoreExpectationRequests(service, tracingService, performanceEvaluationId, batchRequest);
                AddKeyResultRequests(service, tracingService, performanceEvaluationId, batchRequest);
                AddParentRequest(tracingService, performanceEvaluationId, batchRequest);

                tracingService.Trace("Total Requests In Batch: " + batchRequest.Requests.Count.ToString());

                if (batchRequest.Requests.Count > 0)
                {
                    ExecuteMultipleResponse batchResponse = (ExecuteMultipleResponse)service.Execute(batchRequest);
                    tracingService.Trace("Batch executed successfully. Response Count: " + batchResponse.Responses.Count.ToString());
                }

                tracingService.Trace("===== End Of Year Plugin Completed =====");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        #endregion

        #region Parent

        /// <summary>
        /// Add parent deactivate request.
        /// </summary>
        private void AddParentRequest(ITracingService tracingService, Guid performanceEvaluationId, ExecuteMultipleRequest batchRequest)
        {
            tracingService.Trace("Adding parent deactivate request.");

            SetStateRequest setStateRequest = new SetStateRequest
            {
                EntityMoniker = new EntityReference(CommonEntities.PERFORMANCEEVALUATION, performanceEvaluationId),
                State = new OptionSetValue(inactiveState),
                Status = new OptionSetValue(inactiveStatus)
            };

            batchRequest.Requests.Add(setStateRequest);
        }

        #endregion

        #region Core Expectation

        /// <summary>
        /// Fetch Core Expectation Ratings and add deactivate requests.
        /// </summary>
        private void AddCoreExpectationRequests(IOrganizationService service, ITracingService tracingService, Guid performanceEvaluationId, ExecuteMultipleRequest batchRequest)
        {
            QueryExpression query = new QueryExpression(CommonEntities.COREEXPECTATIONRATING);
            query.ColumnSet = new ColumnSet("ink_coreexpectationratingsid");
            query.Criteria.AddCondition(performanceEvaluationLookupField, ConditionOperator.Equal, performanceEvaluationId);
            query.Criteria.AddCondition(stateCodeField, ConditionOperator.Equal, activeState);

            EntityCollection records = service.RetrieveMultiple(query);

            tracingService.Trace("Core Expectation Rating Count: " + records.Entities.Count.ToString());

            foreach (Entity rec in records.Entities)
            {
                SetStateRequest request = new SetStateRequest
                {
                    EntityMoniker = rec.ToEntityReference(),
                    State = new OptionSetValue(inactiveState),
                    Status = new OptionSetValue(inactiveStatus)
                };

                batchRequest.Requests.Add(request);
            }
        }

        #endregion

        #region Key Result

        /// <summary>
        /// Fetch Key Result Ratings and add deactivate requests.
        /// </summary>
        private void AddKeyResultRequests(IOrganizationService service, ITracingService tracingService, Guid performanceEvaluationId, ExecuteMultipleRequest batchRequest)
        {
            QueryExpression query = new QueryExpression(CommonEntities.KEYRESULTRATING);
            query.ColumnSet = new ColumnSet("ink_keyresultratingsid");
            query.Criteria.AddCondition(performanceEvaluationLookupField, ConditionOperator.Equal, performanceEvaluationId);
            query.Criteria.AddCondition(stateCodeField, ConditionOperator.Equal, activeState);

            EntityCollection records = service.RetrieveMultiple(query);

            tracingService.Trace("Key Result Rating Count: " + records.Entities.Count.ToString());

            foreach (Entity rec in records.Entities)
            {
                SetStateRequest request = new SetStateRequest
                {
                    EntityMoniker = rec.ToEntityReference(),
                    State = new OptionSetValue(inactiveState),
                    Status = new OptionSetValue(inactiveStatus)
                };

                batchRequest.Requests.Add(request);
            }
        }

        #endregion
    }
}