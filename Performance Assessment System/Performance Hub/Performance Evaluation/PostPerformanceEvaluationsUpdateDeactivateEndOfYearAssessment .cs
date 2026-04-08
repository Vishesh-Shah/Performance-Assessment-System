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
    /// When status becomes Q4 Acknowledged → Full assessment becomes inactive/read-only.
    /// It also deactivates related Core Expectation Ratings and Key Result Ratings.
    /// Filter Attribute: statuscode
    /// Pre-Image Alias: PostPerformanceEvaluationsUpdateDeactivateEndOfYearAssessmentPreImage
    /// Pre-Image Attributes: statuscode
    /// </summary>
    public class PostPerformanceEvaluationsUpdateDeactivateEndOfYearAssessment : IPlugin
    {
        #region Variable Declaration

        private const string preImageAlias = "PostPerformanceEvaluationsUpdateDeactivateEndOfYearAssessmentPreImage";

        private const int inactiveState = 1;
        private const int inactiveStatus = 2;

        #endregion

        #region Execute Method

        /// <summary>
        /// This method triggers when Performance Evaluation is updated.
        /// It checks whether status is changed to Q4 Acknowledged.
        /// If yes, then all related records are deactivated using ExecuteMultiple.
        /// </summary>
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider
                .GetService(typeof(IPluginExecutionContext));

            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider
                .GetService(typeof(IOrganizationServiceFactory));

            IOrganizationService service = factory.CreateOrganizationService(context.UserId);

            try
            {
                if (Plugin.ValidateTargetAsEntity(CommonEntities.PERFORMANCEEVALUATION, context))
                {
                    Entity target = (Entity)context.InputParameters["Target"];
                    Entity preImage = Plugin.GetPreEntityImage(context, preImageAlias);

                    if (target == null || preImage == null)
                        return;

                    // Get Status
                    OptionSetValue status = Plugin.GetAttributeValue<OptionSetValue>(target, preImage, "statuscode");
                    OptionSetValue oldStatus = Plugin.GetAttributeValue<OptionSetValue>(preImage, "statuscode");

                    // 🔥 Check Q4 Acknowledged
                    if (status != null &&
                        status.Value == CommonEntities.PerformanceEvaluation.StatusField.Q4Acknowledged &&
                        (oldStatus == null || oldStatus.Value != CommonEntities.PerformanceEvaluation.StatusField.Q4Acknowledged))
                    {
                        Guid evaluationId = target.Id != Guid.Empty ? target.Id : preImage.Id;

                        if (evaluationId != Guid.Empty)
                        {
                            // Create Batch Request
                            ExecuteMultipleRequest batch = new ExecuteMultipleRequest()
                            {
                                Settings = new ExecuteMultipleSettings()
                                {
                                    ContinueOnError = false,
                                    ReturnResponses = false
                                },
                                Requests = new OrganizationRequestCollection()
                            };

                            // Add Child Records
                            AddCoreExpectationRequests(service, evaluationId, batch);
                            AddKeyResultRequests(service, evaluationId, batch);

                            // Add Parent
                            AddParentRequest(evaluationId, batch);

                            // Execute Batch
                            if (batch.Requests.Count > 0)
                            {
                                service.Execute(batch);
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

        #region Parent

        /// <summary>
        /// Add parent deactivate request
        /// </summary>
        private void AddParentRequest(Guid id, ExecuteMultipleRequest batch)
        {
            SetStateRequest req = new SetStateRequest
            {
                EntityMoniker = new EntityReference(CommonEntities.PERFORMANCEEVALUATION, id),
                State = new OptionSetValue(inactiveState),
                Status = new OptionSetValue(inactiveStatus)
            };

            batch.Requests.Add(req);
        }

        #endregion

        #region Core Expectation

        /// <summary>
        /// Fetch Core Expectation Ratings and add deactivate requests
        /// </summary>
        private void AddCoreExpectationRequests(IOrganizationService service, Guid evaluationId, ExecuteMultipleRequest batch)
        {
            QueryExpression query = new QueryExpression(CommonEntities.COREEXPECTATIONRATING);
            query.ColumnSet = new ColumnSet("ink_coreexpectationratingsid");

            query.Criteria.AddCondition("ink_performanceevaluations", ConditionOperator.Equal, evaluationId);
            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);

            EntityCollection records = service.RetrieveMultiple(query);

            foreach (Entity rec in records.Entities)
            {
                SetStateRequest req = new SetStateRequest
                {
                    EntityMoniker = rec.ToEntityReference(),
                    State = new OptionSetValue(inactiveState),
                    Status = new OptionSetValue(inactiveStatus)
                };

                batch.Requests.Add(req);
            }
        }

        #endregion

        #region Key Result

        /// <summary>
        /// Fetch Key Result Ratings and add deactivate requests
        /// </summary>
        private void AddKeyResultRequests(IOrganizationService service, Guid evaluationId, ExecuteMultipleRequest batch)
        {
            QueryExpression query = new QueryExpression(CommonEntities.KEYRESULTRATING);
            query.ColumnSet = new ColumnSet("ink_keyresultratingsid");

            query.Criteria.AddCondition("ink_performanceevaluations", ConditionOperator.Equal, evaluationId);
            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);

            EntityCollection records = service.RetrieveMultiple(query);

            foreach (Entity rec in records.Entities)
            {
                SetStateRequest req = new SetStateRequest
                {
                    EntityMoniker = rec.ToEntityReference(),
                    State = new OptionSetValue(inactiveState),
                    Status = new OptionSetValue(inactiveStatus)
                };

                batch.Requests.Add(req);
            }
        }

        #endregion
    }
}