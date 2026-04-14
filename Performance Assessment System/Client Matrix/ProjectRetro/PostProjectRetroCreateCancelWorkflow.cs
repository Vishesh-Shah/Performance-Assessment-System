using Inkey.MSCRM.Plugin_V9._0.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Performance_Assessment_System.Common;
using System;

namespace Performance_Assessment_System.Client_Matrix.ProjectRetro
{
    public class PostProjectRetroCreateCancelWorkflow : IPlugin
    {
        public void Execute(IServiceProvider iServiceProvider)
        {
            // Obtain the execution context from the service provider.
            IPluginExecutionContext iPluginExecutionContext =
                (IPluginExecutionContext)iServiceProvider.GetService(typeof(IPluginExecutionContext));

            // Obtain the organization service factory reference.
            IOrganizationServiceFactory iPluginOrganizationServiceFactory =
                (IOrganizationServiceFactory)iServiceProvider.GetService(typeof(IOrganizationServiceFactory));

            // Obtain the tracing service reference.
            ITracingService iTracingService =
                (ITracingService)iServiceProvider.GetService(typeof(ITracingService));

            // Obtain the organization service reference.
            IOrganizationService iOrganizationService =
                iPluginOrganizationServiceFactory.CreateOrganizationService(iPluginExecutionContext.UserId);

            try
            {
                if (iPluginExecutionContext.Depth > 1)
                {

                    return;
                }

                if (!Plugin.ValidateTargetAsEntity(CommonEntities.PROJECTRETRO, iPluginExecutionContext))
                {

                    return;
                }

                Entity projectRetroEntity = (Entity)iPluginExecutionContext.InputParameters["Target"];

                if (projectRetroEntity == null)
                {

                    return;
                }

                // Get project lookup from target entity using helper method
                EntityReference projectEntityReference = Plugin.GetAttributeValue<EntityReference>(projectRetroEntity, "ink_project");

                if (projectEntityReference == null)
                {

                    return;
                }

                // Retrieve workflow definition id from environment variable
                string workflowDefinitionIdString = CommonFunctions.GetEnvironmentVariable(iOrganizationService, "ink_env_projectretroworkflowid");

                if (!string.IsNullOrEmpty(workflowDefinitionIdString))
                {
                    Guid workflowDefinitionId = new Guid(workflowDefinitionIdString);

                    // Retrieve workflow record to get active workflow id using helper method
                    Entity workflowEntity = Plugin.FetchEntityRecord("workflow", workflowDefinitionId,
                        new ColumnSet("workflowid", "name", "activeworkflowid", "statecode", "statuscode"), iOrganizationService);

                    // Get active workflow id from workflow entity using helper method
                    EntityReference activeWorkflowEntityReference = Plugin.GetAttributeValue<EntityReference>(workflowEntity, "activeworkflowid");

                    Guid activeWorkflowId = Guid.Empty;

                    if (activeWorkflowEntityReference != null)
                    {
                        activeWorkflowId = activeWorkflowEntityReference.Id;

                    }
                    else
                    {
                        // Fallback: use definition id if activeworkflowid is not populated
                        activeWorkflowId = workflowDefinitionId;

                    }

                    // Query pending workflow jobs for this project record
                    QueryExpression asyncOperationQueryExpression = new QueryExpression(Entities.ASYNC_OPERATION);
                    asyncOperationQueryExpression.ColumnSet = new ColumnSet(
                        "asyncoperationid",
                        "name",
                        "statecode",
                        "statuscode",
                        "workflowactivationid",
                        "regardingobjectid",
                        "operationtype"
                    );

                    // Only workflow jobs
                    asyncOperationQueryExpression.Criteria.AddCondition("operationtype", ConditionOperator.Equal, 10);

                    // Only same project record
                    asyncOperationQueryExpression.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, projectEntityReference.Id);

                    // Only this workflow active version
                    asyncOperationQueryExpression.Criteria.AddCondition("workflowactivationid", ConditionOperator.Equal, activeWorkflowId);

                    // Only waiting / ready jobs
                    FilterExpression stateFilterExpression = new FilterExpression(LogicalOperator.Or);
                    stateFilterExpression.AddCondition("statecode", ConditionOperator.Equal, SystemJobStatus.READY); // Ready
                    stateFilterExpression.AddCondition("statecode", ConditionOperator.Equal, SystemJobStatus.SUSPENDED); // Suspended
                    asyncOperationQueryExpression.Criteria.AddFilter(stateFilterExpression);

                    EntityCollection pendingWorkflowJobCollection = iOrganizationService.RetrieveMultiple(asyncOperationQueryExpression);



                    foreach (Entity asyncJobEntity in pendingWorkflowJobCollection.Entities)
                    {
                        // Get job name using helper method
                        string jobName = Plugin.GetStringAttributeValue(asyncJobEntity, null, "name");


                        // Cancel the async job by setting statecode = Completed and statuscode = Canceled
                        Entity asyncJobUpdateEntity = new Entity(Entities.ASYNC_OPERATION, asyncJobEntity.Id);
                        Plugin.AddAttribute(asyncJobUpdateEntity, "statecode", new OptionSetValue(SystemJobStatus.COMPLETED));   // Completed
                        Plugin.AddAttribute(asyncJobUpdateEntity, "statuscode", new OptionSetValue(SystemJobStatusReason.CANCELED)); // Canceled
                        iOrganizationService.Update(asyncJobUpdateEntity);

                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(
                    "Error in PostProjectRetroCreateCancelProjectWorkflow : " + ex.Message
                );
            }
        }
    }
}
