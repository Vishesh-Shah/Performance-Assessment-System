using Inkey.MSCRM.Plugin_V9._0.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
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
                    Plugin.TraceLog("Plugin stopped because Depth > 1", iTracingService);
                    return;
                }

                if (!Plugin.ValidateTargetAsEntity("ink_projectretro", iPluginExecutionContext))
                {
                    Plugin.TraceLog("Target is not valid for entity ink_projectretro", iTracingService);
                    return;
                }

                Entity projectRetroEntity = (Entity)iPluginExecutionContext.InputParameters["Target"];

                if (projectRetroEntity == null)
                {
                    Plugin.TraceLog("Target entity is null.", iTracingService);
                    return;
                }

                // Get project lookup from target entity using helper method
                EntityReference projectEntityReference = Plugin.GetAttributeValue<EntityReference>(projectRetroEntity, "ink_project");

                if (projectEntityReference == null)
                {
                    Plugin.TraceLog("Project lookup is null.", iTracingService);
                    return;
                }

                Plugin.TraceLog("Project Id : " + projectEntityReference.Id.ToString(), iTracingService);

                // Workflow definition id - get this from the URL of the workflow record
                Guid workflowDefinitionId = new Guid("b8457006-f12b-f111-88b4-000d3a3ac0a7");
                Plugin.TraceLog("Workflow Definition Id : " + workflowDefinitionId.ToString(), iTracingService);

                // Retrieve workflow record to get active workflow id using helper method
                Entity workflowEntity = Plugin.FetchEntityRecord("workflow", workflowDefinitionId,
                    new ColumnSet("workflowid", "name", "activeworkflowid", "statecode", "statuscode"), iOrganizationService);

                if (workflowEntity == null)
                {
                    Plugin.TraceLog("Workflow record not found.", iTracingService);
                    return;
                }

                // Get active workflow id from workflow entity using helper method
                EntityReference activeWorkflowEntityReference = Plugin.GetAttributeValue<EntityReference>(workflowEntity, "activeworkflowid");

                Guid activeWorkflowId = Guid.Empty;

                if (activeWorkflowEntityReference != null)
                {
                    activeWorkflowId = activeWorkflowEntityReference.Id;
                    Plugin.TraceLog("Active Workflow Id : " + activeWorkflowId.ToString(), iTracingService);
                }
                else
                {
                    // Fallback: use definition id if activeworkflowid is not populated
                    activeWorkflowId = workflowDefinitionId;
                    Plugin.TraceLog("Active Workflow Id not found. Fallback to Definition Id : " + activeWorkflowId.ToString(), iTracingService);
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
                stateFilterExpression.AddCondition("statecode", ConditionOperator.Equal, 0); // Ready
                stateFilterExpression.AddCondition("statecode", ConditionOperator.Equal, 1); // Suspended
                asyncOperationQueryExpression.Criteria.AddFilter(stateFilterExpression);

                EntityCollection pendingWorkflowJobCollection = iOrganizationService.RetrieveMultiple(asyncOperationQueryExpression);

                Plugin.TraceLog("Pending workflow jobs found : " + pendingWorkflowJobCollection.Entities.Count, iTracingService);

                foreach (Entity asyncJobEntity in pendingWorkflowJobCollection.Entities)
                {
                    // Get job name using helper method
                    string jobName = Plugin.GetStringAttributeValue(asyncJobEntity, null, "name");

                    Plugin.TraceLog("Cancelling workflow job : " + jobName + " | " + asyncJobEntity.Id.ToString(), iTracingService);

                    // Cancel the async job by setting statecode = Completed and statuscode = Canceled
                    Entity asyncJobUpdateEntity = new Entity(Entities.ASYNC_OPERATION, asyncJobEntity.Id);
                    Plugin.AddAttribute(asyncJobUpdateEntity, "statecode", new OptionSetValue(3));   // Completed
                    Plugin.AddAttribute(asyncJobUpdateEntity, "statuscode", new OptionSetValue(32)); // Canceled
                    iOrganizationService.Update(asyncJobUpdateEntity);

                    Plugin.TraceLog("Workflow job canceled successfully : " + asyncJobEntity.Id.ToString(), iTracingService);
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
