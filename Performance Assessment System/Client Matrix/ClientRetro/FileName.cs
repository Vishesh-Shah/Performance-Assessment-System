using Inkey.MSCRM.Plugin_V9._0.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Performance_Assessment_System.Common;
using System;

namespace Performance_Assessment_System.Client_Matrix.ClientRetro
{
    public class PostClientRetroCreateCancelWorkFlow : IPlugin
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

                if (!Plugin.ValidateTargetAsEntity(CommonEntities.CLIENTRETRO, iPluginExecutionContext))
                {
                    return;
                }

                Entity clientRetroEntity = (Entity)iPluginExecutionContext.InputParameters["Target"];

                if (clientRetroEntity == null)
                {
                    return;
                }

                // Get client lookup from target entity using helper method
                EntityReference clientEntityReference = Plugin.GetAttributeValue<EntityReference>(clientRetroEntity, "ink_client");

                if (clientEntityReference == null)
                {
                    return;
                }

                // Retrieve workflow definition id from environment variable
                QueryExpression envVarQueryExpression = new QueryExpression("environmentvariabledefinition");
                envVarQueryExpression.ColumnSet = new ColumnSet("environmentvariabledefinitionid");
                envVarQueryExpression.Criteria.AddCondition("schemaname", ConditionOperator.Equal, "ink_clientretroworkflowid");

                // 1: child table we are joining to get the actual value 
                // 2: first "environmentvariabledefinitionid" — this is the column on the parent table (environmentvariabledefinition) used to join
                // 3: second "environmentvariabledefinitionid" — this is the column on the child table (environmentvariablevalue) used to join

                LinkEntity envVarValueLinkEntity = envVarQueryExpression.AddLink(
                    "environmentvariablevalue", "environmentvariabledefinitionid", "environmentvariabledefinitionid");
                envVarValueLinkEntity.Columns = new ColumnSet("value");
                envVarValueLinkEntity.EntityAlias = "envval";

                EntityCollection envVarCollection = iOrganizationService.RetrieveMultiple(envVarQueryExpression);

                // --- FIX APPLIED HERE: Check if collection has records ---
                if (envVarCollection == null || envVarCollection.Entities.Count == 0)
                {
                    iTracingService.Trace("Environment variable 'ink_clientretroworkflowid' definition or value not found.");
                    return;
                }

                // Get environment variable value using helper method
                string workflowDefinitionIdString = Plugin.GetAttributeValueFromAliasedValue<string>(envVarCollection.Entities[0], "envval.value");

                // --- FIX APPLIED HERE: Check if value is valid ---
                if (string.IsNullOrEmpty(workflowDefinitionIdString))
                {
                    iTracingService.Trace("Environment variable value for 'ink_clientretroworkflowid' is empty.");
                    return;
                }

                Guid workflowDefinitionId;
                if (!Guid.TryParse(workflowDefinitionIdString, out workflowDefinitionId))
                {
                    iTracingService.Trace($"Invalid GUID format in environment variable: {workflowDefinitionIdString}");
                    return;
                }

                // Retrieve workflow record to get active workflow id using helper method
                Entity workflowEntity = Plugin.FetchEntityRecord("workflow", workflowDefinitionId,
                    new ColumnSet("workflowid", "name", "activeworkflowid", "statecode", "statuscode"), iOrganizationService);

                if (workflowEntity == null)
                {
                    iTracingService.Trace("Workflow entity could not be retrieved.");
                    return;
                }

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

                // Query pending workflow jobs for this client record
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

                // Only same client record
                asyncOperationQueryExpression.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, clientEntityReference.Id);

                // Only this workflow active version
                asyncOperationQueryExpression.Criteria.AddCondition("workflowactivationid", ConditionOperator.Equal, activeWorkflowId);

                // Only waiting / ready jobs
                // --- FIX APPLIED HERE: Changed lOperator.Or to LogicalOperator.Or ---
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
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(
                    "Error in PostClientRetroCreateCancelClientWorkflow : " + ex.Message
                );
            }
        }
    }
}