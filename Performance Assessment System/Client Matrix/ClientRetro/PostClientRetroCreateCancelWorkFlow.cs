using Inkey.MSCRM.Plugin_V9._0.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
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
                    iTracingService.Trace("Plugin stopped because Depth > 1");
                    return;
                }

                if (!Plugin.ValidateTargetAsEntity("ink_clientretro", iPluginExecutionContext))
                {
                    iTracingService.Trace("Target is not valid for entity ink_clientretro");
                    return;
                }

                Entity clientRetroEntity = (Entity)iPluginExecutionContext.InputParameters["Target"];

                if (clientRetroEntity == null)
                {
                    iTracingService.Trace("Target entity is null.");
                    return;
                }

                if (!clientRetroEntity.Contains("ink_client"))
                {
                    iTracingService.Trace("Client lookup field 'ink_client' not found in target.");
                    return;
                }

                EntityReference clientRef = clientRetroEntity.GetAttributeValue<EntityReference>("ink_client");

                if (clientRef == null)
                {
                    iTracingService.Trace("Client lookup is null.");
                    return;
                }

                iTracingService.Trace("Client Id : " + clientRef.Id.ToString());

                // URL mathi malelo workflow definition id
                Guid workflowDefinitionId = new Guid("2a97303f-1129-f111-8341-000d3a3ac0a7");
                iTracingService.Trace("Workflow Definition Id : " + workflowDefinitionId.ToString());

                // Get active workflow id from workflow table
                Entity workflowEntity = iOrganizationService.Retrieve(
                    "workflow",
                    workflowDefinitionId,
                    new ColumnSet("workflowid", "name", "activeworkflowid", "statecode", "statuscode")
                );

                if (workflowEntity == null)
                {
                    iTracingService.Trace("Workflow record not found.");
                    return;
                }

                Guid activeWorkflowId = Guid.Empty;

                if (workflowEntity.Contains("activeworkflowid") &&
                    workflowEntity.GetAttributeValue<EntityReference>("activeworkflowid") != null)
                {
                    activeWorkflowId = workflowEntity.GetAttributeValue<EntityReference>("activeworkflowid").Id;
                    iTracingService.Trace("Active Workflow Id : " + activeWorkflowId.ToString());
                }
                else
                {
                    // fallback: in some cases activeworkflowid may not be populated as expected
                    activeWorkflowId = workflowDefinitionId;
                    iTracingService.Trace("Active Workflow Id not found. Fallback to Definition Id : " + activeWorkflowId.ToString());
                }

                QueryExpression queryExpression = new QueryExpression("asyncoperation");
                queryExpression.ColumnSet = new ColumnSet(
                    "asyncoperationid",
                    "name",
                    "statecode",
                    "statuscode",
                    "workflowactivationid",
                    "regardingobjectid",
                    "operationtype"
                );

                // only workflow jobs
                queryExpression.Criteria.AddCondition("operationtype", ConditionOperator.Equal, 10);

                // only same client record
                queryExpression.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, clientRef.Id);

                // only this workflow's active version
                queryExpression.Criteria.AddCondition("workflowactivationid", ConditionOperator.Equal, activeWorkflowId);

                // only waiting / ready jobs
                FilterExpression stateFilter = new FilterExpression(LogicalOperator.Or);
                stateFilter.AddCondition("statecode", ConditionOperator.Equal, 0); // Ready
                stateFilter.AddCondition("statecode", ConditionOperator.Equal, 1); // Suspended
                queryExpression.Criteria.AddFilter(stateFilter);

                EntityCollection pendingWorkflowJobs = iOrganizationService.RetrieveMultiple(queryExpression);

                iTracingService.Trace("Pending workflow jobs found : " + pendingWorkflowJobs.Entities.Count);

                foreach (Entity asyncJob in pendingWorkflowJobs.Entities)
                {
                    Guid asyncJobId = asyncJob.Id;
                    string jobName = asyncJob.Contains("name")
                        ? asyncJob.GetAttributeValue<string>("name")
                        : string.Empty;

                    iTracingService.Trace("Cancelling workflow job : " + jobName + " | " + asyncJobId.ToString());

                    Entity cancelAsyncJob = new Entity("asyncoperation", asyncJobId);
                    cancelAsyncJob["statecode"] = new OptionSetValue(3);   // Completed
                    cancelAsyncJob["statuscode"] = new OptionSetValue(32); // Canceled

                    iOrganizationService.Update(cancelAsyncJob);

                    iTracingService.Trace("Workflow job canceled successfully : " + asyncJobId.ToString());
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(
                    "Error in PostClientRetroCreatCancelClientWorkflow : " + ex.Message
                );
            }
        }
    }
}
