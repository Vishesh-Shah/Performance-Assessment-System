using Inkey.MSCRM.Plugin_V9._0.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Performance_Assessment_System.Resource
{
    public class PreResourceCreateValidateDuplicate : IPlugin
    {
        #region Public Methods

        #region Execute
        public void Execute(IServiceProvider iServiceProvider)
        {
            IPluginExecutionContext iPluginExecutionContext =
                (IPluginExecutionContext)iServiceProvider.GetService(typeof(IPluginExecutionContext));

            IOrganizationServiceFactory iOrganizationServiceFactory =
                (IOrganizationServiceFactory)iServiceProvider.GetService(typeof(IOrganizationServiceFactory));

            ITracingService iTracingService =
                (ITracingService)iServiceProvider.GetService(typeof(ITracingService));

            IOrganizationService iOrganizationService =
                iOrganizationServiceFactory.CreateOrganizationService(iPluginExecutionContext.UserId);

            try
            {
                if (Plugin.ValidateTargetAsEntity("ink_resource", iPluginExecutionContext))
                {
                    Entity resourceEntity =
                        (Entity)iPluginExecutionContext.InputParameters["Target"];

                    if (resourceEntity != null)
                    {
                        // ===== Logic Start =====

                        string firstName =
                            Plugin.GetAttributeValue<string>(resourceEntity, "ink_firstname");

                        string lastName =
                            Plugin.GetAttributeValue<string>(resourceEntity, "ink_lastname");

                        EntityReference reportingManager =
                            Plugin.GetAttributeValue<EntityReference>(resourceEntity, "ink_reportingmanager");

                        EntityReference designation =
                            Plugin.GetAttributeValue<EntityReference>(resourceEntity, "ink_designation");

                        QueryExpression query = new QueryExpression("ink_resource");
                        query.ColumnSet = new ColumnSet(false);
                        query.TopCount = 1;

                        if (!string.IsNullOrWhiteSpace(firstName))
                            query.Criteria.AddCondition("ink_firstname", ConditionOperator.Equal, firstName);

                        if (!string.IsNullOrWhiteSpace(lastName))
                            query.Criteria.AddCondition("ink_lastname", ConditionOperator.Equal, lastName);

                        if (reportingManager != null)
                            query.Criteria.AddCondition("ink_reportingmanager", ConditionOperator.Equal, reportingManager.Id);

                        if (designation != null)
                            query.Criteria.AddCondition("ink_designation", ConditionOperator.Equal, designation.Id);

                        EntityCollection existingRecords =
                            iOrganizationService.RetrieveMultiple(query);

                        if (existingRecords != null && existingRecords.Entities.Count > 0)
                        {
                            Plugin.ThrowManualException("Duplicate Resource record already exists.");
                        }

                        // ===== Logic End =====
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