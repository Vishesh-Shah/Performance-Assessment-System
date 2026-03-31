using Inkey.MSCRM.Plugin_V9._0.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Performance_Assessment_System.Resource
{
    public class PreResourceCreateValidateDuplicateResource : IPlugin
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
                iTracingService.Trace("Plugin Execution Started: PreResourceCreateValidateDuplicateResource");

                if (Plugin.ValidateTargetAsEntity("ink_resource", iPluginExecutionContext))
                {
                    iTracingService.Trace("Target is valid entity");

                    Entity resourceEntity =
                        (Entity)iPluginExecutionContext.InputParameters["Target"];

                    if (resourceEntity != null)
                    {
                        iTracingService.Trace("Resource Entity is not null");

                        string firstName =
                            Plugin.GetAttributeValue<string>(resourceEntity, "ink_firstname");

                        string lastName =
                            Plugin.GetAttributeValue<string>(resourceEntity, "ink_lastname");

                        EntityReference managerRef =
                            Plugin.GetAttributeValue<EntityReference>(resourceEntity, "ink_reportingmanager");

                        EntityReference designationRef =
                            Plugin.GetAttributeValue<EntityReference>(resourceEntity, "ink_designation");

                        iTracingService.Trace($"FirstName: {firstName}");
                        iTracingService.Trace($"LastName: {lastName}");
                        iTracingService.Trace($"ManagerRef: {(managerRef != null ? managerRef.Id.ToString() : "NULL")}");
                        iTracingService.Trace($"DesignationRef: {(designationRef != null ? designationRef.Id.ToString() : "NULL")}");

                        if (!string.IsNullOrWhiteSpace(firstName) &&
                            !string.IsNullOrWhiteSpace(lastName) &&
                            managerRef != null &&
                            designationRef != null)
                        {
                            iTracingService.Trace("All required fields are present. Building query...");

                            QueryExpression query = new QueryExpression("ink_resource");
                            query.ColumnSet = new ColumnSet(false);

                            query.Criteria.AddCondition("ink_firstname", ConditionOperator.Equal, firstName);
                            query.Criteria.AddCondition("ink_lastname", ConditionOperator.Equal, lastName);
                            query.Criteria.AddCondition("ink_reportingmanager", ConditionOperator.Equal, managerRef.Id);
                            query.Criteria.AddCondition("ink_designation", ConditionOperator.Equal, designationRef.Id);

                            iTracingService.Trace("Executing RetrieveMultiple...");

                            EntityCollection result =
                                iOrganizationService.RetrieveMultiple(query);

                            iTracingService.Trace($"Records Found: {result.Entities.Count}");

                            if (result != null && result.Entities.Count > 0)
                            {
                                iTracingService.Trace("Duplicate record found. Throwing exception.");

                                throw new InvalidPluginExecutionException(
                                    "Duplicate Resource record found with same First Name, Last Name, Reporting Manager and Designation.");
                            }
                        }
                        else
                        {
                            iTracingService.Trace("Required fields are missing. Skipping duplicate check.");
                        }
                    }
                    else
                    {
                        iTracingService.Trace("Resource Entity is NULL");
                    }
                }
                else
                {
                    iTracingService.Trace("Target is NOT valid entity");
                }

                iTracingService.Trace("Plugin Execution Completed Successfully");
            }
            catch (Exception ex)
            {
                iTracingService.Trace("Exception Occurred: " + ex.ToString());
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        #endregion

        #endregion
    }
}