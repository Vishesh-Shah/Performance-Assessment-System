using Inkey.MSCRM.Plugin_V9._0.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Performance_Assessment_System.Resource
{
    public class PreResourceUpdateValidateDuplicateResource : IPlugin
    {
        #region Constants
        private const string preImageAlias = "PreResourceUpdateValidateDuplicateResourcePreImage";
        #endregion

        #region Execute
        public void Execute(IServiceProvider iServiceProvider)
        {
            IPluginExecutionContext context =
                (IPluginExecutionContext)iServiceProvider.GetService(typeof(IPluginExecutionContext));

            IOrganizationServiceFactory factory =
                (IOrganizationServiceFactory)iServiceProvider.GetService(typeof(IOrganizationServiceFactory));

            ITracingService trace =
                (ITracingService)iServiceProvider.GetService(typeof(ITracingService));

            IOrganizationService service =
                factory.CreateOrganizationService(context.UserId);

            try
            {
                trace.Trace("Plugin Start: PreResourceUpdateValidateDuplicateResource");

                if (!Plugin.ValidateTargetAsEntity("ink_resource", context))
                    return;

                Entity target = (Entity)context.InputParameters["Target"];

                Entity preImage = null;
                if (context.PreEntityImages.Contains(preImageAlias))
                {
                    preImage = context.PreEntityImages[preImageAlias];
                }

                string firstName =
                    Plugin.GetAttributeValue<string>(target, preImage, "ink_firstname");

                string lastName =
                    Plugin.GetAttributeValue<string>(target, preImage, "ink_lastname");

                EntityReference managerRef =
                    Plugin.GetAttributeValue<EntityReference>(target, preImage, "ink_reportingmanager");

                EntityReference designationRef =
                    Plugin.GetAttributeValue<EntityReference>(target, preImage, "ink_designation");

                trace.Trace($"FirstName: {firstName}");
                trace.Trace($"LastName: {lastName}");
                trace.Trace($"Manager: {(managerRef != null ? managerRef.Id.ToString() : "NULL")}");
                trace.Trace($"Designation: {(designationRef != null ? designationRef.Id.ToString() : "NULL")}");

                if (string.IsNullOrWhiteSpace(firstName) ||
                    string.IsNullOrWhiteSpace(lastName) ||
                    managerRef == null ||
                    designationRef == null)
                {
                    trace.Trace("Required fields missing → exit");
                    return;
                }

                QueryExpression query = new QueryExpression("ink_resource");
                query.ColumnSet = new ColumnSet(false);

                query.Criteria.AddCondition("ink_firstname", ConditionOperator.Equal, firstName);
                query.Criteria.AddCondition("ink_lastname", ConditionOperator.Equal, lastName);
                query.Criteria.AddCondition("ink_reportingmanager", ConditionOperator.Equal, managerRef.Id);
                query.Criteria.AddCondition("ink_designation", ConditionOperator.Equal, designationRef.Id);

                // 🔥 Correct fix
                query.Criteria.AddCondition("ink_resourceid", ConditionOperator.NotEqual, context.PrimaryEntityId);

                trace.Trace("Executing query...");

                EntityCollection result = service.RetrieveMultiple(query);

                trace.Trace("Duplicate Count: " + result.Entities.Count);

                if (result.Entities.Count > 0)
                {
                    throw new InvalidPluginExecutionException(
                        "Duplicate Resource not allowed! Same First Name, Last Name, Reporting Manager and Designation already exists.");
                }

                trace.Trace("Plugin End Successfully");
            }
            catch (Exception ex)
            {
                trace.Trace("Error: " + ex.ToString());
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        #endregion
    }
}