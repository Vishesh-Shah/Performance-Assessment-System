using Inkey.MSCRM.Plugin_V9._0.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Performance_Assessment_System.Resource
{
    public class PreResourceUpdateValidateDuplicate : IPlugin
    {
        #region Constants

        private const string preImageAlias = "PreResourceUpdateValidateDuplicatePreImage";

        #endregion

        #region Public Methods

        #region Execute
        public void Execute(IServiceProvider iServiceProvider)
        {
            IPluginExecutionContext context =
                (IPluginExecutionContext)iServiceProvider.GetService(typeof(IPluginExecutionContext));

            IOrganizationServiceFactory factory =
                (IOrganizationServiceFactory)iServiceProvider.GetService(typeof(IOrganizationServiceFactory));

            IOrganizationService service =
                factory.CreateOrganizationService(context.UserId);

            try
            {
                if (Plugin.ValidateTargetAsEntity("ink_resource", context))
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    Entity preImage = Plugin.GetPreEntityImage(context, preImageAlias);

                    if (entity != null)
                    {
                        string firstName = Plugin.GetAttributeValue<string>(entity, preImage, "ink_firstname");
                        string lastName = Plugin.GetAttributeValue<string>(entity, preImage, "ink_lastname");
                        EntityReference manager = Plugin.GetAttributeValue<EntityReference>(entity, preImage, "ink_reportingmanager");
                        EntityReference designation = Plugin.GetAttributeValue<EntityReference>(entity, preImage, "ink_designation");

                        QueryExpression query = new QueryExpression("ink_resource");
                        query.ColumnSet = new ColumnSet(false);
                        query.TopCount = 1;

                        // 🔥 STRICT MATCH
                        query.Criteria.AddCondition("ink_firstname", ConditionOperator.Equal, firstName);
                        query.Criteria.AddCondition("ink_lastname", ConditionOperator.Equal, lastName);

                        if (manager != null)
                            query.Criteria.AddCondition("ink_reportingmanager", ConditionOperator.Equal, manager.Id);

                        if (designation != null)
                            query.Criteria.AddCondition("ink_designation", ConditionOperator.Equal, designation.Id);

                        // ✅ Exclude current record
                        query.Criteria.AddCondition("ink_resourceid", ConditionOperator.NotEqual, entity.Id);

                        EntityCollection result = service.RetrieveMultiple(query);

                        if (result != null && result.Entities.Count > 0)
                        {
                            Plugin.ThrowManualException("Duplicate Resource record already exists.");
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

        #endregion
    }
}