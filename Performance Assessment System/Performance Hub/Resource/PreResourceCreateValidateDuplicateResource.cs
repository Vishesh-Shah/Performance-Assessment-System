using Inkey.MSCRM.Plugin_V9._0.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace JeetPlugins.Resource
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
                if (Plugin.ValidateTargetAsEntity("ink_resource", iPluginExecutionContext))
                {
                    Entity resourceEntity =
                        (Entity)iPluginExecutionContext.InputParameters["Target"];

                    if (resourceEntity != null)
                    {

                        string firstName =
                            Plugin.GetAttributeValue<string>(resourceEntity, "ink_firstname");

                        string lastName =
                            Plugin.GetAttributeValue<string>(resourceEntity, "ink_lastname");

                        EntityReference managerRef =
                            Plugin.GetAttributeValue<EntityReference>(resourceEntity, "ink_reportingmanager");

                        if (!string.IsNullOrWhiteSpace(firstName) &&
                            !string.IsNullOrWhiteSpace(lastName) &&
                            managerRef != null)
                        {
                            QueryExpression query = new QueryExpression("ink_resource");
                            query.ColumnSet = new ColumnSet(false);

                            query.Criteria.AddCondition("ink_firstname", ConditionOperator.Equal, firstName);
                            query.Criteria.AddCondition("ink_lastname", ConditionOperator.Equal, lastName);
                            query.Criteria.AddCondition("ink_reportingmanager", ConditionOperator.Equal, managerRef.Id);

                            EntityCollection result =
                                iOrganizationService.RetrieveMultiple(query);

                            if (result != null && result.Entities.Count > 0)
                            {
                                throw new InvalidPluginExecutionException("Duplicate Resource record found with same First Name, Last Name and Reporting Manager.");
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

        #endregion
    }
}