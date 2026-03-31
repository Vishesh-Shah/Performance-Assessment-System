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
                    Entity targetEntity =
                        (Entity)iPluginExecutionContext.InputParameters["Target"];

                    Entity preImage =
                        Plugin.GetPreEntityImage(iPluginExecutionContext, preImageAlias);

                    if (targetEntity != null)
                    {
                        string firstName =
                            Plugin.GetAttributeValue<string>(targetEntity, preImage, "ink_firstname");

                        string lastName =
                            Plugin.GetAttributeValue<string>(targetEntity, preImage, "ink_lastname");

                        EntityReference managerRef =
                            Plugin.GetAttributeValue<EntityReference>(targetEntity, preImage, "ink_reportingmanager");

                        EntityReference designationRef =
                            Plugin.GetAttributeValue<EntityReference>(targetEntity, preImage, "ink_designation");

                        if (!string.IsNullOrWhiteSpace(firstName) &&
                            !string.IsNullOrWhiteSpace(lastName) &&
                            managerRef != null &&
                            designationRef != null) 
                        {
                            QueryExpression query = new QueryExpression("ink_resource");
                            query.ColumnSet = new ColumnSet(false);

                            query.Criteria.AddCondition("ink_firstname", ConditionOperator.Equal, firstName);
                            query.Criteria.AddCondition("ink_lastname", ConditionOperator.Equal, lastName);
                            query.Criteria.AddCondition("ink_reportingmanager", ConditionOperator.Equal, managerRef.Id);

                            query.Criteria.AddCondition("ink_designation", ConditionOperator.Equal, designationRef.Id);

                            query.Criteria.AddCondition("ink_resourceid", ConditionOperator.NotEqual, targetEntity.Id);

                            EntityCollection result =
                                iOrganizationService.RetrieveMultiple(query);

                            if (result != null && result.Entities.Count > 0)
                            {
                                throw new InvalidPluginExecutionException("Duplicate Resource record found with same First Name, Last Name, Reporting Manager and Designation.");
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