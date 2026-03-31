using Inkey.MSCRM.Plugin_V9._0.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Performance_Assessment_System.Common;
using System;

namespace Performance_Assessment_System.Performance_Hub.Resource
{
    /// <summary>
    /// PreValidation-Operation Plugin on Create of Resource.
    /// Prevents duplicate Resource records based on First Name + Last Name + Reporting Manager + Designation.
    /// Filter Attribute: N/A
    /// Pre-Image Alias: N/A
    /// </summary>
    public class PreResourceCreateValidateDuplicateResource : IPlugin
    {
        #region Public Methods

        #region Execute
        /// <summary>
        /// Validates whether another active Resource already exists with the same
        /// First Name, Last Name, Reporting Manager, and Designation combination.
        /// </summary>
        public void Execute(IServiceProvider iServiceProvider)
        {
            // Obtain the execution context from the service provider.
            IPluginExecutionContext iPluginExecutionContext = (IPluginExecutionContext)iServiceProvider
                .GetService(typeof(IPluginExecutionContext));

            // Obtain the organization service factory reference.
            IOrganizationServiceFactory iOrganizationServiceFactory = (IOrganizationServiceFactory)iServiceProvider
                .GetService(typeof(IOrganizationServiceFactory));

            // Obtain the tracing service reference.
            ITracingService iTracingService = (ITracingService)iServiceProvider
                .GetService(typeof(ITracingService));

            // Obtain the organization service reference.
            IOrganizationService iOrganizationService = iOrganizationServiceFactory
                .CreateOrganizationService(iPluginExecutionContext.UserId);

            try
            {
                if (Plugin.ValidateTargetAsEntity(CommonEntities.RESOURCE, iPluginExecutionContext))
                {
                    Entity resourceEntity = (Entity)iPluginExecutionContext.InputParameters["Target"];

                    if (resourceEntity != null)
                    {
                        string firstName = Plugin.GetAttributeValue<string>(
                            resourceEntity,
                            CommonEntities.Resource.INK_FIRSTNAME);

                        string lastName = Plugin.GetAttributeValue<string>(
                            resourceEntity,
                            CommonEntities.Resource.INK_LASTNAME);

                        EntityReference reportingManagerEntityReference = Plugin.GetAttributeValue<EntityReference>(
                            resourceEntity,
                            CommonEntities.Resource.INK_REPORTINGMANAGER);

                        OptionSetValue designationOptionSetValue = Plugin.GetAttributeValue<OptionSetValue>(
                            resourceEntity,
                            CommonEntities.Resource.INK_DESIGNATION);

                        firstName = firstName == null ? string.Empty : firstName.Trim();
                        lastName = lastName == null ? string.Empty : lastName.Trim();

                        if (string.IsNullOrWhiteSpace(firstName) ||
                            string.IsNullOrWhiteSpace(lastName) ||
                            reportingManagerEntityReference == null ||
                            designationOptionSetValue == null)
                        {
                            return;
                        }

                        QueryExpression resourceQueryExpression = new QueryExpression(CommonEntities.RESOURCE);
                        resourceQueryExpression.ColumnSet = new ColumnSet(CommonEntities.Resource.INK_FIRSTNAME);
                        resourceQueryExpression.Criteria.AddCondition(
                            CommonEntities.Resource.INK_FIRSTNAME,
                            ConditionOperator.Equal,
                            firstName);
                        resourceQueryExpression.Criteria.AddCondition(
                            CommonEntities.Resource.INK_LASTNAME,
                            ConditionOperator.Equal,
                            lastName);
                        resourceQueryExpression.Criteria.AddCondition(
                            CommonEntities.Resource.INK_REPORTINGMANAGER,
                            ConditionOperator.Equal,
                            reportingManagerEntityReference.Id);
                        resourceQueryExpression.Criteria.AddCondition(
                            CommonEntities.Resource.INK_DESIGNATION,
                            ConditionOperator.Equal,
                            designationOptionSetValue.Value);
                        resourceQueryExpression.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        resourceQueryExpression.TopCount = 1;

                        EntityCollection lstResourceRecords = iOrganizationService.RetrieveMultiple(resourceQueryExpression);

                        if (lstResourceRecords != null && lstResourceRecords.Entities.Count > 0)
                        {
                            Plugin.ThrowManualException(
                                "A Resource record already exists with the same First Name, Last Name, Reporting Manager, and Designation.");
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