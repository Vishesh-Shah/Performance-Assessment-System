using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Inkey.MSCRM.Plugin_V9._0.Common;
using Performance_Assessment_System.Common;

namespace Performance_Assessment_System.Client_Matrix.Project
{
    public class PostProjectCreatSetNextRetroDate: IPlugin
    {
        public void Execute(IServiceProvider iServiceProvider)
        {
            // Obtain the execution context from the service provider.
            IPluginExecutionContext iPluginExecutionContext = (IPluginExecutionContext)iServiceProvider.GetService(typeof(IPluginExecutionContext));

            // Obtain the organization service factory reference.
            IOrganizationServiceFactory iOrganizationServiceFactory = (IOrganizationServiceFactory)iServiceProvider.GetService(typeof(IOrganizationServiceFactory));

            // Obtain the tracing service reference.
            ITracingService iTracingService = (ITracingService)iServiceProvider.GetService(typeof(ITracingService));

            // Obtain the organization service reference.
            IOrganizationService iOrganizationService = iOrganizationServiceFactory.CreateOrganizationService(iPluginExecutionContext.UserId);

            try
            {
               

                if (Plugin.ValidateTargetAsEntity(CommonEntities.PROJECT, iPluginExecutionContext))
                {
                    Entity projectEntity = (Entity)iPluginExecutionContext.InputParameters["Target"];

                    if (projectEntity != null)
                    {
                        // Retrieve full project record with required fields using helper method
                        Entity fullProjectEntity = Plugin.FetchEntityRecord(projectEntity.LogicalName, projectEntity.Id,
                            new ColumnSet("createdon", "ink_retrofrequency"), iOrganizationService);

                        if (fullProjectEntity != null)
                        {
                            // ink_retrofrequency is Whole Number (int) - use GetAttributeValue<int>
                            int frequencyDays = Plugin.GetAttributeValue<int>(fullProjectEntity, "ink_retrofrequency");

                            if (frequencyDays > 0)
                            {
                                // Get created on date using helper method
                                DateTime createdOn = Plugin.GetAttributeValue<DateTime>(fullProjectEntity, "createdon");

                                // Step 1: Calculate next retro date = created on + frequency days
                                DateTime nextRetroDate = createdOn.AddDays(frequencyDays);

                                // Step 2: Apply weekend adjustment
                                if (nextRetroDate.DayOfWeek == DayOfWeek.Saturday)
                                {
                                    // Saturday → move to Monday
                                    nextRetroDate = nextRetroDate.AddDays(2);
                                }
                                else if (nextRetroDate.DayOfWeek == DayOfWeek.Sunday)
                                {
                                    // Sunday → move to Monday
                                    nextRetroDate = nextRetroDate.AddDays(1);
                                }

                                // Update next retro date on project record using helper method
                                Entity projectUpdateEntity = new Entity(CommonEntities.PROJECT);
                                projectUpdateEntity.Id = fullProjectEntity.Id;
                                Plugin.AddAttribute(projectUpdateEntity, "ink_nextretrodate", nextRetroDate);
                                iOrganizationService.Update(projectUpdateEntity);
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
    }
}
