using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Inkey.MSCRM.Plugin_V9._0.Common;
using Performance_Assessment_System.Common;

namespace Performance_Assessment_System.Client_Matrix.Client
{
    public class PostClientCreatSetNextRetroDate : IPlugin
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
               

                if (Plugin.ValidateTargetAsEntity(CommonEntities.CLIENT,iPluginExecutionContext))
                {
                    Entity clientEntity = (Entity)iPluginExecutionContext.InputParameters["Target"];

                    if (clientEntity != null)
                    {
                        // Retrieve full client record with required fields using helper method
                        Entity fullClientEntity = Plugin.FetchEntityRecord(clientEntity.LogicalName, clientEntity.Id,
                            new ColumnSet("createdon", "ink_retrofrequency"), iOrganizationService);

                        if (fullClientEntity != null)
                        {
                            // ink_retrofrequency is Whole Number (int) - use GetAttributeValue<int>
                            int frequencyDays = Plugin.GetAttributeValue<int>(fullClientEntity, "ink_retrofrequency");

                            if (frequencyDays > 0)
                            {
                                // Get created on date using helper method
                                DateTime createdOn = Plugin.GetAttributeValue<DateTime>(fullClientEntity, "createdon");

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

                                // Update next retro date on client record using helper method
                                Entity clientUpdateEntity = new Entity(CommonEntities.CLIENT);
                                clientUpdateEntity.Id = fullClientEntity.Id;
                                Plugin.AddAttribute(clientUpdateEntity, "ink_nextretrodate", nextRetroDate);
                                iOrganizationService.Update(clientUpdateEntity);
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