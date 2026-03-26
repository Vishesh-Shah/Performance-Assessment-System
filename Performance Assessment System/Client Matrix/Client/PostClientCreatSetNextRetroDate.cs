using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Inkey.MSCRM.Plugin_V9._0.Common;

namespace Performance_Assessment_System.Client_Matrix.Client
{
    public class PostClientCreatSetNextRetroDate : IPlugin
    {
        public void Execute(IServiceProvider iServiceProvider)
        {
            // ✅ Correct service initialization
            IPluginExecutionContext context =
                (IPluginExecutionContext)iServiceProvider.GetService(typeof(IPluginExecutionContext));

            IOrganizationServiceFactory factory =
                (IOrganizationServiceFactory)iServiceProvider.GetService(typeof(IOrganizationServiceFactory));

            ITracingService tracingService =
                (ITracingService)iServiceProvider.GetService(typeof(ITracingService));

            IOrganizationService service =
                factory.CreateOrganizationService(context.UserId);

            try
            {
                tracingService.Trace("PostClientCreatSetNextRetroDate plugin execution started.");

                if (Plugin.ValidateTargetAsEntity("ink_client", context))
                {
                    Entity target = (Entity)context.InputParameters["Target"];

                    if (target != null)
                    {
                        // Retrieve full client record with required fields using helper method
                        Entity clientEntity = Plugin.FetchEntityRecord("ink_client", target.Id,
                            new ColumnSet("createdon", "ink_retrofrequency"), service);

                        if (clientEntity != null)
                        {
                            // ink_retrofrequency is Whole Number (int) - use GetAttributeValue<int>
                            int frequencyDays = Plugin.GetAttributeValue<int>(clientEntity, "ink_retrofrequency");

                            if (frequencyDays > 0)
                            {
                                // Get created on date using helper method
                                DateTime createdOn = Plugin.GetAttributeValue<DateTime>(clientEntity, "createdon");

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
                                Entity clientUpdateEntity = new Entity("ink_client");
                                clientUpdateEntity.Id = clientEntity.Id;
                                Plugin.AddAttribute(clientUpdateEntity, "ink_nextretrodate", nextRetroDate);
                                service.Update(clientUpdateEntity);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Plugin.TraceLog("Error: " + ex.Message, tracingService);
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}