using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Inkey.MSCRM.Plugin_V9._0.Common;

namespace Performance_Assessment_System.Client_Matrix.Project
{
    public class PostProjectCreatSetNextRetroDate: IPlugin
    {
        public void Execute(IServiceProvider iServiceProvider)
        {
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
                tracingService.Trace("PostProjectCreateSetNextRetroDate plugin execution started.");

                if (Plugin.ValidateTargetAsEntity("ink_project", context))
                {
                    Entity target = (Entity)context.InputParameters["Target"];

                    if (target != null)
                    {
                        // Retrieve full project record with required fields using helper method
                        Entity projectEntity = Plugin.FetchEntityRecord("ink_project", target.Id,
                            new ColumnSet("createdon", "ink_retrofrequency"), service);

                        if (projectEntity != null)
                        {
                            // ink_retrofrequency is Whole Number (int) - use GetAttributeValue<int>
                            int frequencyDays = Plugin.GetAttributeValue<int>(projectEntity, "ink_retrofrequency");

                            if (frequencyDays > 0)
                            {
                                // Get created on date using helper method
                                DateTime createdOn = Plugin.GetAttributeValue<DateTime>(projectEntity, "createdon");

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
                                Entity projectUpdateEntity = new Entity("ink_project");
                                projectUpdateEntity.Id = projectEntity.Id;
                                Plugin.AddAttribute(projectUpdateEntity, "ink_nextretrodate", nextRetroDate);
                                service.Update(projectUpdateEntity);
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
