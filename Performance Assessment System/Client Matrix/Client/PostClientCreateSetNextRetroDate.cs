using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Inkey.MSCRM.Plugin_V9._0.Common;
using Performance_Assessment_System.Common;


namespace Performance_Assessment_System.Client_Matrix.Client
{
    public class PostClientCreateSetNextRetroDate : IPlugin
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
                if (Plugin.ValidateTargetAsEntity(CommonEntities.CLIENT, iPluginExecutionContext))
                {
                    Entity clientEntity = (Entity)iPluginExecutionContext.InputParameters["Target"];
                    if (clientEntity != null)
                    {
                        int frequencyDays = Plugin.GetAttributeValue<int>(clientEntity, "ink_retrofrequency");
                        if (frequencyDays > 0)
                        {
                            DateTime createdOn = Plugin.GetAttributeValue<DateTime>(clientEntity, "createdon");

                            // Calculate next retro date = created on + frequency days
                            DateTime nextRetroDate = createdOn.AddDays(frequencyDays);

                            // Apply weekend adjustment
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
                            Entity clientUpdateEntity = new Entity(CommonEntities.CLIENT, clientEntity.Id);
                            clientUpdateEntity["ink_nextretrodate"] = nextRetroDate;
                            iOrganizationService.Update(clientUpdateEntity);
                        }
                        else
                        {
                            throw new Exception("Retro Frequency must be greater than 0");
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