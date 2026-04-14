using System;
using Microsoft.Xrm.Sdk;
using Inkey.MSCRM.Plugin_V9._0.Common;
using Performance_Assessment_System.Common;


namespace Performance_Assessment_System.Client_Matrix.Project
{
    public class PreProjectCreateValidateRetroFrequency : IPlugin
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
                        int frequency = Plugin.GetAttributeValue<int>(projectEntity, "ink_retrofrequency");
                        if (frequency <= 0)
                        {
                            throw new Exception("Retro Frequency must be a positive number greater than 0.");
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