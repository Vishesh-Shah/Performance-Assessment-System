using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Performance_Assessment_System.Common;


namespace Performance_Assessment_System.Client_Matrix.Client
{
    public class PostClientCreatSetNextRetroDate :IPlugin
    {
        #region Public Methods

        #region Execute
        /// <summary>
        /// </summary>
        /// 
        /// 
        public void Execute(IServiceProvider iServiceProvider)
        {

            // Obtain the execution context from the service provider.
            IPluginExecutionContext iPluginExecutionContext = (IPluginExecutionContext)iServiceProvider.GetService(typeof(IPluginExecutionContext));

            // Obtain the organization service factory reference.
            IOrganizationServiceFactory iOrganizationServiceFactory = (IOrganizationServiceFactory)iServiceProvider.GetService(typeof(IPluginExecutionContext));

            // Obtain the tracing service reference.
            ITracingService iTracingService = (ITracingService)iServiceProvider.GetService(typeof(ITracingService));

            // Obtain the organization service reference.
            IOrganizationService iOrganizationService = iOrganizationServiceFactory.CreateOrganizationService(iPluginExecutionContext.UserId);

            try
            {
                iTracingService.Trace("PostClientCreatSetNextRetroDate plugin execution started.");

                if (Plugin.ValidateTargetAsEntity("ink_client", iPluginExecutionContext))
                {
                    Entity clientEntity = (Entity)iPluginExecutionContext.InputParameters["Target"];

                    if (clientEntity != null)
                    {
                        // Frequency and CreatedOn direct Target ma create time hamesha male j evu jaruri nathi
                        // etle created record retrieve kariye
                        Entity retrieveClient = iOrganizationService.Retrieve(
                            Entities.CLIENT,
                            clientEntity.Id,
                            new ColumnSet("createdon", "ink_retrofrequency", "ink_nextretrodate")
                        );

                        if (retrieveClient != null)
                        {
                            if (retrieveClient.Contains("ink_retrofrequency") && retrieveClient["ink_retrofrequency"] != null)
                            {
                                int frequency = retrieveClient.GetAttributeValue<int>("ink_retrofrequency");

                                if (frequency > 0)
                                {
                                    DateTime createdOn = retrieveClient.GetAttributeValue<DateTime>("createdon");

                                    Entity updateClient = new Entity(Entities.CLIENT);
                                    updateClient.Id = retrieveClient.Id;
                                    updateClient["ink_nextretrodate"] = createdOn.AddDays(frequency);

                                    iOrganizationService.Update(updateClient);
                                }
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
