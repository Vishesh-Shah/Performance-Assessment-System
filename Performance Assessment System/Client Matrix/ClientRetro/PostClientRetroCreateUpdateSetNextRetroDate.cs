using Inkey.MSCRM.Plugin_V9._0.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Performance_Assessment_System.Client_Matrix.ClientRetro
{
    public class PostClientRetroCreateUpdateSetNextRetroDate:IPlugin
    {
        #region Variable Declaration

        private const string preImageAlias = "PostClientRetroCreateUpdateSetNextRetroDatePreImage";

        #endregion

        #region Public Methods

        #region Execute
        /// <summary>
        /// On Client Retro Create or Update, gets conducted on date from target or pre image,
        /// retrieves related client frequency days and updates Next Retro Date on Client record.
        /// </summary>
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
                if (Plugin.ValidateTargetAsEntity("ink_clientretro", iPluginExecutionContext))
                {
                    Entity clientRetroEntity = (Entity)iPluginExecutionContext.InputParameters["Target"];

                    // Get pre image for Update - to get conducted on and client lookup if not in target
                    Entity clientRetroPreImage = Plugin.GetPreEntityImage(iPluginExecutionContext, preImageAlias);

                    if (clientRetroEntity != null)
                    {
                        // Get conducted on from target or pre image using helper method
                        DateTime conductedOn = Plugin.GetAttributeValue<DateTime>(clientRetroEntity, clientRetroPreImage, "ink_conductedon");

                        // Proceed only if conducted on is set
                        if (conductedOn != DateTime.MinValue)
                        {
                            // Get client lookup from target or pre image using helper method
                            EntityReference clientEntityReference = Plugin.GetAttributeValue<EntityReference>(clientRetroEntity, clientRetroPreImage, "ink_client");

                            if (clientEntityReference != null)
                            {
                                // Retrieve client record to get current frequency days
                                Entity clientEntity = Plugin.FetchEntityRecord("ink_client", clientEntityReference.Id,
                                    new ColumnSet("ink_retrofrequency"), iOrganizationService);

                                if (clientEntity != null)
                                {
                                    // Get retro frequency days from client record
                                    int frequencyDays = Plugin.GetAttributeValue<int>(clientEntity, "ink_retrofrequency");

                                    if (frequencyDays > 0)
                                    {
                                        // Calculate next retro date = conducted on + frequency days
                                        DateTime nextRetroDate = conductedOn.Date.AddDays(frequencyDays);

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

                                        // Update next retro date on client record
                                        Entity clientUpdateEntity = new Entity("ink_client");
                                        clientUpdateEntity.Id = clientEntityReference.Id;
                                        Plugin.AddAttribute(clientUpdateEntity, "ink_nextretrodate", nextRetroDate);
                                        iOrganizationService.Update(clientUpdateEntity);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (InvalidPluginExecutionException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Plugin.TraceLog("Error: " + ex.Message, iTracingService);
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        #endregion

        #endregion
    }
}
