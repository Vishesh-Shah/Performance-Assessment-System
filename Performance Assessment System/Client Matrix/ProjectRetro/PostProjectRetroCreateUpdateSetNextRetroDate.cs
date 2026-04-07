using Inkey.MSCRM.Plugin_V9._0.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Performance_Assessment_System.Common;
using System;

namespace Performance_Assessment_System.Client_Matrix.ProjectRetro
{
    public class PostProjectRetroCreateUpdateSetNextRetroDate: IPlugin
    {
        #region Variable Declaration

        private const string preImageAlias = "PostProjectRetroCreateUpdateSetNextRetroDatePreImage";

        #endregion

        #region Public Methods

        #region Execute
        /// <summary>
        /// On Project Retro Create or Update, gets conducted on date from target or pre image,
        /// retrieves related project frequency days and updates Next Retro Date on Project record.
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
                if (Plugin.ValidateTargetAsEntity(CommonEntities.PROJECTRETRO, iPluginExecutionContext))
                {
                    Entity projectRetroEntity = (Entity)iPluginExecutionContext.InputParameters["Target"];

                    // Get pre image for Update - to get conducted on and project lookup if not in target
                    Entity projectRetroPreImage = Plugin.GetPreEntityImage(iPluginExecutionContext, preImageAlias);

                    if (projectRetroEntity != null)
                    {
                        // Get conducted on from target or pre image using helper method
                        DateTime conductedOn = Plugin.GetAttributeValue<DateTime>(projectRetroEntity, projectRetroPreImage, "ink_conductedon");

                        // Proceed only if conducted on is set
                        if (conductedOn != DateTime.MinValue)
                        {
                            // Get project lookup from target or pre image using helper method
                            EntityReference projectEntityReference = Plugin.GetAttributeValue<EntityReference>(projectRetroEntity, projectRetroPreImage, "ink_project");

                            if (projectEntityReference != null)
                            {
                                // Retrieve project record to get current frequency days
                                Entity projectEntity = Plugin.FetchEntityRecord("ink_project", projectEntityReference.Id,
                                    new ColumnSet("ink_retrofrequency"), iOrganizationService);

                                if (projectEntity != null)
                                {
                                    // Get retro frequency days from project record
                                    int frequencyDays = Plugin.GetAttributeValue<int>(projectEntity, "ink_retrofrequency");

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

                                        // Update next retro date on project record
                                        Entity projectUpdateEntity = new Entity(CommonEntities.PROJECT);
                                        projectUpdateEntity.Id = projectEntityReference.Id;
                                        Plugin.AddAttribute(projectUpdateEntity, "ink_nextretrodate", nextRetroDate);
                                        iOrganizationService.Update(projectUpdateEntity);
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