using Inkey.MSCRM.Plugin_V9._0.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Performance_Assessment_System.Common;
using System;
using System.Collections.Generic;
using System.Web.Services.Description;
using System.Windows.Documents;
using System.Xml.Linq;

namespace Performance_Assessment_System.Performance_Hub.Performance_Evaluation
{
    /// <summary>
    /// PreValidation-Operation Plugin on Create of Resource.
    /// Prevents duplicate Resource records based on First Name + Last Name + Reporting Manager + Designation.
    /// Filter Attribute: N/A
    /// Pre-Image Alias: N/A
    /// </summary>
    public class PostPerformanceEvaluationCreateCreateRatingRecords : IPlugin
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
                if (Plugin.ValidateTargetAsEntity(CommonEntities.PERFORMANCEEVALUATION, iPluginExecutionContext))
                {
                    Entity performanceEvaluationEntity = (Entity)iPluginExecutionContext.InputParameters["Target"];

                    if (performanceEvaluationEntity != null)
                    {
                        EntityReference assesseeReference = Plugin.GetAttributeValue<EntityReference>(performanceEvaluationEntity, "ink_assessee");
                        Entity assesseeEntity = Plugin.FetchEntityRecord(CommonEntities.RESOURCE, assesseeReference.Id, new ColumnSet("ink_designation"), iOrganizationService);
                        OptionSetValue designation = Plugin.GetAttributeValue<OptionSetValue>(assesseeEntity, "ink_designation");

                        EntityReference reviewCycleReference = Plugin.GetAttributeValue<EntityReference>(performanceEvaluationEntity, "ink_performancereviewcycle");

                        if (reviewCycleReference != null && assesseeEntity != null && designation != null)
                        {

                            QueryExpression query = new QueryExpression(CommonEntities.PERFORMANCEEVALUATIONTEMPLATE);
                            query.ColumnSet.AddColumns("ink_name");
                            query.Criteria.AddCondition("ink_designation", ConditionOperator.Equal, designation.Value);
                            query.Criteria.AddCondition("ink_performancereviewcycle", ConditionOperator.Equal, reviewCycleReference.Id);
                            EntityCollection fetchedTemplate = iOrganizationService.RetrieveMultiple(query);

                            Entity templateEntity = null;
                            if (fetchedTemplate.Entities.Count > 0)
                            {
                                foreach (Entity template in fetchedTemplate.Entities)
                                {
                                    templateEntity = template;
                                }

                                #region QueryExpression to get Objectives

                                QueryExpression templateQuery = new QueryExpression(CommonEntities.OBJECTIVE);
                                templateQuery.ColumnSet = new ColumnSet("ink_name", "ink_objectivesid");

                                // Link to intersect table
                                LinkEntity linkIntersect = new LinkEntity(
                                    CommonEntities.OBJECTIVE,
                                    "ink_objectives_ink_performanceevaluatio",
                                    "ink_objectivesid",
                                    "ink_objectivesid",
                                    JoinOperator.Inner
                                );

                                // Link intersect to template
                                LinkEntity linkTemplate = new LinkEntity(
                                    "ink_objectives_ink_performanceevaluatio",
                                    CommonEntities.PERFORMANCEEVALUATIONTEMPLATE,
                                    "ink_performanceevaluationtemplateid",
                                    "ink_performanceevaluationtemplateid",
                                    JoinOperator.Inner
                                );

                                // Filter using Entity Id
                                linkTemplate.LinkCriteria.AddCondition(
                                    "ink_performanceevaluationtemplateid",
                                    ConditionOperator.Equal,
                                    templateEntity.Id
                                );

                                // Chain
                                linkIntersect.LinkEntities.Add(linkTemplate);
                                templateQuery.LinkEntities.Add(linkIntersect);

                                EntityCollection fetchedObjectives = iOrganizationService.RetrieveMultiple(templateQuery);

                                #endregion

                                bool flag = true;
                                int objectiveIndex = 0;
                                decimal keyResultIndex = 0m;
                                if (fetchedObjectives.Entities.Count > 0)
                                {
                                    foreach (Entity objectiveEntity in fetchedObjectives.Entities)
                                    {
                                        flag = true;
                                        keyResultIndex = 0;
                                        objectiveIndex += 1;

                                        string objName = Plugin.GetAttributeValue<string>(objectiveEntity, "ink_name");

                                        QueryExpression keyResultQuery = new QueryExpression(CommonEntities.KEYRESULT);
                                        keyResultQuery.ColumnSet.AddColumns("ink_name");
                                        keyResultQuery.Criteria.AddCondition("ink_objectives", ConditionOperator.Equal, objectiveEntity.Id);

                                        EntityCollection fetchedKeyResults = iOrganizationService.RetrieveMultiple(keyResultQuery);

                                        keyResultIndex += objectiveIndex;
                                        string PEName = Plugin.GetAttributeValue<string>(performanceEvaluationEntity, "ink_name");

                                        if (fetchedKeyResults.Entities.Count > 0)
                                        {
                                            for (int i = 0; i < (fetchedKeyResults.Entities.Count + 1); i++)
                                            {
                                                if (flag == true)
                                                {
                                                    string name = PEName + " " + objName;
                                                    string SrNo = objectiveIndex + " - Objective";
                                                    EntityReference performanceEvaluation = new EntityReference(CommonEntities.PERFORMANCEEVALUATION, performanceEvaluationEntity.Id);
                                                    EntityReference keyResult = new EntityReference(CommonEntities.KEYRESULT, fetchedKeyResults.Entities[i].Id);
                                                    EntityReference objective = new EntityReference(CommonEntities.OBJECTIVE, objectiveEntity.Id);

                                                    Entity keyResultRatingEntity = new Entity(CommonEntities.KEYRESULTRATING);
                                                    Plugin.AddAttribute<string>(keyResultRatingEntity, "ink_name", name);
                                                    Plugin.AddAttribute<EntityReference>(keyResultRatingEntity, "ink_performanceevaluations", performanceEvaluation);
                                                    Plugin.AddAttribute<EntityReference>(keyResultRatingEntity, "ink_keyresults", keyResult);
                                                    Plugin.AddAttribute<EntityReference>(keyResultRatingEntity, "ink_objectives", objective);
                                                    Plugin.AddAttribute<int>(keyResultRatingEntity, "ink_objectivenumbering", objectiveIndex);
                                                    Plugin.AddAttribute<string>(keyResultRatingEntity, "ink_griddisplayname", objName);
                                                    Plugin.AddAttribute<string>(keyResultRatingEntity, "ink_srno", SrNo);
                                                    iOrganizationService.Create(keyResultRatingEntity);

                                                    flag = false;
                                                }
                                                else
                                                {
                                                    Entity keyResultEntity = fetchedKeyResults.Entities[i - 1];
                                                    string KRname = Plugin.GetAttributeValue<string>(keyResultEntity, "ink_name");
                                                    string name = PEName + " " + KRname;
                                                    keyResultIndex += 0.1m;
                                                    string SrNo = keyResultIndex.ToString();

                                                    EntityReference performanceEvaluation = new EntityReference(CommonEntities.PERFORMANCEEVALUATION, performanceEvaluationEntity.Id);
                                                    EntityReference keyResult = new EntityReference(CommonEntities.KEYRESULT, keyResultEntity.Id);
                                                    EntityReference objective = new EntityReference(CommonEntities.OBJECTIVE, objectiveEntity.Id);

                                                    Entity keyResultRatingEntity = new Entity(CommonEntities.KEYRESULTRATING);
                                                    Plugin.AddAttribute<string>(keyResultRatingEntity, "ink_name", name);
                                                    Plugin.AddAttribute<EntityReference>(keyResultRatingEntity, "ink_performanceevaluations", performanceEvaluation);
                                                    Plugin.AddAttribute<EntityReference>(keyResultRatingEntity, "ink_keyresults", keyResult);
                                                    Plugin.AddAttribute<EntityReference>(keyResultRatingEntity, "ink_objectives", objective);
                                                    Plugin.AddAttribute<decimal>(keyResultRatingEntity, "ink_keyresultnumbering", keyResultIndex);
                                                    Plugin.AddAttribute<string>(keyResultRatingEntity, "ink_griddisplayname", KRname);
                                                    Plugin.AddAttribute<string>(keyResultRatingEntity, "ink_srno", SrNo);
                                                    iOrganizationService.Create(keyResultRatingEntity);

                                                }
                                            }
                                        }

                                    }
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