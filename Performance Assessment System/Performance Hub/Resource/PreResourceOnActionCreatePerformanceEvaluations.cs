using Inkey.MSCRM.Plugin_V9._0.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Performance_Assessment_System.Common;
using System;
using System.Collections.Generic;

namespace Performance_Assessment_System.Performance_Hub.Resource
{
   
    public class PreResourceOnActionCreatePerformanceEvaluations : IPlugin
    {
        #region Variable Declaration

        #endregion

        #region Public Methods

        #region Execute
        /// <summary>
        /// Pre-Operation Action Plugin on Create Performance Evaluations Action triggered from Resource form ribbon button.
        /// Loops all active resources, checks for duplicate Performance Evaluations (same Assessee + Assessor + Cycle),
        /// and creates Performance Evaluation records only where no duplicate exists.
        /// Duplicate check uses both in-memory HashSet (same-transaction guard) and DB query (previous-run guard).
        /// Filter Attribute: N/A (Action trigger)
        /// Pre-Image Alias: N/A (Action trigger)
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
                // ── Step 1: Fetch all active resources ───────────────────────────
                QueryExpression resourceQueryExpression = new QueryExpression(CommonEntities.RESOURCE);
                resourceQueryExpression.ColumnSet = new ColumnSet(
                    CommonEntities.Resource.INK_FIRSTNAME,
                    CommonEntities.Resource.INK_LASTNAME,
                    CommonEntities.Resource.INK_REPORTINGMANAGER,
                    CommonEntities.Resource.INK_DESIGNATION
                );
                resourceQueryExpression.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);

                EntityCollection lstResourceRecords = iOrganizationService.RetrieveMultiple(resourceQueryExpression);

                if (lstResourceRecords.Entities.Count > 0)
                {
                    // ── In-memory tracker: guards against duplicates created
                    // within the SAME plugin execution (DB query won't see them
                    // until the transaction commits, so we track them here too).
                    // Key format: assesseeId_assessorId_cycleId
                    HashSet<string> processedCombinations = new HashSet<string>();

                    foreach (Entity resourceEntity in lstResourceRecords.Entities)
                    {
                        // ── Step 2: Skip resource if no Reporting Manager ─────────
                        EntityReference reportingManagerEntityReference = Plugin.GetAttributeValue<EntityReference>(resourceEntity, CommonEntities.Resource.INK_REPORTINGMANAGER);

                        if (reportingManagerEntityReference == null)
                        {
                            Plugin.TraceLog("Skipping resource " + resourceEntity.Id + " — no Reporting Manager.", iTracingService);
                            continue;
                        }

                        // ── Step 3: Get Designation from resource ─────────────────
                        OptionSetValue designationOptionSetValue = Plugin.GetAttributeValue<OptionSetValue>(resourceEntity, CommonEntities.Resource.INK_DESIGNATION);

                        if (designationOptionSetValue == null)
                        {
                            Plugin.TraceLog("Skipping resource " + resourceEntity.Id + " — no Designation.", iTracingService);
                            continue;
                        }

                        // ── Step 4: Find the LATEST active template for this Designation
                        // TopCount = 1 + order by createdon desc ensures only ONE
                        // template is returned even if multiple exist for the same
                        // designation — preventing 3x duplicates per person.
                        QueryExpression templateQueryExpression = new QueryExpression(CommonEntities.PERFORMANCEEVALUATIONTEMPLATE);
                        templateQueryExpression.ColumnSet = new ColumnSet(
                            CommonEntities.PerformanceEvaluationTemplate.INK_PERFORMANCEREVIEWCYCLE
                        );
                        templateQueryExpression.Criteria.AddCondition(CommonEntities.PerformanceEvaluationTemplate.INK_DESIGNATION, ConditionOperator.Equal, designationOptionSetValue.Value);
                        templateQueryExpression.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        templateQueryExpression.TopCount = 1;
                        templateQueryExpression.AddOrder("createdon", OrderType.Descending);

                        EntityCollection lstTemplateRecords = iOrganizationService.RetrieveMultiple(templateQueryExpression);

                        if (lstTemplateRecords.Entities.Count == 0)
                        {
                            Plugin.TraceLog("Skipping resource " + resourceEntity.Id + " — no matching template for designation " + designationOptionSetValue.Value + ".", iTracingService);
                            continue;
                        }

                        // ── Step 5: Get Cycle from the matched template ────────────
                        Entity templateEntity = lstTemplateRecords.Entities[0];

                        EntityReference performanceReviewCycleEntityReference = Plugin.GetAttributeValue<EntityReference>(templateEntity, CommonEntities.PerformanceEvaluationTemplate.INK_PERFORMANCEREVIEWCYCLE);

                        if (performanceReviewCycleEntityReference == null)
                        {
                            Plugin.TraceLog("Skipping resource " + resourceEntity.Id + " — template has no Performance Review Cycle.", iTracingService);
                            continue;
                        }

                        // ── Step 6: In-memory duplicate check ─────────────────────
                        // Catches combinations already processed in this same execution
                        // (these won't be visible in DB yet so the DB query below would miss them).
                        string combinationKey = resourceEntity.Id.ToString()
                            + "_" + reportingManagerEntityReference.Id.ToString()
                            + "_" + performanceReviewCycleEntityReference.Id.ToString();

                        if (processedCombinations.Contains(combinationKey))
                        {
                            Plugin.TraceLog("Skipping resource " + resourceEntity.Id + " — combination already processed in this run.", iTracingService);
                            continue;
                        }

                        // ── Step 7: Database duplicate check ──────────────────────
                        // Catches combinations that already exist from a previous run.
                        QueryExpression duplicateCheckQueryExpression = new QueryExpression(CommonEntities.PERFORMANCEEVALUATION);
                        duplicateCheckQueryExpression.ColumnSet = new ColumnSet(CommonEntities.PerformanceEvaluation.INK_NAME);
                        duplicateCheckQueryExpression.Criteria.AddCondition(
                            CommonEntities.PerformanceEvaluation.INK_ASSESSEE,
                            ConditionOperator.Equal,
                            resourceEntity.Id);
                        duplicateCheckQueryExpression.Criteria.AddCondition(
                            CommonEntities.PerformanceEvaluation.INK_ASSESSOR,
                            ConditionOperator.Equal,
                            reportingManagerEntityReference.Id);
                        duplicateCheckQueryExpression.Criteria.AddCondition(
                            CommonEntities.PerformanceEvaluation.INK_PERFORMANCEREVIEWCYCLE,
                            ConditionOperator.Equal,
                            performanceReviewCycleEntityReference.Id);
                        duplicateCheckQueryExpression.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        duplicateCheckQueryExpression.TopCount = 1;

                        EntityCollection lstDuplicateRecords = iOrganizationService.RetrieveMultiple(duplicateCheckQueryExpression);

                        if (lstDuplicateRecords != null && lstDuplicateRecords.Entities.Count > 0)
                        {
                            Plugin.TraceLog("Skipping resource " + resourceEntity.Id + " — Performance Evaluation already exists in DB.", iTracingService);
                            continue;
                        }

                        // ── Step 8: Register combination in-memory BEFORE creating ─
                        // Must happen before iOrganizationService.Create so that if the
                        // same combination appears again later in this loop it is caught.
                        processedCombinations.Add(combinationKey);

                        // ── Step 9: Build PE name ─────────────────────────────────
                        string firstName = Plugin.GetAttributeValue<string>(resourceEntity, CommonEntities.Resource.INK_FIRSTNAME);
                        string lastName = Plugin.GetAttributeValue<string>(resourceEntity, CommonEntities.Resource.INK_LASTNAME);
                        string cycleName = performanceReviewCycleEntityReference.Name ?? string.Empty;
                        string performanceEvaluationName = firstName + " " + lastName + " " + cycleName;

                        // ── Step 10: Create Performance Evaluation record ──────────
                        EntityReference assesseeEntityReference = new EntityReference(resourceEntity.LogicalName, resourceEntity.Id);

                        Entity performanceEvaluationEntity = new Entity(CommonEntities.PERFORMANCEEVALUATION);
                        Plugin.AddAttribute<string>(performanceEvaluationEntity, CommonEntities.PerformanceEvaluation.INK_NAME, performanceEvaluationName);
                        Plugin.AddAttribute<EntityReference>(performanceEvaluationEntity, CommonEntities.PerformanceEvaluation.INK_ASSESSEE, assesseeEntityReference);
                        Plugin.AddAttribute<EntityReference>(performanceEvaluationEntity, CommonEntities.PerformanceEvaluation.INK_ASSESSOR, reportingManagerEntityReference);
                        Plugin.AddAttribute<EntityReference>(performanceEvaluationEntity, CommonEntities.PerformanceEvaluation.INK_PERFORMANCEREVIEWCYCLE, performanceReviewCycleEntityReference);

                        iOrganizationService.Create(performanceEvaluationEntity);

                        Plugin.TraceLog("Performance Evaluation created: " + performanceEvaluationName, iTracingService);
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