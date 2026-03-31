using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inkey.MSCRM.Plugin_V9._0.Common;


namespace Performance_Assessment_System.Common
{
    public struct CommonEntities
    {
        public const string RESOURCE = "ink_resource";
        public const string PERFORMANCEEVALUATIONTEMPLATE = "ink_performanceevaluationtemplate";
        public const string PERFORMANCEEVALUATION = "ink_performanceevaluations";
        public const string PROJECT = "ink_project";
        public const string AUDITMATRIX = "ink_auditmatrix";
        public const string AUDITFREQUENCYMATRIX = "ink_auditfrequencymatrix";
        public const string AUDIT = "ink_audit";

        public const string OBJECTIVE = "ink_objectives";


        public struct Resource
        {
            public const string INK_RESOURCEID = "ink_resourceid";
            public const string INK_FIRSTNAME = "ink_firstname";
            public const string INK_LASTNAME = "ink_lastname";
            public const string INK_REPORTINGMANAGER = "ink_reportingmanager";
            public const string INK_DESIGNATION = "ink_designation";
        }

        public struct PerformanceEvaluationTemplate
        {
            public const string INK_DESIGNATION = "ink_designation";
            public const string INK_PERFORMANCEREVIEWCYCLE = "ink_performancereviewcycle";
        }

        public struct PerformanceEvaluation
        {
            public const string INK_PERFORMANCEEVALUATIONSID = "ink_performanceevaluationsid";
            
            public const string INK_NAME = "ink_name";
            public const string INK_ASSESSEE = "ink_assessee";
            public const string INK_ASSESSOR = "ink_assessor";
            public const string INK_PERFORMANCEREVIEWCYCLE = "ink_performancereviewcycle";
        }
    }
}
