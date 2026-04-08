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
        public const string CHECKLISTITEM = "ink_checklistitem";
        public const string TASK = "task";

        public const string CLIENT = "ink_client";
        public const string CLIENTRETRO = "ink_clientretro";
        public const string PROJECTRETRO = "ink_projectretro";

        public const string OBJECTIVE = "ink_objectives";
        public const string KEYRESULT = "ink_keyresults";

        public const string KEYRESULTRATING = "ink_keyresultratings";

        public const string COREEXPECTATION = "ink_coreexpectations";
        public const string COREEXPECTATIONRATING = "ink_coreexpectationratings";
        public const string STATUSCODE = "statuscode";
        public const string STATECODE = "statecode";
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

            public struct StatusField
            {
                public const int Q1Submitted = 826460000;
                public const int Q1Acknowledged= 826460001;
                public const int Q2Submitted = 826460002;
                public const int Q2Acknowledged = 826460003;    
                public const int Q3Submitted = 826460004;
                public const int Q3Acknowledged = 826460005;    
                public const int Q4Submitted = 826460006;
                public const int Q4Acknowledged = 826460007;    
            }
        }
    }

    
    public struct Status 
    {
            public const int NEW = 826460000;
            public const int STARTED = 826460001;
            public const int SUBMITTED = 826460002;
            public const int CLOSED = 826460003;

    }

    public struct SystemJobStatus
    {
        public const int READY = 0;
        public const int SUSPENDED = 1;
        public const int LOCKED = 2;
        public const int COMPLETED = 3;

    }♥

    public struct SystemJobStatusReason
    {
        public const int SUCCEEDED = 30;
        public const int FAILED = 31;
        public const int CANCELED = 32;

    }
}
