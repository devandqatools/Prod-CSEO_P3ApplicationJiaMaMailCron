using Microsoft.TeamFoundation.Build.WebApi;
using Microsoft.TeamFoundation.Core.WebApi;
using Microsoft.TeamFoundation.SourceControl.WebApi;
using Microsoft.TeamFoundation.TestManagement.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Data;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Configuration;
using System.Text.RegularExpressions;
using System.IO;

namespace TFRestApiApp
{
    class Program
    {
        static readonly string TFUrl = "https://microsoftit.visualstudio.com/";
        //https://dev.azure.com/<your_org>/ for devops azure 
        static readonly string UserAccount = "";
        static readonly string UserPassword = "";
        static readonly string UserPAT = "";
        //https://docs.microsoft.com/en-us/azure/devops/organizations/accounts/use-personal-access-tokens-to-authenticate?view=azure-devops

        static WorkItemTrackingHttpClient WitClient;
        static BuildHttpClient BuildClient;
        static ProjectHttpClient ProjectClient;
        static GitHttpClient GitClient;
        static TfvcHttpClient TfvsClient;
        static TestManagementHttpClient TestManagementClient;

        public static DataTable allP3UserStories = new DataTable();
        public static DataTable dt_AllRecordsInAirt = new DataTable();
        static readonly string TN_RecordsInAIRT = "AllRecords";

        public static DataTable dt_SelfAssessmentToStart = new DataTable(); //Condition-1: Self-Assessment Not Started
        public static DataTable dt_SelfAssessmentInProgress = new DataTable(); //Condition-2: Self-Assessment In Progress
        public static DataTable dt_SelfAssessmentCompleted = new DataTable(); //Condition-3: Self-Assessment Completed
        public static DataTable dt_BugRemediationPostSelfAssessmentNotStarted = new DataTable(); //Condition-4: Bug Remediation Post Self-Assessment Not Started
        public static DataTable dt_BugRemediationPostSelfAssessmentStarted = new DataTable(); //Condition-5: Bug Remediation Post Self-Assessment In Progress
        public static DataTable dt_BugRemediationPostSelfAssessmentCompleted = new DataTable(); //Condition-6: Bug Remediation Post Self-Assessment Completed
        public static DataTable dt_ReadyForAssessmentServiceOnboarding = new DataTable(); //Condition-8: CSEO Assessment Onboarding Not Started
        public static DataTable dt_AssessmentServiceOnboardingInProgress = new DataTable(); //Condition-9: CSEO Assessment Onboarding In Progress 
        public static DataTable dt_ScheduledForGradeReviewAssessment = new DataTable(); //Condition-10: CSEO Assessment Onboarding Completed
        public static DataTable dt_GradeReviewAssessmentInProgress = new DataTable(); //Condition-11: CSEO Assessment Grade Review In Progress
        public static DataTable dt_GradeReviewAssessmentCompleted = new DataTable(); //Condition-12: CSEO Assessment Grade Review Completed
        public static DataTable dt_FY20GradeC = new DataTable(); //Condition-15: CSEO Assessment Grade C Received
        public static DataTable dt_FY20UserStoryCount = new DataTable(); //Last Count: P3 Applications User Story Count
        public static DataTable dt_BugRemediationPostCSEOAssessmentNotStarted = new DataTable(); //Condition-13: Bug Remediation Post CSEO Assessment Not Started 
        public static DataTable dt_BugRemediationPostCSEOAssessmentInProgress = new DataTable(); //Condition-14: Bug Remediation Post CSEO Assessment In Progress
        public static DataTable dt_awaitingOnboardingDocumentPostSelfAssessment = new DataTable(); //Condition-7: Awaiting Onboarding Document Post Self-Assessment

        static void Main(string[] args)
        {
            dt_AllRecordsInAirt = connectAndGetDataFromAIRTDB(TN_RecordsInAIRT, ConfigurationManager.AppSettings["allRecords"]);

            /* WorkItemType, WorkItemID, Title, Tags, RecordID, NameDesc,group, Subgroup,Priority,
             * Grade,OpsStatus,ServiceOffering,isDeletedRecordInAIRT */
            allP3UserStories.Columns.Add("WorkItemType", typeof(string));
            allP3UserStories.Columns.Add("WorkItemID", typeof(string));
            allP3UserStories.Columns.Add("Title", typeof(string));
            allP3UserStories.Columns.Add("Tags", typeof(string));
            allP3UserStories.Columns.Add("RecordID", typeof(string));
            allP3UserStories.Columns.Add("NameDesc", typeof(string));
            allP3UserStories.Columns.Add("group", typeof(string));
            allP3UserStories.Columns.Add("Subgroup", typeof(string));
            allP3UserStories.Columns.Add("Priority", typeof(string));
            allP3UserStories.Columns.Add("Grade", typeof(string));
            allP3UserStories.Columns.Add("OpsStatus", typeof(string));
            allP3UserStories.Columns.Add("ServiceOffering", typeof(string));
            allP3UserStories.Columns.Add("isDeletedRecordInAIRT", typeof(string));
            allP3UserStories.Columns.Add("UserStoryState", typeof(string));


            //Below are colums for results data 
            //_____________________________Stat_____________________________________
            dt_SelfAssessmentToStart.Columns.Add("WorkItemID", typeof(string));
            dt_SelfAssessmentToStart.Columns.Add("SubGroup", typeof(string));
            dt_SelfAssessmentToStart.Columns.Add("ServiceLine", typeof(string));

            dt_SelfAssessmentInProgress.Columns.Add("WorkItemID", typeof(string));
            dt_SelfAssessmentInProgress.Columns.Add("SubGroup", typeof(string));
            dt_SelfAssessmentInProgress.Columns.Add("ServiceLine", typeof(string));

            dt_SelfAssessmentCompleted.Columns.Add("WorkItemID", typeof(string));
            dt_SelfAssessmentCompleted.Columns.Add("SubGroup", typeof(string));
            dt_SelfAssessmentCompleted.Columns.Add("ServiceLine", typeof(string));

            dt_BugRemediationPostSelfAssessmentNotStarted.Columns.Add("WorkItemID", typeof(string));
            dt_BugRemediationPostSelfAssessmentNotStarted.Columns.Add("SubGroup", typeof(string));
            dt_BugRemediationPostSelfAssessmentNotStarted.Columns.Add("ServiceLine", typeof(string));

            dt_BugRemediationPostSelfAssessmentStarted.Columns.Add("WorkItemID", typeof(string));
            dt_BugRemediationPostSelfAssessmentStarted.Columns.Add("SubGroup", typeof(string));
            dt_BugRemediationPostSelfAssessmentStarted.Columns.Add("ServiceLine", typeof(string));

            dt_BugRemediationPostSelfAssessmentCompleted.Columns.Add("WorkItemID", typeof(string));
            dt_BugRemediationPostSelfAssessmentCompleted.Columns.Add("SubGroup", typeof(string));
            dt_BugRemediationPostSelfAssessmentCompleted.Columns.Add("ServiceLine", typeof(string));

            dt_ReadyForAssessmentServiceOnboarding.Columns.Add("WorkItemID", typeof(string));
            dt_ReadyForAssessmentServiceOnboarding.Columns.Add("SubGroup", typeof(string));
            dt_ReadyForAssessmentServiceOnboarding.Columns.Add("ServiceLine", typeof(string));

            dt_AssessmentServiceOnboardingInProgress.Columns.Add("WorkItemID", typeof(string));
            dt_AssessmentServiceOnboardingInProgress.Columns.Add("SubGroup", typeof(string));
            dt_AssessmentServiceOnboardingInProgress.Columns.Add("ServiceLine", typeof(string));

            dt_ScheduledForGradeReviewAssessment.Columns.Add("WorkItemID", typeof(string));
            dt_ScheduledForGradeReviewAssessment.Columns.Add("SubGroup", typeof(string));
            dt_ScheduledForGradeReviewAssessment.Columns.Add("ServiceLine", typeof(string));

            dt_GradeReviewAssessmentInProgress.Columns.Add("WorkItemID", typeof(string));
            dt_GradeReviewAssessmentInProgress.Columns.Add("SubGroup", typeof(string));
            dt_GradeReviewAssessmentInProgress.Columns.Add("ServiceLine", typeof(string));

            dt_GradeReviewAssessmentCompleted.Columns.Add("WorkItemID", typeof(string));
            dt_GradeReviewAssessmentCompleted.Columns.Add("SubGroup", typeof(string));
            dt_GradeReviewAssessmentCompleted.Columns.Add("ServiceLine", typeof(string));

            //dt_BugRemediationPostCSEOAssessmentNotStarted
            dt_BugRemediationPostCSEOAssessmentNotStarted.Columns.Add("WorkItemID", typeof(string));
            dt_BugRemediationPostCSEOAssessmentNotStarted.Columns.Add("SubGroup", typeof(string));
            dt_BugRemediationPostCSEOAssessmentNotStarted.Columns.Add("ServiceLine", typeof(string));

            //dt_BugRemediationPostCSEOAssessmentInProgress
            dt_BugRemediationPostCSEOAssessmentInProgress.Columns.Add("WorkItemID", typeof(string));
            dt_BugRemediationPostCSEOAssessmentInProgress.Columns.Add("SubGroup", typeof(string));
            dt_BugRemediationPostCSEOAssessmentInProgress.Columns.Add("ServiceLine", typeof(string));

            dt_FY20GradeC.Columns.Add("WorkItemID", typeof(string));
            dt_FY20GradeC.Columns.Add("SubGroup", typeof(string));
            dt_FY20GradeC.Columns.Add("ServiceLine", typeof(string));
            dt_FY20GradeC.Columns.Add("Grade", typeof(string));

            dt_FY20UserStoryCount.Columns.Add("WorkItemID", typeof(string));
            dt_FY20UserStoryCount.Columns.Add("SubGroup", typeof(string));
            dt_FY20UserStoryCount.Columns.Add("ServiceLine", typeof(string));

            //dt_awaitingOnboardingDocumentPostSelfAssessment
            dt_awaitingOnboardingDocumentPostSelfAssessment.Columns.Add("WorkItemID", typeof(string));
            dt_awaitingOnboardingDocumentPostSelfAssessment.Columns.Add("SubGroup", typeof(string));
            dt_awaitingOnboardingDocumentPostSelfAssessment.Columns.Add("ServiceLine", typeof(string));

            /*_______________________________End__________________________________*/


            ConnectWithPAT(TFUrl, UserPAT);

            //Get user stories for P3 applications

            string queryWiqlList = @"select [System.Id], [System.WorkItemType], [System.Title], [System.AssignedTo], [System.State], [System.Tags] from WorkItemLinks where (Source.[System.TeamProject] = @project and Source.[System.WorkItemType] = 'Scenario' and Source.[System.Id] = 5324882) and ([System.Links.LinkType] = 'System.LinkTypes.Hierarchy-Forward') and (Target.[System.TeamProject] = @project and Target.[System.WorkItemType] in ('Feature', 'User Story') and Target.[System.State] <> 'Removed') order by [System.Title] mode (Recursive)";
            //string queryWiqlList = @"SELECT
            //    [System.Id],
            //    [System.WorkItemType],
            //    [System.Title],
            //    [System.AssignedTo],
            //    [System.State],
            //    [System.Tags]
            //FROM workitemLinks
            //WHERE
            //    (
            //        [Source].[System.TeamProject] = @project
            //        AND [Source].[System.WorkItemType] = 'Scenario'
            //        AND [Source].[System.Id] = 5324882
            //    )
            //    AND (
            //        [System.Links.LinkType] = 'System.LinkTypes.Hierarchy-Forward'
            //    )
            //    AND (
            //        [Target].[System.TeamProject] = @project
            //        AND [Target].[System.Id] IN (5549042)
            //    )
            //MODE (Recursive)";

            string teamProject = "OneITVSO";

            Console.WriteLine("_________________________________");
            Console.WriteLine("             Get Flat List");
            Console.WriteLine("_________________________________");

            GetQueryResult(queryWiqlList, teamProject);

            /*If an application is P3, has Fy 20 Tag, based on Subgroup and Service Offering and 
             * if the applications Current Grade is C*/
            var condition12GradeC = (from r in allP3UserStories.AsEnumerable()
                                     where r.Field<string>("Grade") == "C" &&
                                     r.Field<string>("isDeletedRecordInAIRT") =="False"
                                     select new
                                     {
                                         cdT10_workitemID = r.Field<string>("WorkItemID"),
                                         cdT10_SubGroup = r.Field<string>("Subgroup"),
                                         cdT10_ServiceOffering = r.Field<string>("ServiceOffering"),
                                         cdT10_Grade = r.Field<string>("Grade")
                                     });
            foreach (var condition10Values in condition12GradeC)
            {
                dt_FY20GradeC.Rows.Add(condition10Values.cdT10_workitemID, condition10Values.cdT10_SubGroup, condition10Values.cdT10_ServiceOffering,condition10Values.cdT10_Grade);
            }

            //Get the User Story ID's
            var userIDs = allP3UserStories.AsEnumerable().Select(r => r.Field<string>("WorkItemID")).ToList();

            //Verify the User Stories and get those linked Tasks
            foreach (string userid in userIDs)
            {
                int countTask = 0;
                int wiID = int.Parse(userid); //set the work item ID
                var wi = GetWorkItemWithRelations(wiID);
                //var fieldValue = CheckFieldAndGetFieldValue(wi, "System.Title"); // or just: var fieldValue = GetFieldValue(wi, "System.Title");

                //Console.WriteLine("____________________________________________________");
                //Console.WriteLine("                  USER STORY FIELDS");
                //Console.WriteLine("____________________________________________________");

                //foreach (var fieldName in wi.Fields.Keys)
                //    Console.WriteLine("{0,-40}: {1}", fieldName, wi.Fields[fieldName].ToString());

                if (wi.Relations != null)
                {
                    Console.WriteLine("____________________________________________________");
                    Console.WriteLine("                   USER STORY LINKS");
                    Console.WriteLine("____________________________________________________");
                    string Condition3 = "FALSE";
                    string Condition11 = "FALSE";
                    foreach (var wiLink in wi.Relations)
                    {

                        string childTaskURL = wiLink.Url;

                        //Remove the Task Before string
                        string childTaskID = childTaskURL.Replace("https://microsoftit.visualstudio.com/_apis/wit/workItems/", "");
                        var regexItem = new Regex("^[0-9 ]*$");
                        if (regexItem.IsMatch(childTaskID))
                        {
                            //TaskID
                            int wid_childTaskID = int.Parse(childTaskID);
                            Console.WriteLine("Verifyting User Story Liks: {0}", userid);

                            //Add the Task Childs 
                            Dictionary<string, string> parentList = new Dictionary<string, string>();
                            foreach (var newparent in wiLink.Attributes)
                            {
                                parentList.Add(newparent.Key, newparent.Value.ToString());
                            }

                            if (parentList.ContainsValue("Child"))
                            {
                                Console.WriteLine("{0} This is Child..!", wid_childTaskID);

                                //Get the Task Fields
                                var wi_taskWorkFields = GetWorkItem(wid_childTaskID);

                                if (wi_taskWorkFields.Fields["System.WorkItemType"].ToString() == "Task")
                                {
                                    Console.WriteLine("{0} This is Task...!", countTask);
                                    countTask = countTask + 1;

                                    //Get the Task Fields
                                    //var wi_taskWorkFields = GetWorkItem(wid_childTaskID);
                                    //if (wi_taskWorkFields.Fields["System.WorkItemType"].ToString() == "Task")
                                    //{
                                    var taskTitleVar = CheckFieldAndGetFieldValue(wi_taskWorkFields, "System.Title");
                                    string taskTitle = taskTitleVar.ToString().ToLower();

                                    var stateTask = CheckFieldAndGetFieldValue(wi_taskWorkFields, "System.State");
                                    string taskFieldValue_TaskState = stateTask.ToString();

                                    //Condition-1:[Eng][Activity: Self-assessment and Bugs Logging --> New
                                    if (taskTitle.Contains("self-assessment and bugs logging") && taskFieldValue_TaskState.ToString() == "New")
                                    {
                                        string cndt1_userStoryID = wiID.ToString();
                                        string cndt1_ServiceOffering = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt1_userStoryID)["ServiceOffering"].ToString());
                                        string cndt1_SubGroup = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt1_userStoryID)["Subgroup"].ToString());

                                        dt_SelfAssessmentToStart.Rows.Add(cndt1_userStoryID, cndt1_SubGroup, cndt1_ServiceOffering);
                                        Console.WriteLine("_______________Condition-1 DATA Updated[SelfAssessmentToStart]_________________________");

                                        //Console.WriteLine("{0}:{1} Identified Condition-1:SelfAssessmentToStart..!", taskFieldValue_TaskWorkItemType.ToString(), taskID);
                                    }

                                    //Condition-2:[Eng][Activity: Self-assessment and Bugs Logging --> Active
                                    if (taskTitle.Contains("self-assessment and bugs logging") && taskFieldValue_TaskState.ToString() == "Active")
                                    {
                                        string cndt2_userStoryID = wiID.ToString();
                                        string cndt2_ServiceOffering = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt2_userStoryID)["ServiceOffering"].ToString());
                                        string cndt2_SubGroup = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt2_userStoryID)["Subgroup"].ToString());

                                        dt_SelfAssessmentInProgress.Rows.Add(cndt2_userStoryID, cndt2_SubGroup, cndt2_ServiceOffering);

                                        Console.WriteLine("Condition_2 SelfAssessmentInProgress identified.....!");
                                    }

                                    //Condition-3:[Eng][Activity: Self-assessment and Bugs Logging -->Closed
                                    if (taskTitle.Contains("self-assessment and bugs logging") && taskFieldValue_TaskState.ToString() == "Closed")
                                    {
                                        string cndt3_userStoryID = wiID.ToString();
                                        string cndt3_ServiceOffering = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt3_userStoryID)["ServiceOffering"].ToString());
                                        string cndt3_SubGroup = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt3_userStoryID)["Subgroup"].ToString());

                                        dt_SelfAssessmentCompleted.Rows.Add(cndt3_userStoryID, cndt3_SubGroup, cndt3_ServiceOffering);

                                        Console.WriteLine("Condition_3 SelfAssessmentCompleted identified.....!");
                                        Condition3 = "TRUE";
                                    }

                                    //[Eng][Activity: Self-assessment and Bugs Logging -->Closed && Bug Remediation Post Self Assessment Not Started --> New
                                    //if (taskTitle.Contains("bugs remediation post self-assessment") && taskFieldValue_TaskState.ToString() == "New")
                                    //{
                                    //    string cndt4_userStoryID = wiID.ToString();
                                    //    string cndt4_ServiceOffering = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt4_userStoryID)["ServiceOffering"].ToString());
                                    //    string cndt4_SubGroup = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt4_userStoryID)["Subgroup"].ToString());

                                    //    dt_BugRemediationPostSelfAssessmentNotStarted.Rows.Add(cndt4_userStoryID, cndt4_SubGroup, cndt4_ServiceOffering);

                                    //    Console.WriteLine("Condition_4 Bug Remediation Post Self Assessment Not Started identified.....!");
                                    //}

                                    //Condition-4: ([Eng][Activity: Bugs Remediation post Self-assessment] --> New && ([Eng][Activity: Self-assessment and Bugs Logging) --> Close
                                    if (taskTitle.Contains("bugs remediation post self-assessment") && taskFieldValue_TaskState.ToString() == "New")
                                    {
                                        int parentIdForThisTask = wiID;
                                        var cndt4_userStoryID = GetWorkItemWithRelations(parentIdForThisTask);
                                        if (cndt4_userStoryID.Relations != null)
                                        {
                                            Console.WriteLine("Verifying subcondition for 4th Parent Conditon...!");
                                            foreach (var wiLink_Condition4 in cndt4_userStoryID.Relations)
                                            {
                                                string childTaskURL_Condition4 = wiLink_Condition4.Url;
                                                //Remove the Task Before string
                                                string childTaskID_Condition4 = childTaskURL_Condition4.Replace("https://microsoftit.visualstudio.com/_apis/wit/workItems/", "");
                                                var regexItem_Condition4 = new Regex("^[0-9 ]*$");
                                                if (regexItem_Condition4.IsMatch(childTaskID_Condition4))
                                                {
                                                    //TaskID
                                                    int wid_childTaskIDCondition4 = int.Parse(childTaskID_Condition4);
                                                    //Add the Task Childs 
                                                    Dictionary<string, string> parentList_Condition4 = new Dictionary<string, string>();
                                                    foreach (var newparent_Condition4 in wiLink_Condition4.Attributes)
                                                    {
                                                        parentList_Condition4.Add(newparent_Condition4.Key, newparent_Condition4.Value.ToString());
                                                    }
                                                    if (parentList_Condition4.ContainsValue("Child"))
                                                    {
                                                        //Get the Task Fields
                                                        var wi_taskWorkFields_Condition4 = GetWorkItem(wid_childTaskIDCondition4);
                                                        if (wi_taskWorkFields_Condition4.Fields["System.WorkItemType"].ToString() == "Task" && wi_taskWorkFields_Condition4.Fields["System.State"].ToString() == "Closed")
                                                        {
                                                            var taskTitleVar_Condition4 = CheckFieldAndGetFieldValue(wi_taskWorkFields_Condition4, "System.Title");
                                                            string taskTitle_Condition4 = taskTitleVar_Condition4.ToString().ToLower();
                                                            if (taskTitle_Condition4.Contains("self-assessment and bugs logging"))
                                                            {
                                                                string cndt4New_userStoryID = wiID.ToString();
                                                                string cndt4New_ServiceOffering = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt4New_userStoryID)["ServiceOffering"].ToString());
                                                                string cndt4New_SubGroup = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt4New_userStoryID)["Subgroup"].ToString());
                                                                dt_BugRemediationPostSelfAssessmentNotStarted.Rows.Add(cndt4New_userStoryID, cndt4New_SubGroup, cndt4New_ServiceOffering);
                                                            }
                                                            else
                                                            {
                                                                Console.WriteLine("Sub scenario not satistied In Conditon 4..!");
                                                            }
                                                        }
                                                        else
                                                        {
                                                            Console.WriteLine("WorkItemType not in New state for  Conditon 4 in sub scenario");
                                                        }

                                                    }
                                                    else
                                                    {
                                                        Console.WriteLine("This not Child task workItem in subcondition-4..!");
                                                    }

                                                }
                                                else
                                                {
                                                    Console.WriteLine("Currently Identified Not child workItem verifying next one...!");
                                                }
                                            }
                                        }
                                    }

                                    //Condition-5: Bug Remediation Post Self Assessment Started --> Active
                                    if (taskTitle.Contains("bugs remediation post self-assessment") && taskFieldValue_TaskState.ToString() == "Active")
                                    {
                                        string cndt5_userStoryID = wiID.ToString();
                                        string cndt5_ServiceOffering = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt5_userStoryID)["ServiceOffering"].ToString());
                                        string cndt5_SubGroup = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt5_userStoryID)["Subgroup"].ToString());

                                        dt_BugRemediationPostSelfAssessmentStarted.Rows.Add(cndt5_userStoryID, cndt5_SubGroup, cndt5_ServiceOffering);

                                        Console.WriteLine("Condition 5 identified...!");
                                    }

                                    //Condition-6: [Eng][Activity: Bugs Remediation post Self-assessment --> Closed
                                    if (taskTitle.Contains("bugs remediation post self-assessment") && taskFieldValue_TaskState.ToString() == "Closed")
                                    {
                                        string cndt6_userStoryID = wiID.ToString();
                                        string cndt6_ServiceOffering = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt6_userStoryID)["ServiceOffering"].ToString());
                                        string cndt6_SubGroup = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt6_userStoryID)["Subgroup"].ToString());

                                        dt_BugRemediationPostSelfAssessmentCompleted.Rows.Add(cndt6_userStoryID, cndt6_SubGroup, cndt6_ServiceOffering);

                                        Console.WriteLine("Condition 6 identified...!");
                                    }

                                    //Condition-7: Activity: Bugs Remediation post Self-assessment --> Closed (&&) Activity: Create Grade Review Onboarding Request is New or Active
                                    if (taskTitle.Contains("activity: bugs remediation post self-assessment") && taskFieldValue_TaskState.ToString() == "Closed")
                                    {
                                        int parentIdForThisTask = wiID;
                                        var wi_Condition7_New = GetWorkItemWithRelations(parentIdForThisTask);
                                        if (wi_Condition7_New.Relations != null)
                                        {
                                            Console.WriteLine("Conditon:7 activity: bugs remediation post self-assessment--> Close");
                                            foreach (var wiLink_Condition7_New in wi_Condition7_New.Relations)
                                            {
                                                string childTaskURL_Condition7_New = wiLink_Condition7_New.Url;
                                                //Remove the Task Before string
                                                string childTaskID_Condition7_New = childTaskURL_Condition7_New.Replace("https://microsoftit.visualstudio.com/_apis/wit/workItems/", "");
                                                var regexItem_Condition7_New = new Regex("^[0-9 ]*$");
                                                if (regexItem_Condition7_New.IsMatch(childTaskID_Condition7_New))
                                                {
                                                    //TaskID
                                                    int wid_childTaskIDCondition7_New = int.Parse(childTaskID_Condition7_New);
                                                    //Add the Task Childs 
                                                    Dictionary<string, string> parentList_Condition7_New = new Dictionary<string, string>();
                                                    foreach (var newparent_Condition7_New in wiLink_Condition7_New.Attributes)
                                                    {
                                                        parentList_Condition7_New.Add(newparent_Condition7_New.Key, newparent_Condition7_New.Value.ToString());
                                                    }
                                                    if (parentList_Condition7_New.ContainsValue("Child"))
                                                    {
                                                        //Get the Task Fields
                                                        var wi_taskWorkFields_Condition7_New = GetWorkItem(wid_childTaskIDCondition7_New);
                                                        if (wi_taskWorkFields_Condition7_New.Fields["System.WorkItemType"].ToString() == "Task" && (wi_taskWorkFields_Condition7_New.Fields["System.State"].ToString() == "New" || wi_taskWorkFields_Condition7_New.Fields["System.State"].ToString() == "Active"))
                                                        {
                                                            var taskTitleVar_Condition7_New = CheckFieldAndGetFieldValue(wi_taskWorkFields_Condition7_New, "System.Title");
                                                            string taskTitle_Condition7_New = taskTitleVar_Condition7_New.ToString().ToLower();
                                                            //For Conditon 7 in Sub scenario title: [Activity:Onboarding]---> New
                                                            if (taskTitle_Condition7_New.Contains("create grade review onboarding request"))
                                                            {
                                                                string cndt7_userStoryID_New = wiID.ToString();
                                                                string cndt7_ServiceOffering_New = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt7_userStoryID_New)["ServiceOffering"].ToString());
                                                                string cndt7_SubGroup_New = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt7_userStoryID_New)["Subgroup"].ToString());
                                                                dt_awaitingOnboardingDocumentPostSelfAssessment.Rows.Add(cndt7_userStoryID_New, cndt7_SubGroup_New, cndt7_ServiceOffering_New);
                                                            }
                                                            else
                                                            {
                                                                Console.WriteLine("Conditon:7: Sub Conditon: Create Grade Review Onboarding Request: New or Active");
                                                            }
                                                        }
                                                        else
                                                        {
                                                            Console.WriteLine("WorkItemType not in New state for  Conditon 7 in sub scenario");
                                                        }

                                                    }
                                                    else
                                                    {
                                                        Console.WriteLine("This not Child task workItem in subcondition-7..!");
                                                    }

                                                }
                                                else
                                                {
                                                    Console.WriteLine("Currently Identified Not child workItem verifying next one...!");
                                                }
                                            }
                                        }
                                    }

                                    //Condition-8: Create Grade Review Onboarding Request --> Closed && Activity:Onboarding --> New
                                    if (taskTitle.Contains("create grade review onboarding request") && taskFieldValue_TaskState.ToString() == "Closed")
                                    {
                                        /*Get the parent id and restof taks here 
                                          Title Shouldbe: [Activity:Onboarding] and State -> New*/
                                        int parentIdForThisTask = wiID;
                                        var wi_Condition7 = GetWorkItemWithRelations(parentIdForThisTask);
                                        if (wi_Condition7.Relations != null)
                                        {
                                            Console.WriteLine("Verifying subcondition If Task (Assessment Service Onboarding is in NewState) in Conditon 7");
                                            foreach (var wiLink_Condition7 in wi_Condition7.Relations)
                                            {
                                                string childTaskURL_Condition7 = wiLink_Condition7.Url;
                                                //Remove the Task Before string
                                                string childTaskID_Condition7 = childTaskURL_Condition7.Replace("https://microsoftit.visualstudio.com/_apis/wit/workItems/", "");
                                                var regexItem_Condition7 = new Regex("^[0-9 ]*$");
                                                if (regexItem_Condition7.IsMatch(childTaskID_Condition7))
                                                {
                                                    //TaskID
                                                    int wid_childTaskIDCondition7 = int.Parse(childTaskID_Condition7);
                                                    //Add the Task Childs 
                                                    Dictionary<string, string> parentList_Condition7 = new Dictionary<string, string>();
                                                    foreach (var newparent_Condition7 in wiLink_Condition7.Attributes)
                                                    {
                                                        parentList_Condition7.Add(newparent_Condition7.Key, newparent_Condition7.Value.ToString());
                                                    }
                                                    if (parentList_Condition7.ContainsValue("Child"))
                                                    {
                                                        //Get the Task Fields
                                                        var wi_taskWorkFields_Condition7 = GetWorkItem(wid_childTaskIDCondition7);
                                                        if (wi_taskWorkFields_Condition7.Fields["System.WorkItemType"].ToString() == "Task" && wi_taskWorkFields_Condition7.Fields["System.State"].ToString() == "New")
                                                        {
                                                            var taskTitleVar_Condition7 = CheckFieldAndGetFieldValue(wi_taskWorkFields_Condition7, "System.Title");
                                                            string taskTitle_Condition7 = taskTitleVar_Condition7.ToString().ToLower();
                                                            //For Conditon 7 in Sub scenario title: [Activity:Onboarding]---> New
                                                            if (taskTitle_Condition7.Contains("activity:onboarding"))
                                                            {
                                                                string cndt7_userStoryID = wiID.ToString();
                                                                string cndt7_ServiceOffering = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt7_userStoryID)["ServiceOffering"].ToString());
                                                                string cndt7_SubGroup = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt7_userStoryID)["Subgroup"].ToString());
                                                                dt_ReadyForAssessmentServiceOnboarding.Rows.Add(cndt7_userStoryID, cndt7_SubGroup, cndt7_ServiceOffering);
                                                            }
                                                            else
                                                            {
                                                                Console.WriteLine("Sub scenario not satistied In Conditon 7, that means Title:[Activity:Onboarding] --> Not in New state..!");
                                                            }
                                                        }
                                                        else
                                                        {
                                                            Console.WriteLine("WorkItemType not in New state for  Conditon 7 in sub scenario");
                                                        }

                                                    }
                                                    else
                                                    {
                                                        Console.WriteLine("This not Child task workItem in subcondition-7..!");
                                                    }

                                                }
                                                else
                                                {
                                                    Console.WriteLine("Currently Identified Not child workItem verifying next one...!");
                                                }
                                            }
                                        }
                                    }

                                    //Condition-9: Activity:Onboarding]) is in Active State 
                                    if (taskTitle.Contains("activity:onboarding") && taskFieldValue_TaskState.ToString() == "Active")
                                    {
                                        string cndt09_userStoryID = wiID.ToString();
                                        string cndt09_ServiceOffering = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt09_userStoryID)["ServiceOffering"].ToString());
                                        string cndt09_SubGroup = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt09_userStoryID)["Subgroup"].ToString());

                                        dt_AssessmentServiceOnboardingInProgress.Rows.Add(cndt09_userStoryID, cndt09_SubGroup, cndt09_ServiceOffering);

                                        Console.WriteLine("Condition 10 identified...!");
                                    }

                                    //Condition-10: Activity:Onboarding is in Closed
                                    if (taskTitle.Contains("activity:onboarding") && taskFieldValue_TaskState.ToString() == "Closed")
                                    {
                                        string cndt10_userStoryID = wiID.ToString();
                                        string cndt10_ServiceOffering = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt10_userStoryID)["ServiceOffering"].ToString());
                                        string cndt10_SubGroup = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt10_userStoryID)["Subgroup"].ToString());

                                        dt_ScheduledForGradeReviewAssessment.Rows.Add(cndt10_userStoryID, cndt10_SubGroup, cndt10_ServiceOffering);

                                        Console.WriteLine("Condition 10 identified...!");
                                    }

                                    //Condition-11: Activity:Assessment is in Active state
                                    if (taskTitle.Contains("activity:assessment") && taskFieldValue_TaskState.ToString() == "Active")
                                    {
                                        string cndt10_userStoryID = wiID.ToString();
                                        string cndt10_ServiceOffering = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt10_userStoryID)["ServiceOffering"].ToString());
                                        string cndt10_SubGroup = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt10_userStoryID)["Subgroup"].ToString());

                                        dt_GradeReviewAssessmentInProgress.Rows.Add(cndt10_userStoryID, cndt10_SubGroup, cndt10_ServiceOffering);

                                        Console.WriteLine("Condition 10 identified...!");
                                    }

                                    //Condition-12: Activity:Assessment is in Closed
                                    if (taskTitle.Contains("activity:assessment") && taskFieldValue_TaskState.ToString() == "Closed")
                                    {
                                        string cndt11_userStoryID = wiID.ToString();
                                        string cndt11_ServiceOffering = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt11_userStoryID)["ServiceOffering"].ToString());
                                        string cndt11_SubGroup = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt11_userStoryID)["Subgroup"].ToString());

                                        dt_GradeReviewAssessmentCompleted.Rows.Add(cndt11_userStoryID, cndt11_SubGroup, cndt11_ServiceOffering);

                                        Console.WriteLine("Condition11 identified...!");
                                        Condition11 = "TRUE";
                                    }

                                    //Condition-13: Activity:Assessment--> Closed && Activity: Bug Remediation and Verification post Grade Review --> New
                                    if (taskTitle.Contains("bug remediation and verification post grade review") && taskFieldValue_TaskState.ToString() == "New")
                                    {
                                        int parentIdForThisTask = wiID;
                                        var wi_Condition13 = GetWorkItemWithRelations(parentIdForThisTask);
                                        if (wi_Condition13.Relations != null)
                                        {
                                            Console.WriteLine("Verifying Conditon 13....!");
                                            foreach (var wiLink_Condition13 in wi_Condition13.Relations)
                                            {
                                                string childTaskURL_Condition13 = wiLink_Condition13.Url;
                                                //Remove the Task Before string
                                                string childTaskID_Condition13 = childTaskURL_Condition13.Replace("https://microsoftit.visualstudio.com/_apis/wit/workItems/", "");
                                                var regexItem_Condition13 = new Regex("^[0-9 ]*$");
                                                if (regexItem_Condition13.IsMatch(childTaskID_Condition13))
                                                {
                                                    //TaskID
                                                    int wid_childTaskIDCondition13 = int.Parse(childTaskID_Condition13);
                                                    //Add the Task Childs 
                                                    Dictionary<string, string> parentList_Condition13 = new Dictionary<string, string>();
                                                    foreach (var newparent_Condition13 in wiLink_Condition13.Attributes)
                                                    {
                                                        parentList_Condition13.Add(newparent_Condition13.Key, newparent_Condition13.Value.ToString());
                                                    }
                                                    if (parentList_Condition13.ContainsValue("Child"))
                                                    {
                                                        //Get the Task Fields
                                                        var wi_taskWorkFields_Condition13 = GetWorkItem(wid_childTaskIDCondition13);
                                                        if (wi_taskWorkFields_Condition13.Fields["System.WorkItemType"].ToString() == "Task" && wi_taskWorkFields_Condition13.Fields["System.State"].ToString() == "Closed")
                                                        {
                                                            var taskTitleVar_Condition13 = CheckFieldAndGetFieldValue(wi_taskWorkFields_Condition13, "System.Title");
                                                            string taskTitle_Condition13 = taskTitleVar_Condition13.ToString().ToLower();
                                                            //For Conditon 7 in Sub scenario title: [Activity:Onboarding]---> New
                                                            if (taskTitle_Condition13.Contains("activity:assessment"))
                                                            {
                                                                string cndt13New_userStoryID = wiID.ToString();
                                                                string cndt13New_ServiceOffering = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt13New_userStoryID)["ServiceOffering"].ToString());
                                                                string cndt13New_SubGroup = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt13New_userStoryID)["Subgroup"].ToString());
                                                                dt_BugRemediationPostCSEOAssessmentNotStarted.Rows.Add(cndt13New_userStoryID, cndt13New_SubGroup, cndt13New_ServiceOffering);
                                                            }
                                                            else
                                                            {
                                                                Console.WriteLine("Sub scenario Condition 13...!");
                                                            }
                                                        }
                                                        else
                                                        {
                                                            Console.WriteLine("WorkItemType not in New state for  Conditon 13 in sub scenario");
                                                        }

                                                    }
                                                    else
                                                    {
                                                        Console.WriteLine("This not Child task workItem in subcondition-13..!");
                                                    }

                                                }
                                                else
                                                {
                                                    Console.WriteLine("Currently Identified Not child workItem verifying next one...!");
                                                }
                                            }
                                        }
                                    }

                                    //Condition-14: Activity: Bug Remediation and Verification post Grade Review --> Active
                                    if (taskTitle.Contains("bug remediation and verification post grade review") && taskFieldValue_TaskState.ToString() == "Active")
                                    {
                                        string cndt15_userStoryID = wiID.ToString();
                                        string cndt15_ServiceOffering = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt15_userStoryID)["ServiceOffering"].ToString());
                                        string cndt15_SubGroup = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt15_userStoryID)["Subgroup"].ToString());

                                        dt_BugRemediationPostCSEOAssessmentInProgress.Rows.Add(cndt15_userStoryID, cndt15_SubGroup, cndt15_ServiceOffering);

                                        Console.WriteLine("Condition14 identified...!");
                                    }

                                }
                            }
                        } 
                    }
                }
            }

            List<DataTable> dataTableList = new List<DataTable>();
            dataTableList.Add(dt_SelfAssessmentToStart);
            dataTableList.Add(dt_SelfAssessmentInProgress);
            dataTableList.Add(dt_SelfAssessmentCompleted);
            dataTableList.Add(dt_BugRemediationPostSelfAssessmentNotStarted);
            dataTableList.Add(dt_BugRemediationPostSelfAssessmentStarted);
            dataTableList.Add(dt_BugRemediationPostSelfAssessmentCompleted);
            dataTableList.Add(dt_awaitingOnboardingDocumentPostSelfAssessment);
            dataTableList.Add(dt_ReadyForAssessmentServiceOnboarding);
            dataTableList.Add(dt_AssessmentServiceOnboardingInProgress);
            dataTableList.Add(dt_ScheduledForGradeReviewAssessment);
            dataTableList.Add(dt_GradeReviewAssessmentInProgress);
            dataTableList.Add(dt_GradeReviewAssessmentCompleted);
            dataTableList.Add(dt_BugRemediationPostCSEOAssessmentNotStarted);
            dataTableList.Add(dt_BugRemediationPostCSEOAssessmentInProgress);
            dataTableList.Add(allP3UserStories);
            //dataTableList.Add(dt_FY20UserStoryCount);
            generateXlsDataSheet(dataTableList);

            Console.WriteLine("======================================================================");
            Console.WriteLine("                           EXECUTION COMPLETED                        ");
           Console.WriteLine("======================================================================");
            
            sendMail();

        }


        private static void generateXlsDataSheet(List<DataTable> listDataTables)
        {
            object missing = Type.Missing;
            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook oWB = oXL.Workbooks.Add(missing);
            for (int i = 0; i < listDataTables.Count; i++)
            {
                Microsoft.Office.Interop.Excel.Worksheet Sheet = null;

                if (i == 0)
                {
                    Sheet = oWB.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                }
                else
                {
                    Sheet = oWB.Sheets.Add(missing, missing, 1, missing) as Microsoft.Office.Interop.Excel.Worksheet;
                }

                Sheet.get_Range("A1", "Z1").Font.Bold = true;
                Sheet.Name = ConfigurationManager.AppSettings[i.ToString()];
                DataTable currentDataTable = listDataTables[i];
                int iCol1 = 0;
                foreach (DataColumn c in currentDataTable.Columns)
                {
                    iCol1++;
                    Sheet.Cells[1, iCol1] = c.ColumnName;
                }


                int iRow1 = 0;
                foreach (DataRow r in currentDataTable.Rows)
                {
                    iRow1++;

                    for (int index = 1; index < currentDataTable.Columns.Count + 1; index++)
                    {

                        if (iRow1 == 1)
                        {
                            Sheet.Cells[iRow1, index] = currentDataTable.Columns[index - 1].ColumnName;
                        }

                        Sheet.Cells[iRow1 + 1, index] = r[index - 1].ToString();
                    }
                }
            }
            //string fileName1 = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
            //                        + "\\DataAttachment.xlsx";
            string fileName1 = "D:\\Reports Excel\\AppP3UserStories\\P3UserStories.xlsx";
            oWB.SaveAs(fileName1, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook,
                missing, missing, missing, missing,
                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                missing, missing, missing, missing, missing);
            oWB.Close(missing, missing, missing);
            oXL.UserControl = true;
            oXL.Quit();
        }

        /// <summary>
        /// FUNCTION FOR FORMATTING EXCEL CELLS
        /// </summary>
        /// <param name="range"></param>
        /// <param name="HTMLcolorCode"></param>
        /// <param name="fontColor"></param>
        /// <param name="IsFontbool"></param>
        public static void FormattingExcelCells(Microsoft.Office.Interop.Excel.Range range, string HTMLcolorCode, System.Drawing.Color fontColor, bool IsFontbool)
        {
            range.Interior.Color = System.Drawing.ColorTranslator.FromHtml(HTMLcolorCode);
            range.Font.Color = System.Drawing.ColorTranslator.ToOle(fontColor);
            if (IsFontbool == true)
            {
                range.Font.Bold = IsFontbool;
            }
        }


        /// <summary>
        /// Run query and show result
        /// </summary>
        /// <param name="wiqlStr">Wiql String</param>
        static void GetQueryResult(string wiqlStr, string teamProject)
        {
            WorkItemQueryResult result = RunQueryByWiql(wiqlStr, teamProject);

            string workItemType = null;
            string workItemID = null;
            string UserStorytitle = null;
            string tags = null;
            //string tagsAddString = null;
            string recordID = null;
            string recordIdFinal = null;
            string userStoryState = null;

            if (result != null)
            {
                if (result.WorkItems != null) // this is Flat List 
                    foreach (var wiRef in result.WorkItems)
                    {
                        var wi = GetWorkItem(wiRef.Id);
                        Console.WriteLine(String.Format("{0} - {1}", wi.Id, wi.Fields["System.WorkItemType"].ToString()));
                        if(wi.Fields["System.WorkItemType"].ToString()== "User Story")
                        {
                            workItemType = CheckFieldAndGetFieldValue(wi, "System.WorkItemType").ToString();
                            workItemID = CheckFieldAndGetFieldValue(wi, "System.Id").ToString();
                            UserStorytitle = CheckFieldAndGetFieldValue(wi, "System.Title").ToString().ToLower();
                            tags = CheckFieldAndGetFieldValue(wi, "System.Tags").ToString();

                            if(UserStorytitle.Contains("recid"))
                            {
                                Console.WriteLine("_____________________________________________________________________");
                                Console.WriteLine("                 WORKITEMS RESULTS NOT NULL USER STORY FIELDS");
                                Console.WriteLine("______________________________________________________________________");

                                //Verify the Record Tag pattern Contains RecID_ or RecID- then Remove the _ or - special Symboles in userstory title
                                if (UserStorytitle.Contains("-") || UserStorytitle.Contains("_"))
                                {
                                    List<char> charsToRemove = new List<char>() { '-', '_' };
                                    foreach (char c in charsToRemove)
                                    {
                                        UserStorytitle = UserStorytitle.Replace(c.ToString(), String.Empty);
                                    }
                                }
                                recordIdFinal = betweenStrings(UserStorytitle, "recid", "]");
                                //Remove the emptlye spaces if any
                                recordIdFinal = recordIdFinal.Replace(" ", ""); //Final String
                            }
                            else
                            {
                                Console.WriteLine("{0}: {1} not contains RecID in title....!", workItemType, workItemID);
                            }

                            //WorkItemType, WorkItemID, Title, Tags, RecordID
                            allP3UserStories.Rows.Add(workItemType, workItemID, UserStorytitle, tags, recordIdFinal);
                        }
                    }
                else 
                if (result.WorkItemRelations != null) // this is Tree of Work Items or Work Items and Direct Links
                {
                    foreach (var wiRel in result.WorkItemRelations)
                    {
                        if (wiRel.Source == null)
                        {
                            var wi = GetWorkItem(wiRel.Target.Id);
                            Console.WriteLine(String.Format("Top Level: {0} - {1}", wi.Id, wi.Fields["System.WorkItemType"].ToString()));
                        }
                        else
                        {
                            var wiParent = GetWorkItem(wiRel.Source.Id);
                            var wiChild = GetWorkItem(wiRel.Target.Id);
                            Console.WriteLine(String.Format("{0} --> {1} - {2}", wiParent.Id, wiChild.Id, wiChild.Fields["System.WorkItemType"].ToString()));


                            if (wiChild.Fields["System.WorkItemType"].ToString() == "User Story")
                            {
                                workItemType = CheckFieldAndGetFieldValue(wiChild, "System.WorkItemType").ToString();
                                workItemID = wiChild.Id.ToString();
                                UserStorytitle = CheckFieldAndGetFieldValue(wiChild, "System.Title").ToString().ToLower();
                                tags = CheckFieldAndGetFieldValue(wiChild, "System.Tags").ToString();
                                userStoryState = CheckFieldAndGetFieldValue(wiChild, "System.State").ToString();

                                if (UserStorytitle.Contains("recid"))
                                {
                                    //Verify the Record Tag pattern Contains RecID_ or RecID- then Remove the _ or - special Symboles in userstory title
                                    if (UserStorytitle.Contains("-") || UserStorytitle.Contains("_"))
                                    {
                                        List<char> charsToRemove = new List<char>() { '-', '_' };
                                        foreach (char c in charsToRemove)
                                        {
                                            UserStorytitle = UserStorytitle.Replace(c.ToString(), String.Empty);
                                        }
                                    }
                                    recordIdFinal = betweenStrings(UserStorytitle, "recid", "]");
                                    //Remove the emptlye spaces if any
                                    recordIdFinal = recordIdFinal.Replace(" ", ""); //Final String
                                }
                                else
                                {
                                    Console.WriteLine("{0}: {1} not contains RecID in title....!", workItemType, workItemID);
                                }

                                string NameDesc = (dt_AllRecordsInAirt.AsEnumerable().FirstOrDefault(p => p["RecId"].ToString() == recordIdFinal)["NameDesc"].ToString());
                                string group = (dt_AllRecordsInAirt.AsEnumerable().FirstOrDefault(p => p["RecId"].ToString() == recordIdFinal)["Grp"].ToString());
                                string Subgroup = (dt_AllRecordsInAirt.AsEnumerable().FirstOrDefault(p => p["RecId"].ToString() == recordIdFinal)["SubGrp"].ToString());
                                string Priority = (dt_AllRecordsInAirt.AsEnumerable().FirstOrDefault(p => p["RecId"].ToString() == recordIdFinal)["Priority"].ToString());
                                string Grade = (dt_AllRecordsInAirt.AsEnumerable().FirstOrDefault(p => p["RecId"].ToString() == recordIdFinal)["Grade"].ToString());
                                string OpsStatus = (dt_AllRecordsInAirt.AsEnumerable().FirstOrDefault(p => p["RecId"].ToString() == recordIdFinal)["OpsStatus"].ToString());
                                string ServiceOffering = (dt_AllRecordsInAirt.AsEnumerable().FirstOrDefault(p => p["RecId"].ToString() == recordIdFinal)["SrvOffering"].ToString());
                                string isDeletedRecordInAIRT = (dt_AllRecordsInAirt.AsEnumerable().FirstOrDefault(p => p["RecId"].ToString() == recordIdFinal)["isDeleted"].ToString());

                                if(tags.Contains("FY20") && Priority=="3" && group == "CSEO")
                                {
                                    //WorkItemType, WorkItemID, Title, Tags, RecordID, NameDesc,group, Subgroup,Priority,Grade,OpsStatus,ServiceOffering,isDeletedRecordInAIRT
                                    allP3UserStories.Rows.Add(workItemType, workItemID, UserStorytitle, tags, recordIdFinal, NameDesc, group, Subgroup, Priority, Grade, OpsStatus, ServiceOffering, isDeletedRecordInAIRT, userStoryState);

                                    Console.Write("Added WorkITemType:{0} workItemID:{1}.....!", workItemType, workItemID);
                                }
                                else
                                {
                                    Console.Write("Not Added WorkITemType:{0} workItemID:{1} because of tag not contains FY20.....!", workItemType, workItemID);
                                }
                            }
                        }
                    }
                }
                else Console.WriteLine("There is no query result");
            }
        }


        #region AIRT DB Connection
        private static System.Data.DataTable connectAndGetDataFromAIRTDB(String tableName, string QueryName)
        {
            Console.WriteLine("___________________________________________________________");
            Console.WriteLine("            CONNECTING AIRT ANG GETTING RECORDS");
            Console.WriteLine("___________________________________________________________");

            System.Data.DataTable dt = new System.Data.DataTable();
            SqlCommand cmd = new SqlCommand();
            string dbConn = null;
            dbConn = @"";
            cmd.CommandText = QueryName;
            SqlConnection sqlConnection1 = new SqlConnection(dbConn);
            cmd.Connection = sqlConnection1;
            sqlConnection1.Open();
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            sda.Fill(dt);
            dt.TableName = tableName;
            sqlConnection1.Close();
            Console.WriteLine("In {0} table identified {1} rows and {2} colums.......", dt.TableName, dt.Rows.Count, dt.Columns.Count);
            return dt;
        }
        #endregion


        /// <summary>
        /// Get the string inbetween two string in single string
        /// </summary>
        /// <param name="text"></param>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <returns></returns>
        public static String betweenStrings(String text, String start, String end)
        {
            int p1 = text.IndexOf(start) + start.Length;
            int p2 = text.IndexOf(end, p1);

            if (end == "") return (text.Substring(p1));
            else return text.Substring(p1, p2 - p1);
        }

        /// <summary>
        /// Run Query with Wiql
        /// </summary>
        /// <param name="wiqlStr">Wiql String</param>
        /// <returns></returns>
        static WorkItemQueryResult RunQueryByWiql(string wiqlStr, string teamProject)
        {
            Wiql wiql = new Wiql();
            wiql.Query = wiqlStr;

            if (teamProject == "") return WitClient.QueryByWiqlAsync(wiql).Result;
            else return WitClient.QueryByWiqlAsync(wiql, teamProject).Result;
        }

        /// <summary>
        /// Get one work item
        /// </summary>
        /// <param name="Id"></param>
        /// <returns></returns>
        static WorkItem GetWorkItem(int Id)
        {
            return WitClient.GetWorkItemAsync(Id).Result;
        }

        /// <summary>
        /// Get several work items
        /// </summary>
        /// <param name="Ids"></param>
        /// <returns></returns>
        static List<WorkItem> GetWorkItems(List<int> Ids)
        {
            return WitClient.GetWorkItemsAsync(Ids).Result;
        }

        /// <summary>
        /// Get one work item with information about linked work items
        /// </summary>
        /// <param name="Id"></param>
        /// <returns></returns>
        static WorkItem GetWorkItemWithRelations(int Id)
        {
            return WitClient.GetWorkItemAsync(Id, expand: WorkItemExpand.Relations).Result;
        }

        /// <summary>
        /// Get a string value of a field
        /// </summary>
        /// <param name="WI"></param>
        /// <param name="FieldName"></param>
        /// <returns></returns>
        static string GetFieldValue(WorkItem WI, string FieldName)
        {
            if (!WI.Fields.Keys.Contains(FieldName)) return null;

            return (string)WI.Fields[FieldName];
        }

        /// <summary>
        /// Check field in a work item type then get a string field value
        /// </summary>
        /// <param name="WI"></param>
        /// <param name="FieldName"></param>
        /// <returns></returns>
        static string CheckFieldAndGetFieldValue(WorkItem WI, string FieldName)
        {
            WorkItemType wiType = GetWorkItemType(WI);

            var fields = from field in wiType.Fields where field.Name == FieldName || field.ReferenceName == FieldName select field;

            if (fields.Count() < 1) throw new ArgumentException("Work Item Type " + wiType.Name + " does not contain the field " + FieldName, "CheckFieldAndGetFieldValue");

            return GetFieldValue(WI, FieldName);
        }

        /// <summary>
        /// Get a work item type definition for existing work item
        /// </summary>
        /// <param name="WI"></param>
        /// <returns></returns>
        static WorkItemType GetWorkItemType(WorkItem WI)
        {
            if (!WI.Fields.Keys.Contains("System.WorkItemType")) throw new ArgumentException("There is no WorkItemType field in the workitem", "GetWorkItemType");
            if (!WI.Fields.Keys.Contains("System.TeamProject")) throw new ArgumentException("There is no TeamProject field in the workitem", "GetWorkItemType");

            return WitClient.GetWorkItemTypeAsync((string)WI.Fields["System.TeamProject"], (string)WI.Fields["System.WorkItemType"]).Result;
        }

        /// <summary>
        /// Extract an id from an url
        /// </summary>
        /// <param name="Url"></param>
        /// <returns></returns>
        static int ExtractWiIdFromUrl(string Url)
        {
            int id = -1;

            string splitStr = "_apis/wit/workItems/";

            if (Url.Contains(splitStr))
            {
                string[] strarr = Url.Split(new string[] { splitStr }, StringSplitOptions.RemoveEmptyEntries);

                if (strarr.Length == 2 && int.TryParse(strarr[1], out id))
                    return id;
            }

            return id;
        }

        /// <summary>
        /// Get a work item type definition from a project
        /// </summary>
        /// <param name="TeamProjectName"></param>
        /// <param name="WITypeName"></param>
        /// <returns></returns>
        static WorkItemType GetWorkItemType(string TeamProjectName, string WITypeName)
        {
            return WitClient.GetWorkItemTypeAsync(TeamProjectName, WITypeName).Result;
        }


        #region connection

        static void InitClients(VssConnection Connection)
        {
            WitClient = Connection.GetClient<WorkItemTrackingHttpClient>();
            BuildClient = Connection.GetClient<BuildHttpClient>();
            ProjectClient = Connection.GetClient<ProjectHttpClient>();
            GitClient = Connection.GetClient<GitHttpClient>();
            TfvsClient = Connection.GetClient<TfvcHttpClient>();
            TestManagementClient = Connection.GetClient<TestManagementHttpClient>();
        }

        static void ConnectWithDefaultCreds(string ServiceURL)
        {
            VssConnection connection = new VssConnection(new Uri(ServiceURL), new VssCredentials());
            InitClients(connection);
        }

        static void ConnectWithCustomCreds(string ServiceURL, string User, string Password)
        {
            VssConnection connection = new VssConnection(new Uri(ServiceURL), new WindowsCredential(new NetworkCredential(User, Password)));
            InitClients(connection);
        }

        static void ConnectWithPAT(string ServiceURL, string PAT)
        {
            VssConnection connection = new VssConnection(new Uri(ServiceURL), new VssBasicCredential(string.Empty, PAT));
            InitClients(connection);
        }

        #endregion

        private static void sendMail()
        {
            StringBuilder builder = new StringBuilder();
            Microsoft.Office.Interop.Outlook.Application oApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.MailItem oMsg = (Microsoft.Office.Interop.Outlook.MailItem)oApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);

            builder.Append("<html xmlns:v='urn:schemas-microsoft-com:vml' xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns:m='http://schemas.microsoft.com/office/2004/12/omml' xmlns='http://www.w3.org/TR/REC-html40>");
            builder.Append("<head>");
            builder.Append("<meta http-equiv='Content-Type' content='text/html; charset=Windows-1252'>");
            builder.Append("<meta name='Generator' content='Microsoft Word 15 (filtered medium)'>");
            builder.Append("<style>");
            builder.Append("table,td,th{border:1px solid #000;border-collapse:collapse;font-family:Segoe UI,sans-serif}.rowheader{background-color:#b4c6e7}.rowheader th{height:59px;padding:5px;display:table-cell;vertical-align:sub;font-size:16px;font-weight:400;font-family:Segoe UI,sans-serif;text-align:center}.text-td{text-align:center}.percentage{font-size:11px;font-weight:700}.txtcc{font-size:11px;font-weight:400;padding-left:9px}");
            builder.Append("</style>");
            builder.Append("</head>");
            builder.Append("<body>");

            //builder.Append("<span style=\"font-size:13.0pt;font-family:&quot;Segoe UI&quot;,sans-serif\" id=\"pageH2\"><b>CSEO FY20 P3 Accessibility Progress Scorecard</b></span>");
            builder.Append("<table aria-describedby=\"pageH2\">");
            builder.Append("<tr class=\"rowheader rh1\">");
            builder.Append("<th colspan=\"18\" scope=\"colgroup\">CSEO FY20 P3 Accessibility Progress Scorecard</th>");
            builder.Append("</tr>");
            builder.Append("<tr class=\"rowheader\">");
            builder.Append("<th><b>CSEO Organization</b></th>");
            builder.Append("<th><b>Service Group</b></th>");
            builder.Append("<th><b>Self-Assessment Not Started</b><sup> [<a href=\"#c1\">1</a>]</sup></th>");
            builder.Append("<th><b>Self-Assessment In Progress</b><sup> [<a href=\"#c2\">2</a>]</sup></th>");
            builder.Append("<th bgcolor=\"#FFF2CC\"><b>Self-Assessment Completed</b><sup> [<a href=\"#c3\">3</a>]</sup></th>");
            builder.Append("<th><b>Bug Remediation Post Self-Assessment Not Started</b><sup> [<a href=\"#c4\">4</a>]</sup></th>");
            builder.Append("<th><b>Bug Remediation Post Self-Assessment In Progress</b><sup> [<a href=\"#c5\">5</a>]</sup></th>");
            builder.Append("<th bgcolor=\"#ffe699\"><b>Bug Remediation Post Self-Assessment Completed</b><sup> [<a href=\"#c6\">6</a>]</sup></th>");
            builder.Append("<th><b>Awaiting Onboarding Document Post Self-Assessment</b><sup> [<a href=\"#c7\">7</a>]</sup></th>");
            builder.Append("<th><b>CSEO Assessment Onboarding Not Started</b><sup> [<a href=\"#c8\">8</a>]</sup></th>");
            builder.Append("<th><b>CSEO Assessment Onboarding In Progress</b><sup> [<a href=\"#c9\">9</a>]</sup></th>");
            builder.Append("<th bgcolor=\"#FFD966\"><b>CSEO Assessment Onboarding Completed</b><sup> [<a href=\"#c10\">10</a>]</sup></th>");
            builder.Append("<th><b>CSEO Assessment Grade Review In Progress</b><sup> [<a href=\"#c11\">11</a>]</sup></th>");
            builder.Append("<th bgcolor=\"#92D050\"><b>CSEO Assessment Grade Review Completed</b><sup> [<a href=\"#c12\">12</a>]</sup></th>");
            builder.Append("<th bgcolor=\"#A8D08D\"><b>Bug Remediation Post CSEO Assessment Not Started</b><sup> [<a href=\"#c13\">13</a>]</sup></th>");
            builder.Append("<th bgcolor=\"#A8D08D\"><b>Bug Remediation Post CSEO Assessment In Progress</b><sup> [<a href=\"#c14\">14</a>]</sup></th>");
            builder.Append("<th bgcolor=\"#00B050\"><b>CSEO Assessment Grade C Received</b><sup> [<a href=\"#c15\">15</a>]</sup></th>");
            builder.Append("<th><b>P3 Applications User Story Count</b></th>");
            builder.Append("</tr>");
            

            //distinct SubGroups and Service Lines
            DataTable distinct_SubGroup = allP3UserStories.DefaultView.ToTable(true, "Subgroup"); // get unique SubGroup column name
            DataTable distinct_ServiceOffering = allP3UserStories.DefaultView.ToTable(true, "ServiceOffering"); //Get the unique Service Offering Names

            int condition_1_UserStoryCount = 0;
            int condition_2_UserStoryCount = 0;
            int condition_3_UserStoryCount = 0;
            int condition_4_UserStoryCount = 0;
            int condition_5_UserStoryCount = 0;
            int condition_6_UserStoryCount = 0;
            int condition_7_UserStoryCount = 0;
            int condition_8_UserStoryCount = 0;
            int condition_9_UserStoryCount = 0;
            int condition_10_UserStoryCount = 0;
            int condition_11_UserStoryCount = 0;
            int condition_12_UserStoryCount = 0;
            int condition_13_UserStoryCount = 0;
            int condition_14_BugRemediationPostCSEOAssessmentNotStarted = 0;
            int condition_15_BugRemediationPostCSEOAssessmentInProgress = 0;
            int condition_7New_UserStoryCount = 0;

            foreach (DataRow SubGroup in distinct_SubGroup.Rows)
            {
                foreach (var itemSubGroup in SubGroup.ItemArray)
                {
                    string sugroupStrFmt = itemSubGroup.ToString();

                    var userIDTb_2_2 = (from rTb2_2 in allP3UserStories.AsEnumerable()
                                        where rTb2_2.Field<string>("Subgroup") == sugroupStrFmt
                                        && rTb2_2.Field<string>("ServiceOffering") != null
                                        select new
                                        {
                                            workItemID_2_2 = rTb2_2.Field<string>("ServiceOffering")
                                        });
                    var orderedUserIDTb_2_2 = userIDTb_2_2.OrderBy(fTb2_2 => fTb2_2.workItemID_2_2).Distinct();
                   
                    foreach (var r in orderedUserIDTb_2_2)
                    {
                        builder.Append("<tr class=\"txtcc\">");
                        builder.Append("<td class=\"txtcc\"><b>" + sugroupStrFmt + "</b></td>");
                        string serviceOfferingName = r.workItemID_2_2;
                        builder.Append("<td class=\"txtcc\"><b>"+ r.workItemID_2_2 + "</b></td>");
                        
                        //Condition-1: Self Assessment To Start [Table:dt_SelfAssessmentToStart]
                        var condition1Data = (from rSubCond in dt_SelfAssessmentToStart.AsEnumerable()
                                              where rSubCond.Field<string>("SubGroup") == sugroupStrFmt
                                              && (rSubCond.Field<string>("ServiceLine") == serviceOfferingName)
                                              select rSubCond.Field<string>("WorkItemID"));

                        condition_1_UserStoryCount = condition1Data.Distinct().Count();
                        if(condition_1_UserStoryCount==0)
                        {
                            builder.Append("<td class=\"text-td\">" + condition_1_UserStoryCount + "</td>");
                        }
                        else
                        {
                            builder.Append("<td class=\"text-td\"><b>" + condition_1_UserStoryCount + "</b></td>");
                        }
                        

                        //Condition-2: Self-Assessment In Progress [Table: dt_SelfAssessmentInProgress
                        var condition2Data = (from rSubCond2 in dt_SelfAssessmentInProgress.AsEnumerable()
                                                       where rSubCond2.Field<string>("SubGroup") == sugroupStrFmt
                                                       && (rSubCond2.Field<string>("ServiceLine") == serviceOfferingName)
                                                       select rSubCond2.Field<string>("WorkItemID"));
                        condition_2_UserStoryCount = condition2Data.Distinct().Count();
                        if(condition_2_UserStoryCount==0)
                        {
                            builder.Append("<td class=\"text-td\">" + condition_2_UserStoryCount + "</td>");
                        } else
                        {
                            builder.Append("<td class=\"text-td\"><b>" + condition_2_UserStoryCount + "</b></td>");
                        }
                        

                        //Condition-3: Self-Assessment Completed [Table:dt_SelfAssessmentCompleted
                        var condition3Data = (from rSubCond3 in dt_SelfAssessmentCompleted.AsEnumerable()
                                              where rSubCond3.Field<string>("SubGroup") == sugroupStrFmt
                                              && (rSubCond3.Field<string>("ServiceLine") == serviceOfferingName)
                                              select rSubCond3.Field<string>("WorkItemID"));
                        condition_3_UserStoryCount = condition3Data.Distinct().Count();
                        if(condition_3_UserStoryCount==0)
                        {
                            builder.Append("<td bgcolor=\"#FFF2CC\" class=\"text-td\">" + condition_3_UserStoryCount + "</td>");
                        }
                        else
                        {
                            builder.Append("<td bgcolor=\"#FFF2CC\" class=\"text-td\"><b>" + condition_3_UserStoryCount + "</b></td>");
                        }

                        //dt_BugRemediationPostSelfAssessmentNotStarted
                        //Condition-4: Bug Remediation Post Self Assessment Not Started [Table: dt_BugRemediationPostSelfAssessmentNotStarted]
                        var condition4Data = (from rSubCond4 in dt_BugRemediationPostSelfAssessmentNotStarted.AsEnumerable()
                                              where rSubCond4.Field<string>("SubGroup") == sugroupStrFmt
                                              && (rSubCond4.Field<string>("ServiceLine") == serviceOfferingName)
                                              select rSubCond4.Field<string>("WorkItemID"));
                        condition_4_UserStoryCount = condition4Data.Distinct().Count();
                        int testCOunt = condition4Data.Distinct().Count();

                        if (condition_4_UserStoryCount==0)
                        {
                            builder.Append("<td class=\"text-td\">" + condition_4_UserStoryCount + "</td>");
                        }
                        else
                        {
                            builder.Append("<td class=\"text-td\"><b>" + testCOunt + "</b></td>");
                        }


                        //Condition-5: Bug Remediation Post Self Assessment Started [Table: dt_BugRemediationPostSelfAssessmentStarted]
                        var condition5Data = (from rSubCond5 in dt_BugRemediationPostSelfAssessmentStarted.AsEnumerable()
                                              where rSubCond5.Field<string>("SubGroup") == sugroupStrFmt
                                              && (rSubCond5.Field<string>("ServiceLine") == serviceOfferingName)
                                              select rSubCond5.Field<string>("WorkItemID"));
                        condition_5_UserStoryCount = condition5Data.Distinct().Count();
                        if(condition_5_UserStoryCount==0)
                        {
                            builder.Append("<td class=\"text-td\">" + condition_5_UserStoryCount + "</td>");
                        }
                        else
                        {
                            builder.Append("<td class=\"text-td\"><b>" + condition_5_UserStoryCount + "</b></td>");
                        }
                       

                        //Condition-6: Bug Remediation Post Self Assessment Completed [Table: dt_BugRemediationPostSelfAssessmentCompleted]
                        var condition6Data = (from rSubCond6 in dt_BugRemediationPostSelfAssessmentCompleted.AsEnumerable()
                                              where rSubCond6.Field<string>("SubGroup") == sugroupStrFmt
                                              && (rSubCond6.Field<string>("ServiceLine") == serviceOfferingName)
                                              select rSubCond6.Field<string>("WorkItemID"));
                        condition_6_UserStoryCount = condition6Data.Distinct().Count();
                        
                        if(condition_6_UserStoryCount==0)
                        {
                            builder.Append("<td bgcolor=\"#ffe699\" class=\"text-td\">" + condition_6_UserStoryCount + "</td>");
                        } else
                        {
                            builder.Append("<td bgcolor=\"#ffe699\" class=\"text-td\"><b>" + condition_6_UserStoryCount + "</b></td>");
                        }
                        
                        //Condition:7 New
                        var condition7NewData = (from rSubCond7_New in dt_awaitingOnboardingDocumentPostSelfAssessment.AsEnumerable()
                                                 where rSubCond7_New.Field<string>("SubGroup") == sugroupStrFmt
                                                 && (rSubCond7_New.Field<string>("ServiceLine") == serviceOfferingName)
                                                 select rSubCond7_New.Field<string>("WorkItemID"));
                        condition_7New_UserStoryCount = condition7NewData.Distinct().Count();
                        if (condition_7New_UserStoryCount == 0)
                        {
                            builder.Append("<td class=\"text-td\">" + condition_7New_UserStoryCount + "</td>");
                        }
                        else
                        {
                            builder.Append("<td class=\"text-td\"><b>" + condition_7New_UserStoryCount + "</b></td>");
                        }

                        //Condition-8: Ready for Assessment Service Onboarding [Table: dt_ReadyForAssessmentServiceOnboarding]
                        var condition7Data = (from rSubCond7 in dt_ReadyForAssessmentServiceOnboarding.AsEnumerable()
                                              where rSubCond7.Field<string>("SubGroup") == sugroupStrFmt
                                              && (rSubCond7.Field<string>("ServiceLine") == serviceOfferingName)
                                              select rSubCond7.Field<string>("WorkItemID"));
                        condition_7_UserStoryCount = condition7Data.Distinct().Count();
                        
                        if(condition_7_UserStoryCount==0)
                        {
                            builder.Append("<td class=\"text-td\">" + condition_7_UserStoryCount + "</td>");
                        }
                        else
                        {
                            builder.Append("<td class=\"text-td\"><b>" + condition_7_UserStoryCount + "</b></td>");
                        }

                        //Condition-9: Assessment Service Onboarding In Progress [Table: dt_AssessmentServiceOnboardingInProgress]
                        var condition8Data = (from rSubCond8 in dt_AssessmentServiceOnboardingInProgress.AsEnumerable()
                                              where rSubCond8.Field<string>("SubGroup") == sugroupStrFmt
                                              && (rSubCond8.Field<string>("ServiceLine") == serviceOfferingName)
                                              select rSubCond8.Field<string>("WorkItemID"));
                        condition_8_UserStoryCount = condition8Data.Distinct().Count();
                        
                        if(condition_8_UserStoryCount==0)
                        {
                            builder.Append("<td class=\"text-td\">" + condition_8_UserStoryCount + "</td>");
                        }
                        else
                        {
                            builder.Append("<td class=\"text-td\"><b>" + condition_8_UserStoryCount + "</b></td>");
                        }

                        //Condition-10: Scheduled for Grade Review Assessment [Table: dt_ScheduledForGradeReviewAssessment]
                        var condition9Data = (from rSubCond9 in dt_ScheduledForGradeReviewAssessment.AsEnumerable()
                                              where rSubCond9.Field<string>("SubGroup") == sugroupStrFmt
                                              && (rSubCond9.Field<string>("ServiceLine") == serviceOfferingName)
                                              select rSubCond9.Field<string>("WorkItemID"));
                        condition_9_UserStoryCount = condition9Data.Distinct().Count();
                        
                        if(condition_9_UserStoryCount==0)
                        {
                            builder.Append("<td bgcolor=\"#FFD966\" class=\"text-td\">" + condition_9_UserStoryCount + "</td>");
                        }
                        else
                        {
                            builder.Append("<td bgcolor=\"#FFD966\" class=\"text-td\"><b>" + condition_9_UserStoryCount + "</b></td>");
                        }

                        //Condition-11: Grade Review Assessment in Progress [Table:dt_GradeReviewAssessmentInProgress]
                        var condition10Data = (from rSubCond10 in dt_GradeReviewAssessmentInProgress.AsEnumerable()
                                              where rSubCond10.Field<string>("SubGroup") == sugroupStrFmt
                                              && (rSubCond10.Field<string>("ServiceLine") == serviceOfferingName)
                                              select rSubCond10.Field<string>("WorkItemID"));
                        condition_10_UserStoryCount = condition10Data.Distinct().Count();
                        
                        if(condition_10_UserStoryCount==0)
                        {
                            builder.Append("<td class=\"text-td\">" + condition_10_UserStoryCount + "</td>");
                        }
                        else
                        {
                            builder.Append("<td class=\"text-td\"><b>" + condition_10_UserStoryCount + "</b></td>");
                        }

                        //Condition-12: Grade Review Assessment Completed [Table: dt_GradeReviewAssessmentCompleted]
                        var condition11Data = (from rSubCond11 in dt_GradeReviewAssessmentCompleted.AsEnumerable()
                                               where rSubCond11.Field<string>("SubGroup") == sugroupStrFmt
                                               && (rSubCond11.Field<string>("ServiceLine") == serviceOfferingName)
                                               select rSubCond11.Field<string>("WorkItemID"));
                        condition_11_UserStoryCount = condition11Data.Distinct().Count();

                        if (condition_11_UserStoryCount == 0)
                        {
                            builder.Append("<td bgcolor=\"#92D050\" class=\"text-td\">" + condition_11_UserStoryCount + "</td>");
                        }
                        else
                        {
                            builder.Append("<td bgcolor=\"#92D050\" class=\"text-td\"><b>" + condition_11_UserStoryCount + "</b></td>");
                        }

                        /*Condition-13: If Task ([Eng][Activity: Bug Remediation and Verification post Grade Review]) is in 
                        New state - User Story count aggregated to show the numbers in “Bug Remediation Post CSEO Assessment 
                        Not Started" */
                        var condition14Data = (from rSubCond14 in dt_BugRemediationPostCSEOAssessmentNotStarted.AsEnumerable()
                                               where rSubCond14.Field<string>("SubGroup") == sugroupStrFmt
                                               && (rSubCond14.Field<string>("ServiceLine") == serviceOfferingName)
                                               select rSubCond14.Field<string>("WorkItemID"));
                        condition_14_BugRemediationPostCSEOAssessmentNotStarted = condition14Data.Distinct().Count();

                        if (condition_14_BugRemediationPostCSEOAssessmentNotStarted == 0)
                        {
                            builder.Append("<td bgcolor=\"#A8D08D\" class=\"text-td\">" + condition_14_BugRemediationPostCSEOAssessmentNotStarted + "</td>");
                        }
                        else
                        {
                            builder.Append("<td bgcolor=\"#A8D08D\" class=\"text-td\"><b>" + condition_14_BugRemediationPostCSEOAssessmentNotStarted + "</b></td>");
                        }


                        //Condition-14:	If Task ([Eng][Activity: Bug Remediation and Verification post Grade Review]) is in Active state - User Story count aggregated to show the numbers in " Bug Remediation Post CSEO Assessment In Progress"
                        var condition15Data = (from rSubCond14 in dt_BugRemediationPostCSEOAssessmentInProgress.AsEnumerable()
                                               where rSubCond14.Field<string>("SubGroup") == sugroupStrFmt
                                               && (rSubCond14.Field<string>("ServiceLine") == serviceOfferingName)
                                               select rSubCond14.Field<string>("WorkItemID"));
                        condition_15_BugRemediationPostCSEOAssessmentInProgress = condition15Data.Distinct().Count();

                        if (condition_15_BugRemediationPostCSEOAssessmentInProgress == 0)
                        {
                            builder.Append("<td bgcolor=\"#A8D08D\" class=\"text-td\">" + condition_15_BugRemediationPostCSEOAssessmentInProgress + "</td>");
                        }
                        else
                        {
                            builder.Append("<td bgcolor=\"#A8D08D\" class=\"text-td\"><b>" + condition_15_BugRemediationPostCSEOAssessmentInProgress + "</b></td>");
                        }


                        //Condition-15: P3 Applications Grade C Applications Count [Table: dt_FY20GradeC]
                        var condition12Data = (from rSubCond12 in dt_FY20GradeC.AsEnumerable()
                                               where rSubCond12.Field<string>("SubGroup") == sugroupStrFmt
                                               && (rSubCond12.Field<string>("ServiceLine") == serviceOfferingName)
                                               select rSubCond12.Field<string>("WorkItemID"));
                        condition_12_UserStoryCount = condition12Data.Distinct().Count();
                        
                        if(condition_12_UserStoryCount==0)
                        {
                            builder.Append("<td bgcolor=\"#00B050\" class=\"text-td\">" + condition_12_UserStoryCount + "</td>");
                        }
                        else
                        {
                            builder.Append("<td bgcolor=\"#00B050\" class=\"text-td\"><b>" + condition_12_UserStoryCount + "</b></td>");
                        }

                        //OveralCount: P3 Applications User Story Count [Table: dt_FY20UserStoryCount]
                        var condition13Data = (from rSubCond13 in allP3UserStories.AsEnumerable()
                                               where rSubCond13.Field<string>("Subgroup") == sugroupStrFmt
                                               && (rSubCond13.Field<string>("ServiceOffering") == serviceOfferingName)
                                               && (rSubCond13.Field<string>("UserStoryState") == "New" ||
                                               rSubCond13.Field<string>("UserStoryState") == "Active" ||
                                               rSubCond13.Field<string>("UserStoryState") == "Closed")
                                               select rSubCond13.Field<string>("WorkItemID"));
                        condition_13_UserStoryCount = condition13Data.Distinct().Count();
                        builder.Append("<td class=\"text-td\">" + condition_13_UserStoryCount + "</td>");


                        //dt_FY20GradeC.Columns.Add("WorkItemID", typeof(string));
                        //dt_FY20GradeC.Columns.Add("SubGroup", typeof(string));
                        //dt_FY20GradeC.Columns.Add("ServiceLine", typeof(string));
                        //dt_FY20GradeC.Columns.Add("Grade", typeof(string));

                        //dt_FY20UserStoryCount.Columns.Add("WorkItemID", typeof(string));
                        //dt_FY20UserStoryCount.Columns.Add("SubGroup", typeof(string));
                        //dt_FY20UserStoryCount.Columns.Add("ServiceLine", typeof(string));
                        builder.Append("</tr>");

                    }
                    
                }
            }

            builder.Append("<tr class=\"percentage\">");
            builder.Append("<th colspan=\"2\" scope=\"colgroup\">Total / Percentage</th>");
            builder.Append("<th>" + dt_SelfAssessmentToStart.Rows.Count + " / "+ (int)Math.Round((double)(100 * dt_SelfAssessmentToStart.Rows.Count) / allP3UserStories.Rows.Count) +"%"+"</th>");
            builder.Append("<th>" + dt_SelfAssessmentInProgress.Rows.Count + " / " + (int)Math.Round((double)(100 * dt_SelfAssessmentInProgress.Rows.Count) / allP3UserStories.Rows.Count) + "%" + "</th>");
            builder.Append("<th bgcolor=\"#FFF2CC\">" + dt_SelfAssessmentCompleted.Rows.Count + " / " + (int)Math.Round((double)(100 * dt_SelfAssessmentCompleted.Rows.Count) / allP3UserStories.Rows.Count) + "%" + "</th>");
            builder.Append("<th>" + dt_BugRemediationPostSelfAssessmentNotStarted.Rows.Count + " / " + (int)Math.Round((double)(100 * dt_BugRemediationPostSelfAssessmentNotStarted.Rows.Count) / allP3UserStories.Rows.Count) + "%" + "</th>");
            builder.Append("<th>" + dt_BugRemediationPostSelfAssessmentStarted.Rows.Count + " / " + (int)Math.Round((double)(100 * dt_BugRemediationPostSelfAssessmentStarted.Rows.Count) / allP3UserStories.Rows.Count) + "%" + "</th>");
            builder.Append("<th bgcolor=\"#ffe699\">" + dt_BugRemediationPostSelfAssessmentCompleted.Rows.Count + " / " + (int)Math.Round((double)(100 * dt_BugRemediationPostSelfAssessmentCompleted.Rows.Count) / allP3UserStories.Rows.Count) + "%" + "</th>");
            builder.Append("<th>" + dt_awaitingOnboardingDocumentPostSelfAssessment.Rows.Count + " / " + (int)Math.Round((double)(100 * dt_awaitingOnboardingDocumentPostSelfAssessment.Rows.Count) / allP3UserStories.Rows.Count) + "%" + "</th>");
            builder.Append("<th>" + dt_ReadyForAssessmentServiceOnboarding.Rows.Count + " / " + (int)Math.Round((double)(100 * dt_ReadyForAssessmentServiceOnboarding.Rows.Count) / allP3UserStories.Rows.Count) + "%" + "</th>");
            builder.Append("<th>" + dt_AssessmentServiceOnboardingInProgress.Rows.Count + " / " + (int)Math.Round((double)(100 * dt_AssessmentServiceOnboardingInProgress.Rows.Count) / allP3UserStories.Rows.Count) + "%" + "</th>");
            builder.Append("<th bgcolor=\"#FFD966\">" + dt_ScheduledForGradeReviewAssessment.Rows.Count + " / " + (int)Math.Round((double)(100 * dt_ScheduledForGradeReviewAssessment.Rows.Count) / allP3UserStories.Rows.Count) + "%" + "</th>");
            builder.Append("<th>" + dt_GradeReviewAssessmentInProgress.Rows.Count + " / " + (int)Math.Round((double)(100 * dt_GradeReviewAssessmentInProgress.Rows.Count) / allP3UserStories.Rows.Count) + "%" + "</th>");
            builder.Append("<th bgcolor=\"#91CF50\">" + dt_GradeReviewAssessmentCompleted.Rows.Count + " / " + (int)Math.Round((double)(100 * dt_GradeReviewAssessmentCompleted.Rows.Count) / allP3UserStories.Rows.Count) + "%" + "</th>");
            builder.Append("<th bgcolor=\"#A8D08D\">" + dt_BugRemediationPostCSEOAssessmentNotStarted.Rows.Count + " / " + (int)Math.Round((double)(100 * dt_BugRemediationPostCSEOAssessmentNotStarted.Rows.Count) / allP3UserStories.Rows.Count) + "%" + "</th>");
            builder.Append("<th bgcolor=\"#A8D08D\">" + dt_BugRemediationPostCSEOAssessmentInProgress.Rows.Count + " / " + (int)Math.Round((double)(100 * dt_BugRemediationPostCSEOAssessmentInProgress.Rows.Count) / allP3UserStories.Rows.Count) + "%" + "</th>");
            builder.Append("<th bgcolor=\"#00B050\">" + dt_FY20GradeC.Rows.Count + " / " + (int)Math.Round((double)(100 * dt_FY20GradeC.Rows.Count) / allP3UserStories.Rows.Count) + "%" + "</th>");
            builder.Append("<th>" + allP3UserStories.Rows.Count + "</th>");
            builder.Append("</tr>");
            builder.Append("</table>");
            builder.Append("<br/>" +
                "<ol lang=\"EN-IN\" style=\"font-size:9.0pt;font-family:&quot;Segoe UI&quot;,sans-serif\">" +
                "<li><span lang=\"EN-IN\" style=\"font-size:9.0pt;font-family:&quot;Segoe UI&quot;,sans-serif\"><a name=\"c1\">If Task </a> (<b><i>[Eng][Activity: Self-assessment and Bugs Logging</i></b>) is in <b>New</b> state – User Story count aggregated to show the numbers in &quot;<b>Self-Assessment Not Started</b>&quot; <o:p></o:p></span></li>" +
                "<li><span lang=\"EN-IN\" style=\"font-size:9.0pt;font-family:&quot;Segoe UI&quot;,sans-serif\"><a name=\"c2\">If Task </a> (<b><i>[Eng][Activity: Self-assessment and Bugs Logging</i></b>) is in <b>Active</b> state - User Story count aggregated to show the numbers in &quot;<b>Self-Assessment In Progress</b>&quot;<o:p></o:p></span></li>" +
                "<li><span lang=\"EN-IN\" style=\"font-size:9.0pt;font-family:&quot;Segoe UI&quot;,sans-serif\"><a name=\"c3\">If Task </a> (<b><i>[Eng][Activity: Self-assessment and Bugs Logging</i></b>) is in <b>Closed</b> state - User Story count aggregated to show the numbers in &quot;<b>Self-Assessment Completed</b>&quot; <o:p></o:p></span></li>" +
                //"<li><span lang=\"EN-IN\" style=\"font-size:9.0pt;font-family:&quot;Segoe UI&quot;,sans-serif\"><a name=\"c4\">If Task </a> (<b><i>[Eng][Activity: Bugs Remediation post Self-assessment]</i></b>) is in <b>New</b> state - User Story count aggregated to show the numbers in “<b>Bug Remediation Post Self-Assessment Not Started</b>&quot;<o:p></o:p></span></li>" +
                "<li><span lang=\"EN-IN\" style=\"font-size:9.0pt;font-family:&quot;Segoe UI&quot;,sans-serif\"><a name=\"c4\">If Task </a> (<b><i>[Eng][Activity: Self-assessment and Bugs Logging</i></b>) is in <b>Closed</b> state and </span><span lang=\"EN-IN\" style=\"font-size:9.0pt;font-family:&quot;Segoe UI&quot;,sans-serif\">If Task (<b><i>[Eng][Activity: Bugs Remediation post Self-assessment]</i></b>) is in <b>New</b> state - User Story count aggregated to show the numbers in “<b>Bug Remediation Post Self-Assessment Not Started</b>&quot;<o:p></o:p></span></li>" +
                "<li><span lang=\"EN-IN\" style=\"font-size:9.0pt;font-family:&quot;Segoe UI&quot;,sans-serif\"><a name=\"c5\">If Task</a> (<b><i>[Eng][Activity: Bugs Remediation post Self-assessment]</i></b>) is in <b>Active</b> state - User Story count aggregated to show the numbers in &quot;<b>Bug Remediation Post Self-Assessment In Progress</b>&quot;<o:p></o:p></span></li>" +
                "<li><span lang=\"EN-IN\" style=\"font-size:9.0pt;font-family:&quot;Segoe UI&quot;,sans-serif\"><a name=\"c6\">If Task</a> (<b><i>[Eng][Activity: Bugs Remediation post Self-assessment]</i></b>) is in <b>Closed</b> state - User Story count aggregated to show the numbers in &quot;<b>Bug Remediation Post Self-Assessment Completed</b>&quot; <o:p></o:p></span></li>" +

                "<li><span lang=\"EN-IN\" style=\"font-size:9.0pt;font-family:&quot;Segoe UI&quot;,sans-serif\"><a name=\"c7\">If Task</a> (<b><i>([Eng][Activity: Bugs Remediation post Self-assessment])</i></b><i>)<b> </b></i>is <b>Closed</b> &amp; If Task (<b><i>([Eng][Activity: Create Grade Review Onboarding Request])</i></b><i>)</i> is not in <b>Closed</b> State - User Story count aggregated to show the numbers in “<b>Awaiting Onboarding Document Post Self-Assessment</b>&quot; <o:p></o:p></span></li>" +
                "<li><span lang=\"EN-IN\" style=\"font-size:9.0pt;font-family:&quot;Segoe UI&quot;,sans-serif\"><a name=\"c8\">If Task</a> (<b><i>[Eng][Activity: Create Grade Review Onboarding Request]</i></b><i>)<b> </b></i>is <b>Closed</b> &amp; If Task (<b><i>[Activity:Onboarding]</i></b><i>)</i> is in <b>New</b> State - User Story count aggregated to show the numbers in “<b>CSEO</b> <b>Assessment Onboarding Not Started</b>&quot; <o:p></o:p></span></li>" +
                "<li><span lang=\"EN-IN\" style=\"font-size:9.0pt;font-family:&quot;Segoe UI&quot;,sans-serif\"><a name=\"c9\">If Task</a> (<b><i>[Activity:Onboarding]</i></b><i>)</i> is in <b>Active</b> State - User Story count aggregated to show the numbers &nbsp;in &quot;<b> CSEO</b> <b>Assessment Onboarding In Progress</b> &quot;<o:p></o:p></span></li>" +
                //"<li><span lang=\"EN-IN\" style=\"font-size:9.0pt;font-family:&quot;Segoe UI&quot;,sans-serif\"><a name=\"c9\">If Task</a> (<b><i>[Activity:Onboarding]) </i></b>is in <b>Closed</b> state and &nbsp;If Task (<b><i>[Activity:Assessment]</i></b>) is in <b>New</b> State - User Story count aggregated to show the numbers &nbsp;in &quot;<b>CSEO Assessment Onboarding Completed</b>&quot;<o:p></o:p></span></li>" +
                //"<li><span lang=\"EN-IN\" style=\"font-size:9.0pt;font-family:&quot;Segoe UI&quot;,sans-serif\"><a name=\"c9\">If Task</a> (<b><i>[Activity:Onboarding]) </i></b>is in <b>Closed</b> state - User Story count aggregated to show the numbers &nbsp;in &quot;<b>CSEO Assessment Onboarding Completed</b>&quot;<o:p></o:p></span></li>" +
                "<li><span lang=\"EN-IN\" style=\"font-size:9.0pt;font-family:&quot;Segoe UI&quot;,sans-serif\"><a name=\"c10\">If Task</a> (<b><i>[Activity:Onboarding]) </i></b>is in <b>Closed</b> state - User Story count aggregated to show the numbers &nbsp;in &quot;<b>CSEO Assessment Onboarding Completed</b>&quot;<o:p></o:p></span></li>" +
                "<li><span lang=\"EN-IN\" style=\"font-size:9.0pt;font-family:&quot;Segoe UI&quot;,sans-serif\"><a name=\"c11\">If Task</a> (<b><i>[Activity:Assessment]</i></b>) is in <b>Active</b> state - User Story count aggregated to show the numbers &nbsp;in &quot;<b>CSEO Assessment Grade Review In Progress</b>&quot; <o:p></o:p></span></li>" +
                "<li><span lang=\"EN-IN\" style=\"font-size:9.0pt;font-family:&quot;Segoe UI&quot;,sans-serif\"><a name=\"c12\">If Task</a> (<b><i>[Activity:Assessment]</i></b>) is in <b>Closed</b> state - User Story count aggregated to show the numbers &nbsp;in &quot;<b>CSEO Assessment Grade Review Completed</b>&quot; <o:p></o:p></span></li>" +
                //"<li><span style=\"font-size:9.0pt;font-family:&quot;Segoe UI&quot;,sans-serif\"><a name=\"c12\">If Task</a> (<b><i>[Eng][Activity: Bug Remediation and Verification post Grade Review]</i></b>) is in <b>New</b> state - User Story count aggregated to show the numbers in “<b>Bug Remediation Post CSEO Assessment Not Started</b> </span></li>" +
                "<li><span lang=\"EN-IN\" style=\"font-size:9.0pt;font-family:&quot;Segoe UI&quot;,sans-serif\"><a name=\"c13\">If Task</a> (<b><i>[Activity:Assessment]</i></b>) is in <b>Closed</b> state and </span> <span style=\"font-size:9.0pt;font-family:&quot;Segoe UI&quot;,sans-serif\">If Task (<b><i>[Eng][Activity: Bug Remediation and Verification post Grade Review]</i></b>) is in <b>New</b> state - User Story count aggregated to show the numbers in “<b>Bug Remediation Post CSEO Assessment Not Started</b> </span></li>" +
                "<li><span><a name=\"c14\">If Task</a> (<b><i>[Eng][Activity: Bug Remediation and Verification post Grade Review]</i></b>) is in <b>Active</b> state - User Story count aggregated to show the numbers in &quot;<b> Bug Remediation Post CSEO Assessment In Progress</b>&quot;</span></li>" +
                "<li><span style=\"font-size:9.0pt;font-family:&quot;Segoe UI&quot;,sans-serif\"><a name=\"c15\">If Task</a> (<b><i>[Eng][Activity: Bug Remediation and Verification post Grade Review]</i></b>) is in <b>Closed</b> state - User Story count aggregated to show the numbers in &quot;<b> CSEO Assessment Grade C Received</b>&quot;</span></li></ol>");
            builder.Append("</body>");
            builder.Append("</html>");
            string HtmlFile = builder.ToString();
            oMsg.Attachments.Add(Path.GetFullPath("D:\\Reports Excel\\AppP3UserStories\\P3UserStories.xlsx"));
            //Attachment emailAttachment = oMsg.Attachments.Add("D:\\Reports Excel\\AppP3UserStories\\P3UserStories.xlsx", Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue);
            
            //QA Test
            //oMsg.To = ConfigurationManager.AppSettings["Prod_To_Draft_P3ApplicationSend"];

            //QA Draft
            //oMsg.To = ConfigurationManager.AppSettings["Prod_To_Draft_P3ApplicationSend"];
            

            //Prod
            oMsg.To = ConfigurationManager.AppSettings["Prod_To_P3ApplicationSend"];
            oMsg.CC = ConfigurationManager.AppSettings["Prod_CC_P3ApplicationSend"];

            DateTime startAtMonday = DateTime.Now.AddDays(DayOfWeek.Monday - DateTime.Now.DayOfWeek);
            oMsg.Subject = "CSEO Assessment Service: CSEO FY20 P3 Accessibility Progress Scorecard - " + DateTime.Now.ToString("MM/dd/yyyy");
            oMsg.HTMLBody = HtmlFile;
            Console.WriteLine("ADO Features are created and sent mail to respective folks");
            oMsg.Send();
        }
    }
}
