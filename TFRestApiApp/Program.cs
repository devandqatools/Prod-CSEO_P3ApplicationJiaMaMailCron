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
        static readonly string UserPAT = "fcigllokhrdkt5njzrgchjkszypluxmfcwcigeu5edlkizvif63q";
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

        public static DataTable dt_SelfAssessmentToStart = new DataTable();
        public static DataTable dt_SelfAssessmentInProgress = new DataTable();
        public static DataTable dt_SelfAssessmentCompleted = new DataTable();
        public static DataTable dt_BugRemediationPostSelfAssessmentNotStarted = new DataTable();
        public static DataTable dt_BugRemediationPostSelfAssessmentStarted = new DataTable();
        public static DataTable dt_BugRemediationPostSelfAssessmentCompleted = new DataTable();
        public static DataTable dt_ReadyForAssessmentServiceOnboarding = new DataTable();
        public static DataTable dt_AssessmentServiceOnboardingInProgress = new DataTable();
        public static DataTable dt_ScheduledForGradeReviewAssessment = new DataTable();
        public static DataTable dt_GradeReviewAssessmentInProgress = new DataTable();
        public static DataTable dt_GradeReviewAssessmentCompleted = new DataTable();
        public static DataTable dt_FY20GradeC = new DataTable();
        public static DataTable dt_FY20UserStoryCount = new DataTable();

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

            dt_FY20GradeC.Columns.Add("WorkItemID", typeof(string));
            dt_FY20GradeC.Columns.Add("SubGroup", typeof(string));
            dt_FY20GradeC.Columns.Add("ServiceLine", typeof(string));
            dt_FY20GradeC.Columns.Add("Grade", typeof(string));

            dt_FY20UserStoryCount.Columns.Add("WorkItemID", typeof(string));
            dt_FY20UserStoryCount.Columns.Add("SubGroup", typeof(string));
            dt_FY20UserStoryCount.Columns.Add("ServiceLine", typeof(string));

            /*_______________________________End__________________________________*/


            ConnectWithPAT(TFUrl, UserPAT);

            //Get user stories for P3 applications

            string queryWiqlList = @"select [System.Id], [System.WorkItemType], [System.Title], [System.AssignedTo], [System.State], [System.Tags] from WorkItemLinks where (Source.[System.TeamProject] = @project and Source.[System.WorkItemType] = 'Scenario' and Source.[System.Id] = 5324882) and ([System.Links.LinkType] = 'System.LinkTypes.Hierarchy-Forward') and (Target.[System.TeamProject] = @project and Target.[System.WorkItemType] in ('Feature', 'User Story') and Target.[System.State] <> 'Removed') order by [System.Title] mode (Recursive)";
//            string queryWiqlList = @"SELECT
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
//        AND [Target].[System.Id] IN (5576002)
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
                                     where r.Field<string>("Grade") == "C"
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

                                    //SelfAssessmentToStart: If Task (Self Assessment and Bug Logging) is in New
                                    if (taskTitle.Contains("assessment and bugs logging") && taskFieldValue_TaskState.ToString() == "New")
                                    {
                                        string cndt1_userStoryID = wiID.ToString();
                                        string cndt1_ServiceOffering = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt1_userStoryID)["ServiceOffering"].ToString());
                                        string cndt1_SubGroup = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt1_userStoryID)["Subgroup"].ToString());

                                        dt_SelfAssessmentToStart.Rows.Add(cndt1_userStoryID, cndt1_SubGroup, cndt1_ServiceOffering);
                                        Console.WriteLine("_______________Condition-1 DATA Updated[SelfAssessmentToStart]_________________________");

                                        //Console.WriteLine("{0}:{1} Identified Condition-1:SelfAssessmentToStart..!", taskFieldValue_TaskWorkItemType.ToString(), taskID);
                                    }

                                    //SelfAssessmentInProgress
                                    if (taskTitle.Contains("assessment and bugs logging") && taskFieldValue_TaskState.ToString() == "Active")
                                    {
                                        string cndt2_userStoryID = wiID.ToString();
                                        string cndt2_ServiceOffering = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt2_userStoryID)["ServiceOffering"].ToString());
                                        string cndt2_SubGroup = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt2_userStoryID)["Subgroup"].ToString());

                                        dt_SelfAssessmentInProgress.Rows.Add(cndt2_userStoryID, cndt2_SubGroup, cndt2_ServiceOffering);

                                        Console.WriteLine("Condition_2 SelfAssessmentInProgress identified.....!");
                                    }
                                    //SelfAssessmentCompleted
                                    if (taskTitle.Contains("assessment and bugs logging") && taskFieldValue_TaskState.ToString() == "Closed")
                                    {
                                        string cndt3_userStoryID = wiID.ToString();
                                        string cndt3_ServiceOffering = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt3_userStoryID)["ServiceOffering"].ToString());
                                        string cndt3_SubGroup = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt3_userStoryID)["Subgroup"].ToString());

                                        dt_SelfAssessmentCompleted.Rows.Add(cndt3_userStoryID, cndt3_SubGroup, cndt3_ServiceOffering);

                                        Console.WriteLine("Condition_3 SelfAssessmentCompleted identified.....!");
                                    }
                                    //Bug Remediation Post Self Assessment Not Started
                                    if (taskTitle.Contains("bugs remediation post self-assessment") && taskFieldValue_TaskState.ToString() == "New")
                                    {
                                        string cndt4_userStoryID = wiID.ToString();
                                        string cndt4_ServiceOffering = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt4_userStoryID)["ServiceOffering"].ToString());
                                        string cndt4_SubGroup = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt4_userStoryID)["Subgroup"].ToString());

                                        dt_BugRemediationPostSelfAssessmentNotStarted.Rows.Add(cndt4_userStoryID, cndt4_SubGroup, cndt4_ServiceOffering);

                                        Console.WriteLine("Condition_4 Bug Remediation Post Self Assessment Not Started identified.....!");
                                    }
                                    //Bug Remediation Post Self Assessment Started
                                    if (taskTitle.Contains("bugs remediation post self-assessment") && taskFieldValue_TaskState.ToString() == "Active")
                                    {
                                        string cndt5_userStoryID = wiID.ToString();
                                        string cndt5_ServiceOffering = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt5_userStoryID)["ServiceOffering"].ToString());
                                        string cndt5_SubGroup = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt5_userStoryID)["Subgroup"].ToString());

                                        dt_BugRemediationPostSelfAssessmentStarted.Rows.Add(cndt5_userStoryID, cndt5_SubGroup, cndt5_ServiceOffering);

                                        Console.WriteLine("Condition 5 identified...!");
                                    }
                                    //Bug Remediation Post Self Assessment Completed
                                    if (taskTitle.Contains("bugs remediation post self-assessment") && taskFieldValue_TaskState.ToString() == "Closed")
                                    {
                                        string cndt6_userStoryID = wiID.ToString();
                                        string cndt6_ServiceOffering = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt6_userStoryID)["ServiceOffering"].ToString());
                                        string cndt6_SubGroup = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt6_userStoryID)["Subgroup"].ToString());

                                        dt_BugRemediationPostSelfAssessmentCompleted.Rows.Add(cndt6_userStoryID, cndt6_SubGroup, cndt6_ServiceOffering);

                                        Console.WriteLine("Condition 6 identified...!");
                                    }
                                    //Ready for Assessment Service Onboarding
                                    //[Eng][Activity: Create Grade Review Onboarding Request] --> Should be Closed
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
                                                if (regexItem.IsMatch(childTaskID_Condition7))
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

                                    //Conditon:8 - Assessment Service Onboarding In Progress
                                    /*If Task (Eng Onboarding) is in Closed state & If Task (Assessment Service Onboarding is in Active State) */
                                    if(taskTitle.Contains("create grade review onboarding request") && taskFieldValue_TaskState.ToString() == "Closed")
                                    {
                                        int parentIdForThisTask = wiID;
                                        var wi_Condition8 = GetWorkItemWithRelations(parentIdForThisTask);
                                        if (wi_Condition8.Relations != null)
                                        {
                                            Console.WriteLine("Verifying subcondition If Task (Assessment Service Onboarding is in NewState) in Conditon 8");
                                            foreach (var wiLink_Condition8 in wi_Condition8.Relations)
                                            {
                                                string childTaskURL_Condition8 = wiLink_Condition8.Url;
                                                //Remove the Task Before string
                                                string childTaskID_Condition8 = childTaskURL_Condition8.Replace("https://microsoftit.visualstudio.com/_apis/wit/workItems/", "");
                                                var regexItem_Condition8 = new Regex("^[0-9 ]*$");
                                                if (regexItem.IsMatch(childTaskID_Condition8))
                                                {
                                                    //TaskID
                                                    int wid_childTaskIDCondition8 = int.Parse(childTaskID_Condition8);
                                                    //Add the Task Childs 
                                                    Dictionary<string, string> parentList_Condition8 = new Dictionary<string, string>();
                                                    foreach (var newparent_Condition8 in wiLink_Condition8.Attributes)
                                                    {
                                                        parentList_Condition8.Add(newparent_Condition8.Key, newparent_Condition8.Value.ToString());
                                                    }
                                                    if (parentList_Condition8.ContainsValue("Child"))
                                                    {
                                                        //Get the Task Fields
                                                        var wi_taskWorkFields_Condition8 = GetWorkItem(wid_childTaskIDCondition8);
                                                        if (wi_taskWorkFields_Condition8.Fields["System.WorkItemType"].ToString() == "Task" && wi_taskWorkFields_Condition8.Fields["System.State"].ToString() == "Active")
                                                        {
                                                            var taskTitleVar_Condition8 = CheckFieldAndGetFieldValue(wi_taskWorkFields_Condition8, "System.Title");
                                                            string taskTitle_Condition8 = taskTitleVar_Condition8.ToString().ToLower();
                                                            //For Conditon 8 in Sub scenario title: [Activity:Onboarding]---> Active
                                                            if (taskTitle_Condition8.Contains("activity:onboarding"))
                                                            {
                                                                string cndt8_userStoryID = wiID.ToString();
                                                                string cndt8_ServiceOffering = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt8_userStoryID)["ServiceOffering"].ToString());
                                                                string cndt8_SubGroup = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt8_userStoryID)["Subgroup"].ToString());
                                                                dt_AssessmentServiceOnboardingInProgress.Rows.Add(cndt8_userStoryID, cndt8_SubGroup,cndt8_ServiceOffering);
                                                            }
                                                            else
                                                            {
                                                                Console.WriteLine("Sub scenario not satistied In Conditon 8, that means Title:[Activity:Onboarding] --> Not in Active state..!");
                                                            }
                                                        }
                                                        else
                                                        {
                                                            Console.WriteLine("WorkItemType not in New state for  Conditon 8 in sub scenario");
                                                        }

                                                    }
                                                    else
                                                    {
                                                        Console.WriteLine("This not Child task workItem in subcondition-8..!");
                                                    }

                                                }
                                                else
                                                {
                                                    Console.WriteLine("Currently Identified Not child workItem verifying next one...!");
                                                }
                                            }
                                        }

                                    }

                                    //Condition 9: Scheduled for Grade Review Assessment
                                    /*[Activity:Onboarding] --> Closed State */
                                    if(taskTitle.Contains("[activity:onboarding]") && taskFieldValue_TaskState.ToString() == "Closed")
                                    {
                                        int parentIdForThisTask = wiID;
                                        var wi_Condition9 = GetWorkItemWithRelations(parentIdForThisTask);
                                        if (wi_Condition9.Relations != null)
                                        {
                                            Console.WriteLine("Verifying subcondition If Task (Assessment Service Onboarding is in NewState) in Conditon 9");
                                            foreach (var wiLink_Condition9 in wi_Condition9.Relations)
                                            {
                                                string childTaskURL_Condition9 = wiLink_Condition9.Url;
                                                //Remove the Task Before string
                                                string childTaskID_Condition9 = childTaskURL_Condition9.Replace("https://microsoftit.visualstudio.com/_apis/wit/workItems/", "");
                                                var regexItem_Condition9 = new Regex("^[0-9 ]*$");
                                                if (regexItem_Condition9.IsMatch(childTaskID_Condition9))
                                                {
                                                    //TaskID
                                                    int wid_childTaskIDCondition9 = int.Parse(childTaskID_Condition9);
                                                    //Add the Task Childs 
                                                    Dictionary<string, string> parentList_Condition9 = new Dictionary<string, string>();
                                                    foreach (var newparent_Condition9 in wiLink_Condition9.Attributes)
                                                    {
                                                        parentList_Condition9.Add(newparent_Condition9.Key, newparent_Condition9.Value.ToString());
                                                    }
                                                    if (parentList_Condition9.ContainsValue("Child"))
                                                    {
                                                        //Get the Task Fields
                                                        var wi_taskWorkFields_Condition9 = GetWorkItem(wid_childTaskIDCondition9);
                                                        if (wi_taskWorkFields_Condition9.Fields["System.WorkItemType"].ToString() == "Task" && wi_taskWorkFields_Condition9.Fields["System.State"].ToString() == "New")
                                                        {
                                                            var taskTitleVar_Condition9 = CheckFieldAndGetFieldValue(wi_taskWorkFields_Condition9, "System.Title");
                                                            string taskTitle_Condition9 = taskTitleVar_Condition9.ToString().ToLower();
                                                            //For Conditon 9 in Sub scenario title: [Activity:Assessment]---> New
                                                            if (taskTitle_Condition9.Contains("[activity:assessment]"))
                                                            {
                                                                string cndt9_userStoryID = wiID.ToString();
                                                                string cndt9_ServiceOffering = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt9_userStoryID)["ServiceOffering"].ToString());
                                                                string cndt9_SubGroup = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt9_userStoryID)["Subgroup"].ToString());
                                                                dt_ScheduledForGradeReviewAssessment.Rows.Add(cndt9_userStoryID, cndt9_SubGroup, cndt9_ServiceOffering);
                                                            }
                                                            else
                                                            {
                                                                Console.WriteLine("Sub scenario not satistied In Conditon 9, that means Title:[Activity:Assessment] --> Not in New state..!");
                                                            }
                                                        }
                                                        else
                                                        {
                                                            Console.WriteLine("WorkItemType not in New state for  Conditon 9 in sub scenario");
                                                        }

                                                    }
                                                    else
                                                    {
                                                        Console.WriteLine("This not Child task workItem in subcondition-9..!");
                                                    }

                                                }
                                                else
                                                {
                                                    Console.WriteLine("Currently Identified Not child workItem verifying next one...!");
                                                }
                                            }
                                        }

                                    }

                                    //Grade Review Assessment in Progress
                                    if (taskTitle.Contains("activity:assessment") && taskFieldValue_TaskState.ToString() == "Active")
                                    {
                                        string cndt10_userStoryID = wiID.ToString();
                                        string cndt10_ServiceOffering = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt10_userStoryID)["ServiceOffering"].ToString());
                                        string cndt10_SubGroup = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt10_userStoryID)["Subgroup"].ToString());

                                        dt_GradeReviewAssessmentInProgress.Rows.Add(cndt10_userStoryID, cndt10_SubGroup,cndt10_ServiceOffering);

                                        Console.WriteLine("Condition 10 identified...!");
                                    }
                                    //Grade Review Assessment Completed
                                    if (taskTitle.Contains("activity:assessment") && taskFieldValue_TaskState.ToString() == "Closed")
                                    {
                                        string cndt11_userStoryID = wiID.ToString();
                                        string cndt11_ServiceOffering = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt11_userStoryID)["ServiceOffering"].ToString());
                                        string cndt11_SubGroup = (allP3UserStories.AsEnumerable().FirstOrDefault(p => p["WorkItemID"].ToString() == cndt11_userStoryID)["Subgroup"].ToString());

                                        dt_GradeReviewAssessmentCompleted.Rows.Add(cndt11_userStoryID, cndt11_SubGroup, cndt11_ServiceOffering);

                                        Console.WriteLine("Condition11 identified...!");
                                    }
                                    //}
                                }
                            }
                        } 
                    }
                }
            }

            List<DataTable> dataTableList = new List<DataTable>();
            dataTableList.Add(dt_FY20GradeC);
            dataTableList.Add(dt_GradeReviewAssessmentCompleted);
            dataTableList.Add(dt_GradeReviewAssessmentInProgress);
            dataTableList.Add(dt_ScheduledForGradeReviewAssessment);
            dataTableList.Add(dt_AssessmentServiceOnboardingInProgress);
            dataTableList.Add(dt_ReadyForAssessmentServiceOnboarding);
            dataTableList.Add(dt_BugRemediationPostSelfAssessmentCompleted);
            dataTableList.Add(dt_BugRemediationPostSelfAssessmentStarted);
            dataTableList.Add(dt_BugRemediationPostSelfAssessmentNotStarted);
            dataTableList.Add(dt_SelfAssessmentCompleted);
            dataTableList.Add(dt_SelfAssessmentInProgress);
            dataTableList.Add(dt_SelfAssessmentToStart);
            dataTableList.Add(allP3UserStories);

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
            dbConn = @"Data Source = airtproddbserver.database.windows.net; user id=AIRTReader; password=Reader_AIRT@12; Initial Catalog = AIRTProd;";
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
            builder.Append("<table>");
            builder.Append("<tr class=\"rowheader rh1\">");
            builder.Append("<th colspan=\"15\" scope=\"colgroup\">CSEO FY20 P3 Accessibility Progress Scorecard</th>");
            builder.Append("</tr>");
            builder.Append("<tr class=\"rowheader\">");
            builder.Append("<th><b>CSEO Organization</b></th>");
            builder.Append("<th><b>Service Group</b></th>");
            builder.Append("<th><b>Self Assessment To Start</b></th>");
            builder.Append("<th><b>Self-Assessment In Progress</b></th>");
            builder.Append("<th><b>Self-Assessment Completed</b></th>");
            builder.Append("<th><b>Bug Remediation Post Self Assessment Not Started</b></th>");
            builder.Append("<th><b>Bug Remediation Post Self Assessment Started</b></th>");
            builder.Append("<th><b>Bug Remediation Post Self Assessment Completed</b></th>");
            builder.Append("<th><b>Ready for Assessment Service Onboarding</b></th>");
            builder.Append("<th><b>Assessment Service Onboarding In Progress</b></th>");
            builder.Append("<th><b>Scheduled for Grade Review Assessment</b></th>");
            builder.Append("<th><b>Grade Review Assessment in Progress</b></th>");
            builder.Append("<th><b>Grade Review Assessment Completed</b></th>");
            builder.Append("<th><b>P3 Applications Grade C Applications Count</b></th>");
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
                            builder.Append("<td class=\"text-td\" bgcolor=\"#7B74F3\">" + condition_1_UserStoryCount + "</td>");
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
                            builder.Append("<td class=\"text-td\" bgcolor=\"#756BCA\">" + condition_2_UserStoryCount + "</td>");
                        }
                        

                        //Condition-3: Self-Assessment Completed [Table:dt_SelfAssessmentCompleted
                        var condition3Data = (from rSubCond3 in dt_SelfAssessmentCompleted.AsEnumerable()
                                              where rSubCond3.Field<string>("SubGroup") == sugroupStrFmt
                                              && (rSubCond3.Field<string>("ServiceLine") == serviceOfferingName)
                                              select rSubCond3.Field<string>("WorkItemID"));
                        condition_3_UserStoryCount = condition3Data.Distinct().Count();
                        if(condition_3_UserStoryCount==0)
                        {
                            builder.Append("<td class=\"text-td\">" + condition_3_UserStoryCount + "</td>");
                        }
                        else
                        {
                            builder.Append("<td class=\"text-td\" bgcolor=\"#8EE07D\">" + condition_3_UserStoryCount + "</td>");
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
                            builder.Append("<td class=\"text-td\" bgcolor=\"#EEABAB\">" + testCOunt + "</td>");
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
                            builder.Append("<td class=\"text-td\" bgcolor=\"#7B74F3\">" + condition_5_UserStoryCount + "</td>");
                        }
                       

                        //Condition-6: Bug Remediation Post Self Assessment Completed [Table: dt_BugRemediationPostSelfAssessmentCompleted]
                        var condition6Data = (from rSubCond6 in dt_BugRemediationPostSelfAssessmentCompleted.AsEnumerable()
                                              where rSubCond6.Field<string>("SubGroup") == sugroupStrFmt
                                              && (rSubCond6.Field<string>("ServiceLine") == serviceOfferingName)
                                              select rSubCond6.Field<string>("WorkItemID"));
                        condition_6_UserStoryCount = condition6Data.Distinct().Count();
                        
                        if(condition_6_UserStoryCount==0)
                        {
                            builder.Append("<td class=\"text-td\">" + condition_6_UserStoryCount + "</td>");
                        } else
                        {
                            builder.Append("<td class=\"text-td\" bgcolor=\"#8EE07D\">" + condition_6_UserStoryCount + "</td>");
                        }

                        //Condition-7: Ready for Assessment Service Onboarding [Table: dt_ReadyForAssessmentServiceOnboarding]
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
                            builder.Append("<td class=\"text-td\" bgcolor=\"#7B74F3\">" + condition_7_UserStoryCount + "</td>");
                        }

                        //Condition-8: Assessment Service Onboarding In Progress [Table: dt_AssessmentServiceOnboardingInProgress]
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
                            builder.Append("<td class=\"text-td\" bgcolor=\"#756BCA\">" + condition_8_UserStoryCount + "</td>");
                        }

                        //Condition-9: Scheduled for Grade Review Assessment [Table: dt_ScheduledForGradeReviewAssessment]
                        var condition9Data = (from rSubCond9 in dt_ScheduledForGradeReviewAssessment.AsEnumerable()
                                              where rSubCond9.Field<string>("SubGroup") == sugroupStrFmt
                                              && (rSubCond9.Field<string>("ServiceLine") == serviceOfferingName)
                                              select rSubCond9.Field<string>("WorkItemID"));
                        condition_9_UserStoryCount = condition9Data.Distinct().Count();
                        
                        if(condition_9_UserStoryCount==0)
                        {
                            builder.Append("<td class=\"text-td\">" + condition_9_UserStoryCount + "</td>");
                        }
                        else
                        {
                            builder.Append("<td class=\"text-td\" bgcolor=\"#7B74F3\">" + condition_9_UserStoryCount + "</td>");
                        }

                        //Condition-10: Grade Review Assessment in Progress [Table:dt_GradeReviewAssessmentInProgress]
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
                            builder.Append("<td class=\"text-td\" bgcolor=\"#756BCA\">" + condition_10_UserStoryCount + "</td>");
                        }

                        //Condition-11: Grade Review Assessment Completed [Table: dt_GradeReviewAssessmentCompleted]
                        var condition11Data = (from rSubCond11 in dt_GradeReviewAssessmentCompleted.AsEnumerable()
                                               where rSubCond11.Field<string>("SubGroup") == sugroupStrFmt
                                               && (rSubCond11.Field<string>("ServiceLine") == serviceOfferingName)
                                               select rSubCond11.Field<string>("WorkItemID"));
                        condition_11_UserStoryCount = condition11Data.Distinct().Count();
                        
                        if(condition_11_UserStoryCount==0)
                        {
                            builder.Append("<td class=\"text-td\">" + condition_11_UserStoryCount + "</td>");
                        }
                        else
                        {
                            builder.Append("<td class=\"text-td\" bgcolor=\"#8EE07D\">" + condition_11_UserStoryCount + "</td>");
                        }

                        //Condition-12: P3 Applications Grade C Applications Count [Table: dt_FY20GradeC]
                        var condition12Data = (from rSubCond12 in dt_FY20GradeC.AsEnumerable()
                                               where rSubCond12.Field<string>("SubGroup") == sugroupStrFmt
                                               && (rSubCond12.Field<string>("ServiceLine") == serviceOfferingName)
                                               select rSubCond12.Field<string>("WorkItemID"));
                        condition_12_UserStoryCount = condition12Data.Distinct().Count();
                        
                        if(condition_12_UserStoryCount==0)
                        {
                            builder.Append("<td class=\"text-td\">" + condition_12_UserStoryCount + "</td>");
                        }
                        else
                        {
                            builder.Append("<td class=\"text-td\" bgcolor=\"#8EE07D\">" + condition_12_UserStoryCount + "</td>");
                        }

                        //Condition-13: P3 Applications User Story Count [Table: dt_FY20UserStoryCount]
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
            builder.Append("<th>" + dt_SelfAssessmentToStart.Rows.Count + " / "+ (int)Math.Round((double)(100 * dt_SelfAssessmentToStart.Rows.Count) / allP3UserStories.Rows.Count) + "</th>");
            builder.Append("<th>" + dt_SelfAssessmentInProgress.Rows.Count + " / " + (int)Math.Round((double)(100 * dt_SelfAssessmentInProgress.Rows.Count) / allP3UserStories.Rows.Count) + "</th>");
            builder.Append("<th>" + dt_SelfAssessmentCompleted.Rows.Count + " / " + (int)Math.Round((double)(100 * dt_SelfAssessmentCompleted.Rows.Count) / allP3UserStories.Rows.Count) + "</th>");
            builder.Append("<th>" + dt_BugRemediationPostSelfAssessmentNotStarted.Rows.Count + " / " + (int)Math.Round((double)(100 * dt_BugRemediationPostSelfAssessmentNotStarted.Rows.Count) / allP3UserStories.Rows.Count) + "</th>");
            builder.Append("<th>" + dt_BugRemediationPostSelfAssessmentStarted.Rows.Count + " / " + (int)Math.Round((double)(100 * dt_BugRemediationPostSelfAssessmentStarted.Rows.Count) / allP3UserStories.Rows.Count) + "</th>");
            builder.Append("<th>" + dt_BugRemediationPostSelfAssessmentCompleted.Rows.Count + " / " + (int)Math.Round((double)(100 * dt_BugRemediationPostSelfAssessmentCompleted.Rows.Count) / allP3UserStories.Rows.Count) + "</th>");
            builder.Append("<th>" + dt_ReadyForAssessmentServiceOnboarding.Rows.Count + " / " + (int)Math.Round((double)(100 * dt_ReadyForAssessmentServiceOnboarding.Rows.Count) / allP3UserStories.Rows.Count) + "</th>");
            builder.Append("<th>" + dt_AssessmentServiceOnboardingInProgress.Rows.Count + " / " + (int)Math.Round((double)(100 * dt_AssessmentServiceOnboardingInProgress.Rows.Count) / allP3UserStories.Rows.Count) + "</th>");
            builder.Append("<th>" + dt_ScheduledForGradeReviewAssessment.Rows.Count + " / " + (int)Math.Round((double)(100 * dt_ScheduledForGradeReviewAssessment.Rows.Count) / allP3UserStories.Rows.Count) + "</th>");
            builder.Append("<th>" + dt_GradeReviewAssessmentInProgress.Rows.Count + " / " + (int)Math.Round((double)(100 * dt_GradeReviewAssessmentInProgress.Rows.Count) / allP3UserStories.Rows.Count) + "</th>");
            builder.Append("<th>" + dt_GradeReviewAssessmentCompleted.Rows.Count + " / " + (int)Math.Round((double)(100 * dt_GradeReviewAssessmentCompleted.Rows.Count) / allP3UserStories.Rows.Count) + "</th>");
            builder.Append("<th>" + dt_FY20GradeC.Rows.Count + " / " + (int)Math.Round((double)(100 * dt_FY20GradeC.Rows.Count) / allP3UserStories.Rows.Count) + "</th>");
            builder.Append("<th>" + allP3UserStories.Rows.Count + "</th>");
            builder.Append("</tr>");
            builder.Append("</table>");
            builder.Append("</body>");
            builder.Append("</html>");
            string HtmlFile = builder.ToString();
            oMsg.Attachments.Add(Path.GetFullPath("D:\\Reports Excel\\AppP3UserStories\\P3UserStories.xlsx"));
            //Attachment emailAttachment = oMsg.Attachments.Add("D:\\Reports Excel\\AppP3UserStories\\P3UserStories.xlsx", Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue);
            
            //QA Test
            //oMsg.To = ConfigurationManager.AppSettings["Prod_To_Test_P3ApplicationSend"];

            //QA Draft
            oMsg.To = ConfigurationManager.AppSettings["Prod_To_Draft_P3ApplicationSend"];

            //Prod
            //oMsg.To = ConfigurationManager.AppSettings["Prod_To_P3ApplicationSend"];
            //oMsg.CC = ConfigurationManager.AppSettings["Prod_CC_P3ApplicationSend"];

            DateTime startAtMonday = DateTime.Now.AddDays(DayOfWeek.Monday - DateTime.Now.DayOfWeek);
            oMsg.Subject = "Draft: CSEO Assessment Services : CSEO FY20 P3 Accessibility Progress Scorecard - " + DateTime.Now.ToString("MM/dd/yyyy");
            oMsg.HTMLBody = HtmlFile;
            Console.WriteLine("ADO Features are created and sent mail to respective folks");
            oMsg.Send();
        }
    }
}
