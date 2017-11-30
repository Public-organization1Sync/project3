// <copyright file="BranchManagement.cs" company="Syncfusion Software">
//     Copyright (c) Id and Name. All rights reserved.
// </copyright>

[module: System.Diagnostics.CodeAnalysis.SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "Reviewed. Suppression is OK here.")]
[module: System.Diagnostics.CodeAnalysis.SuppressMessage("StyleCop.CSharp.NamingRules", "SA1305:FieldNamesMustNotUseHungarianNotation", Justification = "Reviewed. Suppression is OK here.")]
[module: System.Diagnostics.CodeAnalysis.SuppressMessage("StyleCop.CSharp.NamingRules", "SA1300:ElementMustBeginWithUpperCaseLetter", Justification = "Reviewed. Suppression is OK here.")]

/// <summary>
/// GitLab branch management DLL
/// </summary>
namespace Syncfusion.Gitlab
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Collections.Specialized;
    using System.Diagnostics;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Net.Mail;
    using System.Reflection;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading;
    using Amazon.SimpleEmail;
    using Amazon.SimpleEmail.Model;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;
    using Syncfusion.XlsIO;

    /// <summary>
    /// branch management DLL
    /// </summary>
    public class BranchManagement
    {
        /// <summary>
        /// GitLab Access Token
        /// </summary>
        public const string Token = "GMeH1efM9e8umMN4wxBL";

        /// <summary>
        /// object of the ExcelOperation class
        /// </summary>
        private ExcelOperation excel = new ExcelOperation();

        /// <summary>
        /// Object for mail notification
        /// </summary>
        private MailNotification mail = new MailNotification();

        /// <summary>
        /// Enum data for groups
        /// </summary>
        public enum Group
        {
            /// <summary>
            /// Consulting group`s ID 
            /// </summary>
            Consulting = 459,

            /// <summary>
            /// Content group`s ID
            /// </summary>
            Content = 505,

            /// <summary>
            /// DataScience group`s ID
            /// </summary>
            DataScience = 452,

            /// <summary>
            /// EssentialStudio group`s ID
            /// </summary>
            EssentialStudio = 450,

            /// <summary>
            /// General group`s ID
            /// </summary>
            General = 486,

            /// <summary>
            /// HRPortal group`s ID
            /// </summary>
            HRPortal = 458,

            /// <summary>
            /// Infrastructure group`s ID
            /// </summary>
            Infrastructure = 456,

            /// <summary>
            /// InnovationLab group`s ID
            /// </summary>
            InnovationLab = 581,

            /// <summary>
            /// MetroStudio group`s ID
            /// </summary>
            MetroStudio = 457,

            /// <summary>
            /// Syncfusion group`s ID
            /// </summary>
            Syncfusion = 460
        }

        /// <summary>
        /// creating the branch 
        /// </summary>
        /// <param name="groupName">Name of the group</param>
        /// <param name="projectName">Project Name List</param>
        /// <param name="sourceBranch">Source branch name</param>
        /// <param name="destinationBranch">Destination branch name</param>
        public void CreateBranch(Group groupName, List<string> projectName, string sourceBranch, string destinationBranch)
        {
            if (projectName == null || projectName.Count == 0)
            {
                throw new ArgumentNullException(nameof(projectName));
            }

            if (string.IsNullOrEmpty(sourceBranch) == true)
            {
                throw new ArgumentNullException(nameof(sourceBranch));
            }

            if (string.IsNullOrEmpty(destinationBranch) == true)
            {
                throw new ArgumentNullException(nameof(destinationBranch));
            }

            this.BranchCreationExcelDocument(groupName, projectName, sourceBranch, destinationBranch);
        }

        /// <summary>
        /// Creating the single branch
        /// </summary>
        /// <param name="groupName">Name of the group</param>
        /// <param name="projectName">Project Name</param>
        /// <param name="sourceBranch">Source branch name</param>
        /// <param name="destinationBranch">Destination branch name</param>
        /// <returns>return the status of the branch creation</returns>
        public string CreateBranch(Group groupName, string projectName, string sourceBranch, string destinationBranch)
        {
            List<ProjectDetails> projectDetails = this.GetProjectList((int)groupName);

            if (string.IsNullOrEmpty(projectName) == true)
            {
                throw new ArgumentNullException(nameof(projectName));
            }

            if (string.IsNullOrEmpty(sourceBranch) == true)
            {
                throw new ArgumentNullException(nameof(sourceBranch));
            }

            if (string.IsNullOrEmpty(destinationBranch) == true)
            {
                throw new ArgumentNullException(nameof(destinationBranch));
            }

            try
            {
                string jsonstringForCreateBranch;
                string status;

                if (destinationBranch.Contains("/"))
                {
                    var item = projectDetails.FirstOrDefault(x => x.Name == projectName.Trim());
                    if (item != null)
                    {
                        List<string> branchList = this.GetBranchList(item.Id).Select(branch => branch.Name).ToList();
                        List<string> tagsList = this.GetTagsList(item.Id);
                        var concatList = branchList.Concat(tagsList);
                        List<string> branchTagList = concatList.ToList();
                        if (branchTagList.Contains(sourceBranch.Trim()))
                        {
                            jsonstringForCreateBranch = this.BranchCreation(item.Id, destinationBranch, sourceBranch);
                            JObject jsonForCreateBranch = JObject.Parse(jsonstringForCreateBranch);
                            string branchName = (string)jsonForCreateBranch["name"];
                            if (branchName != null)
                            {
                                this.BranchProtection(item.Id, destinationBranch);
                                status = "Created";
                            }
                            else
                            {
                                if ((string)jsonForCreateBranch["message"] == "Branch already exists")
                                {
                                    this.BranchProtection(item.Id, destinationBranch);
                                    status = "Already Exists";
                                }
                                else
                                {
                                    status = (string)jsonForCreateBranch["message"];
                                }
                            }
                        }
                        else
                        {
                            status = "Given source branch not exist";
                        }
                    }
                    else
                    {
                        status = "Project not exist in the GitLab instances. Check the spelling of the project";
                    }
                }
                else
                {
                    status = "Destination branch must contain branch type \"/\" before vx.x.x (ex: hotfix/v1.0.0)";
                }

                return status;
            }
            catch (WebException ex)
            {
                string status;
                status = "Exception in branch creation: " + ex;
                return status;
            }
        }

        /// <summary>
        /// Creating the branch and send the mail log excel sheet to given mail ID
        /// </summary>
        /// <param name="groupName">Name of the group</param>
        /// <param name="projectName">Project Name List</param>
        /// <param name="sourceBranch">Source branch name</param>
        /// <param name="destinationBranch">Destination branch name</param>
        /// <param name="emailId">Mail ID to send the mail</param>
        /// /// <returns>return the status of the branch creation</returns>
        public List<OperationResult> CreateBranch(Group groupName, List<string> projectName, string sourceBranch, string destinationBranch, string emailId)
        {
            List<OperationResult> branchCreationResult = new List<OperationResult>();

            if (projectName == null || projectName.Count == 0)
            {
                throw new ArgumentNullException(nameof(projectName));
            }

            if (string.IsNullOrEmpty(sourceBranch) == true)
            {
                throw new ArgumentNullException(nameof(sourceBranch));
            }

            if (string.IsNullOrEmpty(destinationBranch) == true)
            {
                throw new ArgumentNullException(nameof(destinationBranch));
            }

            if (string.IsNullOrEmpty(emailId) == true)
            {
                throw new ArgumentNullException(nameof(emailId));
            }

            Collection<ProjectData> branchCreatedData = this.BranchCreationExcelDocument(groupName, projectName, sourceBranch, destinationBranch);
            for (int projectCount = 0; projectCount < branchCreatedData.Count; projectCount++)
            {
                if (branchCreatedData[projectCount].Status.Equals("Branch successfully created"))
                {
                    branchCreationResult.Add(new OperationResult { ProjectName = branchCreatedData[projectCount].Project_Name, Comments = "Branch successfully created", Status = true });
                }
                else if (branchCreatedData[projectCount].Status.Equals("Branch already exists"))
                {
                    branchCreationResult.Add(new OperationResult { ProjectName = branchCreatedData[projectCount].Project_Name, Comments = "Branch already exists", Status = true });
                }
                else
                {
                    branchCreationResult.Add(new OperationResult { ProjectName = branchCreatedData[projectCount].Project_Name, Comments = "Branch is not deleted", Status = false });
                }
            }

            string mailContent = "<html><head><style type=\"text / css\">body {font-family: Segoe UI;font-size: 14px;}table tr {border: none;}a, a:link, a:visited, a:active {text-decoration: none;color: #0066ff;}div.content a{color: #29ABE2;}img {border: none;}</style></head><body><table cellspacing=\"1\"  cellpadding=\"0\" align=\"center\" id=\"main - panel\" style=\"width: 100 %; background - color:#f2f2f2; color: rgb(75, 75, 75); border-collapse: collapse; border: 1px solid #f2f2f2;\"><tbody><tr><td ><table cellspacing=\"1\"  cellpadding=\"0\" align=\"center\" id=\"main-panel\" style=\"width: 740px; background-color:#ffffff; color: rgb(75, 75, 75); border-collapse: collapse; border: 1px solid #C6C6C6;\"><tbody><tr style=\"width: 100%;\"><td style=\"padding-top:25px;\"><table cellspacing=\"1\"  cellpadding=\"0\" align=\"center\" id=\"main-panel\" style=\"width: 700px;background-color:#ffffff;color: rgb(75, 75, 75); border-collapse: collapse;\"><tbody><tr style=\"width: 100%;\"><td style=\"padding:0px 0px;\"><table style=\" width: 100%; padding: 0px 0px;\"><tbody><tr><p>Hi Everyone</p><p>Branch creation operation is completed successfully. Kindly follow the excel document for more information.</p></tr></tbody></table></td></tr><tr><td style=\" color:#4D4D4D;padding: 20px 0px 40px 0px ;border-bottom: 1px solid #C6C6C6; font-size: 11px;\"><div>Best regards, </div><div><b>GitLab Team </b></div></td></tr><tr><td style=\"padding: 10px 0 ;\"><table><tbody><tr><td style=\"color: #999999; font-size: 10px;\"><p>This is an automatically generated message; please do not reply to this email.</p></td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table></td></tr><tr width=\"100%\"><td style=\"background: #f2f2f2; padding: 10px 0px 10px 0px;\"><table width=\"740px\" style=\"width:740px;\"align=\"center\"><tbody><tr><td  width=\"490px\" style=\" width:490px;color: #999999; font-size: 9px;\"><p><a style=\"text-decoration: none;color: #999999;\" target=\"_top\" href=\"http://syncfusion.com/copyright\">Copyright © 2001 - 2017 Syncfusion Inc. All Rights Reserved</a></p></td><td align=\"right\" style=\" color: #999999; font-size: 9px;\"><p><a style=\"text-decoration: none;color: #999999;\" target=\"_top\" href=\"http://syncfusion.com/privacy\">Privacy Policy</a> | <a style=\"text-decoration: none;color: #999999;\" target=\"_top\" href=\"http://syncfusion.com/company/contact-us\">Contact Us</a></p></td></tr></tbody></table></td></tr></tbody></table></body></html>";
            this.mail.MailNotifications("Create branch status", mailContent, @"GitLab_Log_file.xlsx", emailId);
            return branchCreationResult;
        }

        /// <summary>
        /// Creating the tag 
        /// </summary>
        /// <param name="groupName">Name of the group</param>
        /// <param name="projectName">Project name list</param>
        /// <param name="sourceBranch">Source branch name</param>
        /// <param name="tagname">tag name</param>
        public void CreateTag(Group groupName, List<string> projectName, string sourceBranch, string tagname)
        {
            Collection<ProjectData> datatoExcel = new Collection<ProjectData>();
            Collection<string> projectUrlList = new Collection<string>();
            Collection<string> tagsUrlList = new Collection<string>();
            List<ProjectDetails> projectDetails = this.GetProjectList((int)groupName);

            if (projectName == null || projectName.Count == 0)
            {
                throw new ArgumentNullException(nameof(projectName));
            }

            if (string.IsNullOrEmpty(sourceBranch) == true)
            {
                throw new ArgumentNullException(nameof(sourceBranch));
            }

            if (string.IsNullOrEmpty(tagname) == true)
            {
                throw new ArgumentNullException(nameof(tagname));
            }

            try
            {
                string output;
                string tagsUrl = string.Empty;
                List<string> distinctProjectName = new List<string>();
                projectName = projectName.Distinct().ToList();
                distinctProjectName = projectName.OrderBy(data => data).ToList();
                for (int projectCount = 0; projectCount < distinctProjectName.Count; projectCount++)
                {
                    var data = projectDetails.FirstOrDefault(x => x.Name == distinctProjectName[projectCount].Trim());
                    if (data != null)
                    {
                        ////Console.WriteLine(data.Name);
                        List<string> branchList = this.GetBranchList(data.Id).Select(branch => branch.Name).ToList();
                        List<string> tagsList = this.GetTagsList(data.Id);
                        var concatList = branchList.Concat(tagsList);
                        List<string> branchTagList = concatList.ToList();
                        if (branchTagList.Contains(sourceBranch.Trim()))
                         {
                            output = this.TagCreation(data.Id, tagname, sourceBranch);
                            JObject json = JObject.Parse(output);
                            tagsUrl = data.Web_Url + "/tags/" + tagname.Trim();                            
                            string tagName = (string)json["name"];
                            if (tagName != null)
                            {
                                datatoExcel.Add(new ProjectData { Project_Name = distinctProjectName[projectCount].Trim(), Source_Branch = sourceBranch.Trim(), Destination_Branch = tagname.Trim(), Status = "Tag successfully created" });
                                projectUrlList.Add(data.Web_Url);
                                tagsUrlList.Add(tagsUrl);
                            }
                            else
                            {
                                if ((string)json["message"] == "Invalid reference name")
                                {
                                    datatoExcel.Add(new ProjectData { Project_Name = distinctProjectName[projectCount].Trim(), Source_Branch = sourceBranch.Trim(), Destination_Branch = "-", Status = (string)json["message"] });
                                    projectUrlList.Add(data.Web_Url);
                                    tagsUrlList.Add("empty");
                                }

                                datatoExcel.Add(new ProjectData { Project_Name = distinctProjectName[projectCount].Trim(), Source_Branch = sourceBranch.Trim(), Destination_Branch = tagname.Trim(), Status = (string)json["message"] });
                                projectUrlList.Add(data.Web_Url);
                                tagsUrlList.Add(tagsUrl);
                            }
                        }
                        else
                        {
                            datatoExcel.Add(new ProjectData { Project_Name = distinctProjectName[projectCount].Trim(), Source_Branch = sourceBranch.Trim(), Destination_Branch = "-", Status = "Given source branch not exist" });
                            projectUrlList.Add(data.Web_Url);
                            tagsUrlList.Add("empty");
                        }
                    }
                    else
                    {
                        datatoExcel.Add(new ProjectData { Project_Name = distinctProjectName[projectCount].Trim(), Source_Branch = sourceBranch.Trim(), Destination_Branch = "-", Status = "Project not exist in the GitLab instances. Check the spelling of the project" });
                        projectUrlList.Add("empty");
                        tagsUrlList.Add("empty");
                    }
                }

                this.excel.GenerateExcelSheet(datatoExcel, projectUrlList, tagsUrlList, "Tag Creation");
            }
            catch (WebException ex)
            {
                datatoExcel.Add(new ProjectData { Project_Name = "-", Source_Branch = "-", Destination_Branch = "-", Status = "Exception in Tag creation. " + ex });
                projectUrlList.Add("empty");
                tagsUrlList.Add("empty");
                this.excel.GenerateExcelSheet(datatoExcel, projectUrlList, tagsUrlList, "Tag Creation");
            }
        }

        /// <summary>
        /// Get the single project`s branch details by passing the project name and project type
        /// </summary>
        /// <param name="groupName">Name of the group</param>
        /// <param name="projectName">Name of the project</param>
        /// <param name="branchType">banch name starts with, if you need all branches just pass empty branch type</param>
        /// <returns> returns the string collection of branch for given branch</returns>
        public StringCollection Branches(Group groupName, string projectName, string branchType)
        {
            if (string.IsNullOrEmpty(projectName) == true)
            {
                throw new ArgumentNullException(nameof(projectName));
            }

            if (branchType == null)
            {
                throw new ArgumentNullException(nameof(branchType));
            }

            StringCollection branchDetails = new StringCollection();
            string branchHint = branchType.ToLower();
            try
            {                 
                List<ProjectDetails> projectList = this.GetProjectList((int)groupName);
                var data = projectList.FirstOrDefault(x => x.Name == projectName.Trim());
                if (data != null)
                {
                    List<string> branchNameList = this.GetBranchList(data.Id).Select(branch => branch.Name).ToList();
                    List<string> tagsNameList = this.GetTagsList(data.Id);
                    List<string> branchesandTagsNameList = branchNameList.Concat(tagsNameList).ToList();
                    if (string.IsNullOrEmpty(branchType) == false)
                    {
                        int branchTypeLength = branchType.Count();
                        foreach (string branchName in branchesandTagsNameList)
                        {
                            if (branchTypeLength <= branchName.Count())
                            {
                                string branchNameStarts = branchName.Substring(0, branchTypeLength).ToLower();
                                if (branchNameStarts == branchHint)
                                {
                                    branchDetails.Add(branchName);
                                }
                            }
                        }

                        if (branchDetails.Count == 0)
                        {
                            branchDetails.Add("No branches starts with " + branchType);
                        }
                    }
                    else
                    {
                        foreach (string branchName in branchesandTagsNameList)
                        {
                            branchDetails.Add(branchName);
                        }
                    }
                }
                else
                {
                    branchDetails.Add("Given project is not exist in our GitLab environment, Check the project`s spelling.");
                }

                return branchDetails;
            }
            catch (IOException ex)
            {
                branchDetails.Add(ex.ToString());
                return branchDetails;
            }
        }

        /// <summary>
        /// Get the common branch list
        /// </summary>
        /// <param name="groupName">Name of the group</param>
        /// <param name="projectName">List of GitLab project List</param>
        /// <returns>returns the common branch list</returns>
        public List<string> Branches(Group groupName, List<string> projectName)
        {
            int projectChecking = 0;
            List<string> commonBranches = new List<string>();
            string branches = string.Empty;
            if (projectName == null || projectName.Count == 0)
            {
                throw new ArgumentNullException(nameof(projectName));
            }

            try
            {
                List<ProjectDetails> projectList = this.GetProjectList((int)groupName);

                int count;
                for (count = 0; count < projectName.Count; count++)
                {
                    var data = projectList.FirstOrDefault(x => x.Name == projectName[count].Trim());
                    if (data != null)
                    {
                        projectChecking = projectChecking + 1;
                    }
                }

                if (projectChecking == projectName.Count)
                {
                    for (count = 0; count < projectName.Count; count++)
                    {
                        var data = projectList.FirstOrDefault(x => x.Name == projectName[count].Trim());
                        List<string> projectBranchesandTags = this.GetBranchList(data.Id).Select(branch => branch.Name).ToList();
                        foreach (string branchdata in projectBranchesandTags)
                        {
                            branches = branches + "%%" + branchdata;
                        }

                        projectBranchesandTags = this.GetTagsList(data.Id);
                        foreach (string branchdata in projectBranchesandTags)
                        {
                            branches = branches + "%%" + branchdata;
                        }
                    }

                    branches = branches + "%%";

                    BranchChecker:
                    if (branches != "%%")
                    {
                        string branchNameCheck = "%" + branches.Split('%')[2].Trim() + "%";
                        branchNameCheck = branchNameCheck.Replace("(", "\\(").Replace(")", "\\)");
                        int repeatedCount = Regex.Matches(branches, branchNameCheck).Count;

                        if (repeatedCount == projectName.Count)
                        {
                            commonBranches.Add(branchNameCheck.Trim('%'));
                        }

                        branchNameCheck = branchNameCheck.Replace("\\(", "(").Replace("\\)", ")");
                        branches = branches.Replace(branchNameCheck, "%");
                        branches = Regex.Replace(branches, "[%]+", "%%");
                        goto BranchChecker;
                    }
                }
                else
                {
                    commonBranches.Add("Some Given Projects are not exists in GitLab");
                }

                return commonBranches;
            }
            catch (IOException ex)
            {
                commonBranches.Add(ex.ToString());
                return commonBranches;
            }
        }

        /// <summary>
        /// Deleting the tags for list of projects
        /// </summary>
        /// <param name="groupName">Name of the group</param>
        /// <param name="projectName">List of project</param>
        /// <param name="tagsName">Tags name which to be deleted</param>
        public void DeleteTag(Group groupName, List<string> projectName, string tagsName)
        {
            List<ProjectDetails> projectDetails = this.GetProjectList((int)groupName);
            Collection<ProjectData> datatoExcel = new Collection<ProjectData>();
            Collection<string> tagsUrlList = new Collection<string>();
            Collection<string> projectUrlList = new Collection<string>();

            if (projectName == null || projectName.Count == 0)
            {
                throw new ArgumentNullException(nameof(projectName));
            }

            if (string.IsNullOrEmpty(tagsName) == true)
            {
                throw new ArgumentNullException(nameof(tagsName));
            }

            try
            {
                string output;
                List<string> distinctProjectName = new List<string>();
                projectName = projectName.Distinct().ToList();
                distinctProjectName = projectName.OrderBy(data => data).ToList();
                for (int projectCount = 0; projectCount < distinctProjectName.Count; projectCount++)
                {
                    var data = projectDetails.FirstOrDefault(x => x.Name == distinctProjectName[projectCount].Trim());
                    if (data != null)
                    {
                        List<string> tagsList = this.GetTagsList(data.Id);
                        if (tagsList.Contains(tagsName.Trim()))
                        {
                            output = this.TagDeletion(data.Id, tagsName);
                            if (string.IsNullOrEmpty(output))
                            {
                                datatoExcel.Add(new ProjectData { Project_Name = distinctProjectName[projectCount].Trim(), Source_Branch = tagsName.Trim(), Status = "Tag successfully deleted" });
                                projectUrlList.Add(data.Web_Url);
                            }
                            else
                            {
                                JObject json = JObject.Parse(output);
                                datatoExcel.Add(new ProjectData { Project_Name = distinctProjectName[projectCount].Trim(), Source_Branch = tagsName.Trim(), Status = (string)json["message"] });
                                projectUrlList.Add(data.Web_Url);
                            }
                        }
                        else
                        {
                            datatoExcel.Add(new ProjectData { Project_Name = distinctProjectName[projectCount].Trim(), Source_Branch = tagsName.Trim(), Status = "Given tag does not exists" });
                            projectUrlList.Add(data.Web_Url);
                        }
                    }
                    else
                    {
                        datatoExcel.Add(new ProjectData { Project_Name = distinctProjectName[projectCount].Trim(), Source_Branch = tagsName.Trim(), Status = "Project not exist in the GitLab instances. Check the spelling of the project" });
                        projectUrlList.Add("empty");
                    }
                }

                this.excel.GenerateExcelSheet(datatoExcel, projectUrlList, tagsUrlList, "Tag Deletion");
            }
            catch (WebException ex)
            {
                datatoExcel.Add(new ProjectData { Project_Name = "-", Source_Branch = "-", Destination_Branch = "-", Status = "Exception in Tag deletion. " + ex });
                projectUrlList.Add("empty");
                this.excel.GenerateExcelSheet(datatoExcel, projectUrlList, tagsUrlList, "Tag Deletion");
            }
        }

        /// <summary>
        /// Deleting the tags for single project
        /// </summary>
        /// <param name="groupName">Name of the group</param>
        /// <param name="projectName">project name</param>
        /// <param name="tagsName">tag name which is to be deleted</param>
        /// <returns> returns the status of the tag deletion operation</returns>
        public string DeleteTag(Group groupName, string projectName, string tagsName)
        {
            List<ProjectDetails> projectDetails = this.GetProjectList((int)groupName);
            string status;

            if (string.IsNullOrEmpty(projectName) == true)
            {
                throw new ArgumentNullException(nameof(projectName));
            }

            if (string.IsNullOrEmpty(tagsName) == true)
            {
                throw new ArgumentNullException(nameof(tagsName));
            }

            try
            {
                string output;
                var data = projectDetails.FirstOrDefault(x => x.Name == projectName.Trim());
                if (data != null)
                {
                    List<string> tagsList = this.GetTagsList(data.Id);
                    if (tagsList.Contains(tagsName.Trim()))
                    {
                        output = this.TagDeletion(data.Id, tagsName);
                        if (string.IsNullOrEmpty(output))
                        {
                            status = "Tag successfully deleted";
                        }
                        else
                        {
                            JObject json = JObject.Parse(output);
                            status = (string)json["message"];
                        }
                    }
                    else
                    {
                        status = "Given tag does not exists";
                    }
                }
                else
                {
                    status = "Project not exist in the GitLab instances. Check the spelling of the project";
                }
            }
            catch (WebException ex)
            {
                status = "Exception in Tag deletion. " + ex;
            }

            return status;
        }

        /// <summary>
        /// Delete the common tags in the given list project
        /// </summary>
        /// <param name="groupName">Name of the group</param>
        /// <param name="projectName">List of the project</param>
        /// <param name="tagName">Name of the tag to be deleted</param>
        /// <returns>Returns list of tag deletion operation result</returns>
        public List<OperationResult> DeleteTags(Group groupName, List<string> projectName, string tagName)
        {
            List<ProjectDetails> projectDetails = this.GetProjectList((int)groupName);
            List<OperationResult> deleteTagResult = new List<OperationResult>();

            if (projectName == null || projectName.Count == 0)
            {
                throw new ArgumentNullException(nameof(projectName));
            }

            if (string.IsNullOrEmpty(tagName) == true)
            {
                throw new ArgumentNullException(nameof(tagName));
            }

            try
            {
                List<string> distinctProjectName = new List<string>();
                projectName = projectName.Distinct().ToList();
                distinctProjectName = projectName.OrderBy(data => data).ToList();
                for (int i = 0; i < distinctProjectName.Count; i++)
                {
                    var data = projectDetails.FirstOrDefault(x => x.Name == distinctProjectName[i].Trim());
                    if (data != null)
                    {
                        List<string> tagsList = this.GetTagsList(data.Id);
                        if (tagsList.Contains(tagName.Trim()))
                        {
                            string output = this.TagDeletion(data.Id, tagName.Trim());

                            if (string.IsNullOrEmpty(output).Equals(true))
                            {
                                deleteTagResult.Add(new OperationResult { ProjectName = distinctProjectName[i].Trim(), Comments = "Tags successfully deleted", Status = true });
                            }
                            else
                            {
                                JObject json = JObject.Parse(output);
                                deleteTagResult.Add(new OperationResult { ProjectName = distinctProjectName[i].Trim(), Comments = (string)json["message"], Status = false });
                            }
                        }
                        else
                        {
                            deleteTagResult.Add(new OperationResult { ProjectName = distinctProjectName[i].Trim(), Comments = "Given tag does not exist", Status = false });
                        }
                    }
                    else
                    {
                        deleteTagResult.Add(new OperationResult { ProjectName = distinctProjectName[i].Trim(), Comments = "Project not exist in the GitLab instances. Check the spelling of the project", Status = false });
                    }
                }
            }
            catch (WebException ex)
            {
                deleteTagResult.Add(new OperationResult { ProjectName = "-", Comments = "Exception in Tag Deletion. " + ex, Status = false });
            }

            return deleteTagResult;
        }

        /// <summary>
        /// Delete the branch particular branch from list of projects
        /// </summary>
        /// <param name="groupName">Name of the group</param>
        /// <param name="projectName">Project name list</param>
        /// <param name="branchName">Branch name which is to be deleted</param>
        public void DeleteBranch(Group groupName, List<string> projectName, string branchName)
        {
            if (projectName == null || projectName.Count == 0)
            {
                throw new ArgumentNullException(nameof(projectName));
            }

            if (string.IsNullOrEmpty(branchName) == true)
            {
                throw new ArgumentNullException(nameof(branchName));
            }

            this.BranchDeletionExcelDocument(groupName, projectName, branchName);
        }

        /// <summary>
        /// Delete the branch form single project
        /// </summary>
        /// <param name="groupName">Name of the group</param>
        /// <param name="projectName">Project name</param>
        /// <param name="branchName">branch which is to be deleted</param>
        /// <returns> returns the status of the branch deletion operation</returns>
        public string DeleteBranch(Group groupName, string projectName, string branchName)
        {
            string status;
            List<ProjectDetails> projectDetails = this.GetProjectList((int)groupName);

            if (string.IsNullOrEmpty(projectName) == true)
            {
                throw new ArgumentNullException(nameof(projectName));
            }

            if (string.IsNullOrEmpty(branchName) == true)
            {
                throw new ArgumentNullException(nameof(branchName));
            }

            try
            {
                string output;
                var data = projectDetails.FirstOrDefault(x => x.Name == projectName.Trim());
                if (data != null)
                {
                    List<string> branchesList = this.GetBranchList(data.Id).Select(branch => branch.Name).ToList();
                    if (branchesList.Contains(branchName.Trim()))
                    {
                        output = this.BranchDeletion(data.Id, branchName);

                        if (string.IsNullOrEmpty(output))
                        {
                            status = "Branch successfully deleted";
                        }
                        else
                        {
                            JObject json = JObject.Parse(output);
                            status = (string)json["message"];
                        }
                    }
                    else
                    {
                        status = "Given branch does not exist";
                    }
                }
                else
                {
                    status = "Project not exist in the GitLab instances. Check the spelling of the project";
                }
            }
            catch (WebException ex)
            {
                status = "Exception in delete branch. " + ex;
            }

            return status;
        }

        /// <summary>
        /// Delete the branches and send the log details to given email ID
        /// </summary>
        /// <param name="groupName">Name of the group</param>
        /// <param name="projectName">List of Project name</param>
        /// <param name="branchName">branch which is to be deleted</param>
        /// <param name="emailId">Mail ID to send the mail</param>
        /// <returns>return the branch deletion status as a list</returns>
        public List<OperationResult> DeleteBranch(Group groupName, List<string> projectName, string branchName, string emailId)
        {
            List<OperationResult> deleteBranchResult = new List<OperationResult>();

            if (projectName == null || projectName.Count == 0)
            {
                throw new ArgumentNullException(nameof(projectName));
            }

            if (string.IsNullOrEmpty(branchName) == true)
            {
                throw new ArgumentNullException(nameof(branchName));
            }

            if (string.IsNullOrEmpty(emailId) == true)
            {
                throw new ArgumentNullException(nameof(emailId));
            }

            Collection<ProjectData> branchDeletionData = this.BranchDeletionExcelDocument(groupName, projectName, branchName);
            for (int projectCount = 0; projectCount < branchDeletionData.Count; projectCount++)
            {
                if (branchDeletionData[projectCount].Status == "Branch successfully deleted")
                {
                    deleteBranchResult.Add(new OperationResult { ProjectName = branchDeletionData[projectCount].Project_Name, Comments = "Branch successfully deleted", Status = true });
                }
                else
                {
                    deleteBranchResult.Add(new OperationResult { ProjectName = branchDeletionData[projectCount].Project_Name, Comments = "Branch is not deleted", Status = false });
                }
            }

            string mailContent = "<html><head><style type=\"text / css\">body {font-family: Segoe UI;font-size: 14px;}table tr {border: none;}a, a:link, a:visited, a:active {text-decoration: none;color: #0066ff;}div.content a{color: #29ABE2;}img {border: none;}</style></head><body><table cellspacing=\"1\"  cellpadding=\"0\" align=\"center\" id=\"main - panel\" style=\"width: 100 %; background - color:#f2f2f2; color: rgb(75, 75, 75); border-collapse: collapse; border: 1px solid #f2f2f2;\"><tbody><tr><td ><table cellspacing=\"1\"  cellpadding=\"0\" align=\"center\" id=\"main-panel\" style=\"width: 740px; background-color:#ffffff; color: rgb(75, 75, 75); border-collapse: collapse; border: 1px solid #C6C6C6;\"><tbody><tr style=\"width: 100%;\"><td style=\"padding-top:25px;\"><table cellspacing=\"1\"  cellpadding=\"0\" align=\"center\" id=\"main-panel\" style=\"width: 700px;background-color:#ffffff;color: rgb(75, 75, 75); border-collapse: collapse;\"><tbody><tr style=\"width: 100%;\"><td style=\"padding:0px 0px;\"><table style=\" width: 100%; padding: 0px 0px;\"><tbody><tr><p>Hi Everyone</p><p>Branch deletion operation is completed successfully. Kindly follow the excel document for more information.</p></tr></tbody></table></td></tr><tr><td style=\" color:#4D4D4D;padding: 20px 0px 40px 0px ;border-bottom: 1px solid #C6C6C6; font-size: 11px;\"><div>Best regards, </div><div><b>GitLab Team </b></div></td></tr><tr><td style=\"padding: 10px 0 ;\"><table><tbody><tr><td style=\"color: #999999; font-size: 10px;\"><p>This is an automatically generated message; please do not reply to this email.</p></td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table></td></tr><tr width=\"100%\"><td style=\"background: #f2f2f2; padding: 10px 0px 10px 0px;\"><table width=\"740px\" style=\"width:740px;\"align=\"center\"><tbody><tr><td  width=\"490px\" style=\" width:490px;color: #999999; font-size: 9px;\"><p><a style=\"text-decoration: none;color: #999999;\" target=\"_top\" href=\"http://syncfusion.com/copyright\">Copyright © 2001 - 2017 Syncfusion Inc. All Rights Reserved</a></p></td><td align=\"right\" style=\" color: #999999; font-size: 9px;\"><p><a style=\"text-decoration: none;color: #999999;\" target=\"_top\" href=\"http://syncfusion.com/privacy\">Privacy Policy</a> | <a style=\"text-decoration: none;color: #999999;\" target=\"_top\" href=\"http://syncfusion.com/company/contact-us\">Contact Us</a></p></td></tr></tbody></table></td></tr></tbody></table></body></html>";
            this.mail.MailNotifications("Delete branch status", mailContent, @"GitLab_Log_file.xlsx", emailId);
            return deleteBranchResult;
        }

        /// <summary>
        /// Check the given project name is exist in the GitLab`s live data or not
        /// </summary>
        /// <param name="groupName">Name of the group</param>
        /// <param name="projectName">Project name</param>
        /// <returns>return the result by boolean</returns>
        public bool IsGitlabProject(Group groupName, string projectName)
        {
            List<ProjectDetails> projectDetails = this.GetProjectList((int)groupName);

            if (string.IsNullOrEmpty(projectName) == true)
            {
                throw new ArgumentNullException(nameof(projectName));
            }
            else
            {
                var data = projectDetails.FirstOrDefault(x => x.Name == projectName.Trim());

                if (data != null)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Protect the branch 
        /// </summary>
        /// <param name="groupName">Name of the group</param>
        /// <param name="projectName">Project name</param>
        /// <param name="branchName">Branch name which is to be protect</param>
        public void ProtectBranch(Group groupName, List<string> projectName, string branchName)
        {
            Collection<ProjectData> datatoExcel = new Collection<ProjectData>();
            Collection<string> projectUrlList = new Collection<string>();
            Collection<string> branchUrlList = new Collection<string>();
            List<ProjectDetails> projectDetails = this.GetProjectList((int)groupName);

            if (projectName == null || projectName.Count == 0)
            {
                throw new ArgumentNullException(nameof(projectName));
            }

            if (string.IsNullOrEmpty(branchName) == true)
            {
                throw new ArgumentNullException(nameof(branchName));
            }

            try
            {
                string jsonstringForProtectBranch;
                List<string> distinctProjectName = new List<string>();
                projectName = projectName.Distinct().ToList();
                distinctProjectName = projectName.OrderBy(data => data).ToList();

                for (int projectCount = 0; projectCount < distinctProjectName.Count; projectCount++)
                {
                    Console.WriteLine(distinctProjectName[projectCount]);
                    var item = projectDetails.FirstOrDefault(x => x.Name == distinctProjectName[projectCount].Trim());
                    if (item != null)
                    {
                        string branchUrl = string.Empty;
                        List<string> branchList = this.GetBranchList(item.Id).Select(branch => branch.Name).ToList();
                        if (branchList.Contains(branchName.Trim()))
                        {
                            jsonstringForProtectBranch = this.BranchProtection(item.Id, branchName);
                            JObject jsonForProtectBranch = JObject.Parse(jsonstringForProtectBranch);
                            string protecedtBranchName = (string)jsonForProtectBranch["name"];
                            branchUrl = item.Web_Url + "/tree/" + branchName.Trim();
                            if (protecedtBranchName.Equals(branchName))
                            {
                                datatoExcel.Add(new ProjectData { Project_Name = distinctProjectName[projectCount].Trim(), Source_Branch = branchName.Trim(), Protect = "Protected" });
                                projectUrlList.Add(item.Web_Url);
                                branchUrlList.Add(branchUrl);
                            }
                            else
                            {
                                datatoExcel.Add(new ProjectData { Project_Name = distinctProjectName[projectCount].Trim(), Source_Branch = branchName.Trim(), Protect = "Not protected" });
                                projectUrlList.Add(item.Web_Url);
                                branchUrlList.Add(branchUrl);
                            }
                        }
                        else
                        {
                            datatoExcel.Add(new ProjectData { Project_Name = distinctProjectName[projectCount].Trim(), Source_Branch = "-", Protect = "Branch does not exist" });
                            projectUrlList.Add(item.Web_Url);
                            branchUrlList.Add("empty");
                        }
                    }
                    else
                    {
                        datatoExcel.Add(new ProjectData { Project_Name = distinctProjectName[projectCount].Trim(), Source_Branch = "-", Protect = "Project does not exist" });
                        projectUrlList.Add("empty");
                        branchUrlList.Add("empty");
                    }
                }

                this.excel.GenerateExcelSheet(datatoExcel, projectUrlList, branchUrlList, "Branch Protection");
            }
            catch (WebException ex)
            {
                datatoExcel.Add(new ProjectData { Project_Name = "-", Source_Branch = "-", Protect = "Exception in branch protection: " + ex });
                projectUrlList.Add("empty");
                branchUrlList.Add("empty");
                this.excel.GenerateExcelSheet(datatoExcel, projectUrlList, branchUrlList, "Branch Protection");
            }
        }

        /// <summary>
        /// Merge two branch and remove the source branch if it successfully merged
        /// </summary>
        /// <param name="groupName">Name of the group</param>
        /// <param name="projectNameList">List of project name</param>
        /// <param name="sourceBranch">Source branch name</param>
        /// <param name="targetBranch">Target branch name</param>
        /// <param name="mergeRequestTitle">Title of the Merge Request</param>
        /// <returns>return the merge branch result by list</returns>
        public List<MergeRequestData> MergeBranch(Group groupName, List<string> projectNameList, string sourceBranch, string targetBranch, string mergeRequestTitle)
        {
            if (projectNameList == null || projectNameList.Count == 0)
            {
                throw new ArgumentNullException(nameof(projectNameList));
            }

            if (string.IsNullOrEmpty(sourceBranch) == true)
            {
                throw new ArgumentNullException(nameof(sourceBranch));
            }

            if (string.IsNullOrEmpty(targetBranch) == true)
            {
                throw new ArgumentNullException(nameof(targetBranch));
            }

            if (string.IsNullOrEmpty(mergeRequestTitle) == true)
            {
                throw new ArgumentNullException(nameof(mergeRequestTitle));
            }

            List<MergeRequestData> mergeBranchData = new List<MergeRequestData>();
            Dictionary<string, string> mergeRequestDetails = new Dictionary<string, string>();
            try
            {
                mergeBranchData = this.MergeRequest(groupName, projectNameList, sourceBranch, targetBranch, mergeRequestTitle);
                foreach (var mergeRequestData in mergeBranchData)
                {
                    if (mergeRequestData.MergeIid != null)
                    {
                        mergeRequestDetails.Add(mergeRequestData.ProjectName, mergeRequestData.MergeIid);
                    }
                    else
                    {
                        mergeRequestDetails.Add(mergeRequestData.ProjectName, mergeRequestData.Status);
                    }
                }

                return this.MergeAccept(groupName, mergeRequestDetails);
            }
            catch (WebException exception)
            {
                mergeBranchData.Add(new MergeRequestData { ProjectName = "-", Status = "Exception in MergeBranch method - " + exception, MergeIid = null, CommitId = null });
            }

            return mergeBranchData;
        }

        /// <summary>
        /// Create the Merge request
        /// </summary>
        /// <param name="groupName">Name of the group</param>
        /// <param name="projectNameList">Project name list</param>
        /// <param name="sourceBranch">Source branch name</param>
        /// <param name="targetBranch">Target branch name</param>
        /// <param name="mergeRequestTitle">Merge request title</param>
        /// <returns>Returns list of merge request created details</returns>
        public List<MergeRequestData> MergeRequest(Group groupName, List<string> projectNameList, string sourceBranch, string targetBranch, string mergeRequestTitle)
        {
            if (projectNameList == null || projectNameList.Count == 0)
            {
                throw new ArgumentNullException(nameof(projectNameList));
            }

            if (string.IsNullOrEmpty(sourceBranch) == true)
            {
                throw new ArgumentNullException(nameof(sourceBranch));
            }

            if (string.IsNullOrEmpty(targetBranch) == true)
            {
                throw new ArgumentNullException(nameof(targetBranch));
            }

            if (string.IsNullOrEmpty(mergeRequestTitle) == true)
            {
                throw new ArgumentNullException(nameof(mergeRequestTitle));
            }

            List<ProjectDetails> projectList = this.GetProjectList((int)groupName);
            string mergeRequestCreationString = string.Empty;
            List<MergeRequestData> mergeBranchData = new List<MergeRequestData>();
            try
            {
                foreach (string project in projectNameList)
                {
                    var projectData = projectList.FirstOrDefault(x => x.Name == project.Trim());
                    if (projectData != null)
                    {
                        List<string> branchList = this.GetBranchList(projectData.Id).Select(branch => branch.Name).ToList();
                        if (branchList.Contains(sourceBranch.Trim()) && branchList.Contains(targetBranch.Trim()))
                        {
                            ////Create the   request
                            mergeRequestCreationString = this.CreateMergeRequest(projectData.Id, sourceBranch, targetBranch, mergeRequestTitle);
                            JObject jsonForCreateMergeRequest = JObject.Parse(mergeRequestCreationString);
                            if (mergeRequestCreationString.StartsWith("{\"id\":", StringComparison.Ordinal))
                            {
                                mergeBranchData.Add(new MergeRequestData { ProjectName = projectData.Name, Status = (string)jsonForCreateMergeRequest["state"], MergeIid = (string)jsonForCreateMergeRequest["iid"], CommitId = null });
                            }
                            else if (mergeRequestCreationString.Contains("This merge request already exists").Equals(true))
                            {
                                string mergeTitle = (string)jsonForCreateMergeRequest["message"].ToString().Split('\\')[1].Split('"')[1];
                                List<MergeRequestDetails> mergeRequestList = this.GetMergeRequestDetails(projectData.Id, string.Empty);
                                var mergeData = mergeRequestList.FirstOrDefault(x => x.Title == mergeTitle);
                                if (mergeData != null)
                                {
                                    mergeBranchData.Add(new MergeRequestData { ProjectName = projectData.Name, Status = mergeData.State, MergeIid = mergeData.Iid, CommitId = null });
                                }
                            }
                            else
                            {
                                mergeBranchData.Add(new MergeRequestData { ProjectName = projectData.Name, Status = (string)jsonForCreateMergeRequest["message"].ToString(), MergeIid = null, CommitId = null });
                            }
                        }
                        else
                        {
                            mergeBranchData.Add(new MergeRequestData { ProjectName = projectData.Name, Status = "Source branch or destination branch does not exist", MergeIid = null, CommitId = null });
                        }
                    }
                    else
                    {
                        mergeBranchData.Add(new MergeRequestData { ProjectName = project, Status = "Given project does not exist", MergeIid = null, CommitId = null });
                    }
                }
            }
            catch (WebException exception)
            {
                mergeBranchData.Add(new MergeRequestData { ProjectName = "-", Status = "Exception in MergeBranch method - " + exception, MergeIid = null, CommitId = null });
            }

            return mergeBranchData;
        }

        /// <summary>
        /// Accept the merge request which is already created and whose state is open
        /// </summary>
        /// <param name="groupName">Name of the group</param>
        /// <param name="mergeRequestDetails">Dictionary of project name (key) and merge Request IID(valve) </param>
        /// <returns>Returns the accepted merge request details</returns>        
        public List<MergeRequestData> MergeAccept(Group groupName, Dictionary<string, string> mergeRequestDetails)
        {
            if (mergeRequestDetails == null || mergeRequestDetails.Count == 0)
            {
                throw new ArgumentNullException(nameof(mergeRequestDetails));
            }

            List<ProjectDetails> projectList = this.GetProjectList((int)groupName);
            string mergeRequestAcceptString = string.Empty;
            List<MergeRequestData> mergeBranchData = new List<MergeRequestData>();
            try
            {
                foreach (var project in mergeRequestDetails)
                {
                    var projectData = projectList.FirstOrDefault(x => x.Name == project.Key.Trim());
                    if (projectData != null)
                    {
                        ////Check the merge request iid 
                        int iid;
                        bool isNumeric = int.TryParse(project.Value, out iid);
                        if (isNumeric.Equals(true))
                        {
                            ////Get the merge request status
                            string mergeRequestStatus = this.GetSingleMergeRequestJson(projectData.Id, project.Value);
                            JObject jsonOfSingleMergeRequest = JObject.Parse(mergeRequestStatus);
                            if (mergeRequestStatus.StartsWith("{\"id\":", StringComparison.Ordinal))
                            {
                                if ((string)jsonOfSingleMergeRequest["state"] == "opened")
                                {
                                    ////Accepting the merge request
                                    mergeRequestAcceptString = this.AcceptMergeRequest(projectData.Id, project.Value);
                                    JObject jsonForAcceptMergeRequest = JObject.Parse(mergeRequestAcceptString);
                                    if (mergeRequestAcceptString.StartsWith("{\"id\":", StringComparison.Ordinal))
                                    {
                                        mergeBranchData.Add(new MergeRequestData { ProjectName = projectData.Name, Status = (string)jsonForAcceptMergeRequest["state"], MergeIid = project.Value, CommitId = (string)jsonForAcceptMergeRequest["merge_commit_sha"].ToString().Substring(0, 8) });
                                    }
                                    else if (mergeRequestAcceptString.StartsWith("{\"message\":", StringComparison.Ordinal))
                                    {
                                        List<string> branchList = this.GetBranchList(projectData.Id).Select(branch => branch.Name).ToList();
                                        if (this.BranchDiff(projectData.Id, (string)jsonOfSingleMergeRequest["target_branch"], (string)jsonOfSingleMergeRequest["source_branch"]).Equals(false) && branchList.Contains((string)jsonOfSingleMergeRequest["target_branch"]) && branchList.Contains((string)jsonOfSingleMergeRequest["source_branch"]))
                                        {
                                            string deletebranch = this.BranchDeletion(projectData.Id, (string)jsonOfSingleMergeRequest["source_branch"]);
                                            if (string.IsNullOrEmpty(deletebranch))
                                            {
                                                mergeBranchData.Add(new MergeRequestData { ProjectName = projectData.Name, Status = "No change, Source branch deleted", MergeIid = project.Value, CommitId = null });
                                            }
                                            else
                                            {
                                                JObject json = JObject.Parse(deletebranch);
                                                mergeBranchData.Add(new MergeRequestData { ProjectName = projectData.Name, Status = "No change, " + (string)json["message"], MergeIid = project.Value, CommitId = null });
                                            }                                            
                                        }
                                        else if (branchList.Contains((string)jsonOfSingleMergeRequest["target_branch"]) && branchList.Contains((string)jsonOfSingleMergeRequest["source_branch"]))
                                        {
                                            mergeBranchData.Add(new MergeRequestData { ProjectName = projectData.Name, Status = "conflict", MergeIid = project.Value, CommitId = null });
                                        }
                                        else
                                        {
                                            mergeBranchData.Add(new MergeRequestData { ProjectName = projectData.Name, Status = "Source or Destination branch does not exist", MergeIid = project.Value, CommitId = null });
                                        }
                                    }
                                }
                                else
                                {
                                    mergeBranchData.Add(new MergeRequestData { ProjectName = projectData.Name, Status = (string)jsonOfSingleMergeRequest["state"], MergeIid = project.Value, CommitId = null });
                                }
                            }
                            else
                            {
                                mergeBranchData.Add(new MergeRequestData { ProjectName = projectData.Name, Status = (string)jsonOfSingleMergeRequest["message"], MergeIid = project.Value, CommitId = null });
                            }
                        }
                        else
                        {
                            mergeBranchData.Add(new MergeRequestData { ProjectName = project.Key, Status = project.Value, CommitId = null, MergeIid = null });
                        }
                    }
                    else
                    {
                        mergeBranchData.Add(new MergeRequestData { ProjectName = project.Key, Status = "Given project does not exist", MergeIid = null, CommitId = null });
                    }
                }
            }
            catch (WebException exception)
            {
                mergeBranchData.Add(new MergeRequestData { ProjectName = "-", Status = "Exception in MergeBranch method - " + exception, MergeIid = null, CommitId = null });
            }

            return mergeBranchData;
        }

        /// <summary>
        /// Freeze/unfreeze the project 
        /// </summary>
        /// <param name="groupName">Name of the group</param>
        /// <param name="projectList">List of project which is to be freeze/unfreeze</param>
        /// <param name="branchName">Name of the branch to be freeze/unfreeze</param>
        /// <param name="engineerEmail">Engineer email ID which is to be freeze/unfreeze, null for freeze all engineers</param>
        /// <param name="freeze">bool value of freeze (freeze = true, unfreeze = false)</param>
        /// <returns>Returns the restricted project details in list</returns>
        public List<OperationResult> GitlabMergePermission(Group groupName, List<string> projectList, string branchName, string engineerEmail, bool freeze)
         {
            //if (projectList == null || projectList.Count == 0)
            //{
            //    throw new ArgumentNullException(nameof(projectList));
            //}
            //List<string> projectList = this.GetProjectList((int)groupName).Select(x => x.Name).ToList();

            if (engineerEmail != null)
            {
                if (engineerEmail.EndsWith("@syncfusion.com", StringComparison.OrdinalIgnoreCase) == false)
                {
                    throw new ArgumentNullException(nameof(engineerEmail));
                }
            }

            List<OperationResult> mergeRestrict = new List<OperationResult>();
            string accessLevel = string.Empty;
            try
            {
                if (freeze.Equals(true))
                {
                    accessLevel = "30";
                }
                else if (freeze.Equals(false))
                {
                    accessLevel = "40";
                }

                List<ProjectDetails> projectDetails = this.GetProjectList((int)groupName);
                List<string> distinctProjectName = new List<string>();
                List<string> projectUrl = new List<string>();
                projectList = projectList.Distinct().ToList();
                distinctProjectName = projectList.OrderBy(data => data).ToList();
                for (int projectCount = 0; projectCount < distinctProjectName.Count; projectCount++)
                {
                    //if(distinctProjectName[projectCount].StartsWith("ej2"))
                    //{
                        var projectData = projectDetails.FirstOrDefault(x => x.Name == distinctProjectName[projectCount].Trim());
                        if (projectData != null)
                        {
                            Console.WriteLine(projectData.Name);
                            projectUrl.Add(projectData.Web_Url);
                            int userAddedCount = 0;
                            List<UserDetails> projectMembers = this.GetMembersList(projectData.Id);
                            if (projectMembers.Count.Equals(0) == false)
                            {
                                if (engineerEmail == null)
                                {
                                    bool userAddingInProject;

                                    for (int memberCount = 0; memberCount < projectMembers.Count; memberCount++)
                                    {
                                        //if (projectMembers[memberCount].State != "40")
                                        //{
                                        userAddingInProject = this.ChangeAccessLevelInProject(projectData.Id, projectMembers[memberCount].Id, accessLevel);


                                        if (userAddingInProject.Equals(true))
                                        {
                                            userAddedCount++;
                                        }
                                        // }
                                    }

                                    if (projectMembers.Count.Equals(userAddedCount))
                                    {
                                        mergeRestrict.Add(new OperationResult { ProjectName = distinctProjectName[projectCount], Comments = "Permission has been changed", Status = true });
                                    }
                                    else
                                    {
                                        mergeRestrict.Add(new OperationResult { ProjectName = distinctProjectName[projectCount], Comments = "Some users permission does not changed", Status = false });
                                    }
                                }
                                else if (engineerEmail.EndsWith("@syncfusion.com", StringComparison.OrdinalIgnoreCase) == true)
                                {
                                    var projectPermission = projectMembers.FirstOrDefault(x => x.Email == engineerEmail.Trim());
                                    if (projectPermission != null)
                                    {
                                        bool userAddingInProject = this.ChangeAccessLevelInProject(projectData.Id, projectPermission.Id, accessLevel);
                                        if (userAddingInProject.Equals(true))
                                        {
                                            mergeRestrict.Add(new OperationResult { ProjectName = distinctProjectName[projectCount], Comments = "Permission has been changed", Status = true });
                                        }
                                        else
                                        {
                                            mergeRestrict.Add(new OperationResult { ProjectName = distinctProjectName[projectCount], Comments = "Permission does not changed", Status = false });
                                        }
                                    }
                                    else
                                    {
                                        mergeRestrict.Add(new OperationResult { ProjectName = distinctProjectName[projectCount], Comments = "Given mail ID does not have permission", Status = false });
                                    }
                                }
                            }
                            else
                            {
                                mergeRestrict.Add(new OperationResult { ProjectName = distinctProjectName[projectCount], Comments = "Project does not have members", Status = false });
                            }
                        }
                        else
                        {
                            // projectUrl.Add("https://gitlab.syncfusion.com/content/xamarinstore-docs");
                            mergeRestrict.Add(new OperationResult { ProjectName = distinctProjectName[projectCount], Comments = "Given project does not exist", Status = false });
                        }
                    //}

                   
                }
                ExcelOperations excels = new ExcelOperations();
                List<string> dummy = new List<string>();
                excels.GenerateExcelSheet(mergeRestrict, projectUrl, dummy, "Merge Protect");
                mail.MailNotifications("Essential Studio project frozen list", "", @"GitLab_Log_file.xlsx", "arunkumar.nagarajan@syncfusion.com");
            }
            catch (WebException exception)
            {
                mergeRestrict.Add(new OperationResult { ProjectName = exception.ToString(), Comments = "Exception in GitLabMergeRestrict method - " + exception, Status = false });
            }

            return mergeRestrict;
        }

        /// <summary>
        /// Check the master permission in particular project
        /// </summary>
        /// <param name="groupName">Name of the group</param>
        /// <param name="project">Project name where you going to check</param>
        /// <param name="emailId">Mail ID of the person to check the master permission</param>
        /// <returns>returns True if the given mail ID exist as a master or else returns Fasle </returns>
        public bool ProjectPermission(Group groupName, string project, string emailId)
        {
            if (string.IsNullOrEmpty(project) == true)
            {
                throw new ArgumentNullException(nameof(project));
            }

            if (emailId.EndsWith("@syncfusion.com", StringComparison.OrdinalIgnoreCase) == false || string.IsNullOrEmpty(emailId) == true)
            {
                throw new ArgumentNullException(nameof(emailId));
            }

            List<ProjectDetails> projectDetails = this.GetProjectList((int)groupName);

            var projectData = projectDetails.FirstOrDefault(x => x.Name == project.Trim());

            if (projectData != null)
            {
                List<UserDetails> masterList = this.GetMembersList(projectData.Id);
                var masterData = masterList.FirstOrDefault(x => x.Email == emailId.Trim() && x.State == "40");
                if (masterData != null)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Validate the Merge request
        /// </summary>
        /// <param name="groupName">Name of the group</param>
        /// <param name="mergeRequest">Merge request URL</param>
        /// <param name="targetBranch">Target branch name</param>
        /// <returns>Returns the merge request status</returns>
        public string ValidateMergeRequest(Group groupName, string mergeRequest, string targetBranch)
        {
            if (mergeRequest.StartsWith("https://gitlab.syncfusion.com/", StringComparison.OrdinalIgnoreCase) == false || string.IsNullOrEmpty(mergeRequest) == true)
            {
                throw new ArgumentNullException(nameof(mergeRequest));
            }

            if (string.IsNullOrEmpty(targetBranch) == true)
            {
                throw new ArgumentNullException(nameof(targetBranch));
            }

            List<ProjectDetails> projectDetails = this.GetProjectList((int)groupName);
            string mergeRequestStatus;

            var iid = mergeRequest.Split('/')[6];
            var projectPath = mergeRequest.Split('/')[3] + "/" + mergeRequest.Split('/')[4];
            var projectData = projectDetails.FirstOrDefault(x => x.Path_With_Namespace == projectPath.Trim());
            if (projectData != null)
            {
                List<string> branchList = this.GetBranchList(projectData.Id).Select(branch => branch.Name).ToList();
                if (branchList.Contains(targetBranch))
                {
                    List<MergeRequestDetails> mergeRequestDetails = this.GetMergeRequestDetails(projectData.Id, "state=opened");
                    var iidDataForOpenMergeRequest = mergeRequestDetails.FirstOrDefault(x => x.Iid == iid);
                    if (iidDataForOpenMergeRequest != null)
                    {
                        if (iidDataForOpenMergeRequest.TargetBranch.Equals(targetBranch.Trim()))
                        {
                            mergeRequestStatus = "Valid";
                        }
                        else
                        {
                            mergeRequestStatus = "The target branch does not match on merge request";
                        }
                    }
                    else
                    {
                        mergeRequestDetails = this.GetMergeRequestDetails(projectData.Id, "state=merged");
                        var iidDataForMergedMergeRequest = mergeRequestDetails.FirstOrDefault(x => x.Iid == iid);
                        if (iidDataForMergedMergeRequest != null)
                        {
                            mergeRequestStatus = "Already Merged";
                        }
                        else
                        {
                            mergeRequestStatus = "The merge request does not exist on the project";
                        }
                    }
                }
                else
                {
                    mergeRequestStatus = "Target branch does not exist";
                }
            }
            else
            {
                mergeRequestStatus = "Given Project does not exist";
            }

            return mergeRequestStatus;
        }

        /// <summary>
        /// To check the status of the merge request
        /// </summary>
        /// <param name="groupName">Name of the group</param>
        /// <param name="mergeRequest">HTTPS URL of the merge request</param>
        /// <returns>Returns the status of the given merge request</returns>
        public string MergeRequestStatus(Group groupName, string mergeRequest)
        {
            if (mergeRequest.StartsWith("https://gitlab.syncfusion.com/", StringComparison.OrdinalIgnoreCase) == false || string.IsNullOrEmpty(mergeRequest) == true)
            {
                throw new ArgumentNullException(nameof(mergeRequest));
            }

            string mergeRequestStatus = string.Empty, jsonOutputForMergeRequest;
            List<ProjectDetails> projectDetails = this.GetProjectList((int)groupName);
            var iid = mergeRequest.Split('/')[6];
            var projectPath = mergeRequest.Split('/')[3] + "/" + mergeRequest.Split('/')[4];
            var projectData = projectDetails.FirstOrDefault(x => x.Path_With_Namespace == projectPath.Trim());
            if (projectData != null)
            {
                List<MergeRequestDetails> mergeRequestDetails = this.GetMergeRequestDetails(projectData.Id, string.Empty);
                var iidDataForOpenMergeRequest = mergeRequestDetails.FirstOrDefault(x => x.Iid == iid);
                if (iidDataForOpenMergeRequest != null)
                {
                    jsonOutputForMergeRequest = this.GetCommitDetails(projectData.Id, iidDataForOpenMergeRequest.Sha);
                    if (jsonOutputForMergeRequest != null)
                    {
                        JObject json = JObject.Parse(jsonOutputForMergeRequest);
                        if ((string)json["status"] == null)
                        {
                            mergeRequestStatus = "null";
                        }
                        else
                        {
                            mergeRequestStatus = (string)json["status"];
                        }
                    }
                }
                else
                {
                    mergeRequestStatus = "The merge request does not exist on the project";
                }
            }
            else
            {
                mergeRequestStatus = "Given Project does not exist";
            }

            return mergeRequestStatus;
        }

        /// <summary>
        /// To get the specific project ID of the Essential Studio or Data Science group`s project
        /// </summary>
        /// <param name="groupName">Name of the group</param>
        /// <param name="projectName">Name of the project</param>
        /// <returns>Returns the project ID of given project</returns>
        public string GetProjectId(Group groupName, string projectName)
        {
            if (string.IsNullOrEmpty(projectName) == true)
            {
                throw new ArgumentNullException(nameof(projectName));
            }

            string projectId = null;
            List<ProjectDetails> projectList = this.GetProjectList((int)groupName);
            var projectData = projectList.FirstOrDefault(x => x.Name == projectName.Trim());
            if (projectData != null)
            {
                projectId = projectData.Id;
            }
            else
            {
                projectId = "Given project not found";
            }

            return projectId;
        }

        /// <summary>
        /// Removing the merged branches which created 10 days before
        /// </summary>
        /// <param name="groupName">Name of the group</param>
        /// <param name="removeProtectedBranch">True for removing protected branch and False for skipping protected branch</param>
        public void RemoveMergedBranches(Group groupName, bool removeProtectedBranch)
        {
            List<ProjectDetails> projectList = this.GetProjectList((int)groupName);
            if (projectList.Count != 0)
            {
                List<string> groupProjectList = projectList.Select(name => name.Name).ToList();
                this.RemoveMergedBranches(groupName, groupProjectList, removeProtectedBranch);
            }
        }
        public List<UserDetails> GetGroupMemberList()
        {
            return this.GetGroupMembersList("581");
        }

        /// <summary>
        /// Removing the merged branches which created 10 days before
        /// </summary>
        /// <param name="groupName">Name of the group</param>
        /// <param name="projectNameList">List of the project where need to remove the merged branchs</param>
        /// <param name="removeProtectedBranch">True for removing protected branch and False for skipping protected branch</param>
        public void RemoveMergedBranches(Group groupName, List<string> projectNameList, bool removeProtectedBranch)
        {
            if (projectNameList.Count == 0 || projectNameList == null)
            {
                throw new ArgumentNullException(nameof(projectNameList));
            }

            List<ProjectDetails> projectList = this.GetProjectList((int)groupName);
            foreach (string projectName in projectNameList)
            {
                var projectData = projectList.FirstOrDefault(x => x.Name == projectName);
                if (projectData != null)
                {
                    List<BranchDetails> branchList = this.GetBranchList(projectData.Id);
                    if (branchList.Count != 0)
                    {
                        foreach (var branch in branchList)
                        {
                            if (branch.Merged.Equals(true) == true)
                            {
                                if (removeProtectedBranch.Equals(true) == true)
                                {
                                    if (this.GetWeekdaysBetween(branch.Commit.AuthoredDate, DateTime.Now) > 10)
                                    {
                                        if (branch.Name.Equals("development") == false || branch.Name.Equals("master") == false)
                                        {
                                            this.BranchDeletion(projectData.Id, branch.Name);
                                        }
                                    }
                                }
                                else
                                {
                                    if (branch.Protected.Equals(false) == true)
                                    {
                                        if (this.GetWeekdaysBetween(branch.Commit.AuthoredDate, DateTime.Now) > 10)
                                        {
                                            if (branch.Name.Equals("development") == false || branch.Name.Equals("master") == false)
                                            {
                                                this.BranchDeletion(projectData.Id, branch.Name);
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

        /// <summary>
        /// Create the branch in common method
        /// </summary>
        /// <param name="groupName">Name of the group</param>
        /// <param name="projectName">Project name</param>
        /// <param name="sourceBranch">Source branch</param>
        /// <param name="destinationBranch">destination branch</param>
        /// <returns>Return the result data List</returns>
        protected Collection<ProjectData> BranchCreationExcelDocument(Group groupName, List<string> projectName, string sourceBranch, string destinationBranch)
        {
            Collection<ProjectData> datatoExcel = new Collection<ProjectData>();
            Collection<string> projectUrlList = new Collection<string>();
            Collection<string> branchUrlList = new Collection<string>();
            List<ProjectDetails> projectDetails = this.GetProjectList((int)groupName);

            if (projectName == null || projectName.Count == 0)
            {
                throw new ArgumentNullException(nameof(projectName));
            }

            if (string.IsNullOrEmpty(sourceBranch) == true)
            {
                throw new ArgumentNullException(nameof(sourceBranch));
            }

            if (string.IsNullOrEmpty(destinationBranch) == true)
            {
                throw new ArgumentNullException(nameof(destinationBranch));
            }

            try
            {
                string jsonstringForCreateBranch, jsonstringForProtectBranch;
                if (destinationBranch.Contains("/"))
                {
                    List<string> distinctProjectName = new List<string>();
                    projectName = projectName.Distinct().ToList();
                    distinctProjectName = projectName.OrderBy(data => data).ToList();

                    for (int projectCount = 0; projectCount < distinctProjectName.Count; projectCount++)
                    {
                        ////Console.WriteLine(distinctProjectName[i]);
                        var item = projectDetails.FirstOrDefault(x => x.Name == distinctProjectName[projectCount].Trim());
                        if (item != null)
                        {
                            string branchUrl = string.Empty;
                            List<string> branchList = this.GetBranchList(item.Id).Select(branch => branch.Name).ToList();
                            List<string> tagsList = this.GetTagsList(item.Id);
                            var concatList = branchList.Concat(tagsList);
                            List<string> branchTagList = concatList.ToList();
                            if (branchTagList.Contains(sourceBranch.Trim()))
                            {
                                jsonstringForCreateBranch = this.BranchCreation(item.Id, destinationBranch, sourceBranch);
                                JObject jsonForCreateBranch = JObject.Parse(jsonstringForCreateBranch);
                                string branchName = (string)jsonForCreateBranch["name"];
                                branchUrl = item.Web_Url + "/tree/" + destinationBranch.Trim();                                
                                if (branchName != null)
                                {
                                    jsonstringForProtectBranch = this.BranchProtection(item.Id, destinationBranch);
                                    JObject jsonForProtectBranch = JObject.Parse(jsonstringForProtectBranch);
                                    string protectBranchName = (string)jsonForProtectBranch["name"];
                                    if (protectBranchName.Equals(destinationBranch))
                                    {
                                        datatoExcel.Add(new ProjectData { Project_Name = distinctProjectName[projectCount].Trim(), Source_Branch = sourceBranch.Trim(), Destination_Branch = destinationBranch.Trim(), Status = "Branch successfully created", Protect = "Protected" });
                                        projectUrlList.Add(item.Web_Url);
                                        branchUrlList.Add(branchUrl);
                                    }
                                    else
                                    {
                                        datatoExcel.Add(new ProjectData { Project_Name = distinctProjectName[projectCount].Trim(), Source_Branch = sourceBranch.Trim(), Destination_Branch = destinationBranch.Trim(), Status = "Branch successfully created", Protect = "Not protected" });
                                        projectUrlList.Add(item.Web_Url);
                                        branchUrlList.Add(branchUrl);
                                    }
                                }
                                else
                                {
                                    if ((string)jsonForCreateBranch["message"] == "Invalid reference name")
                                    {
                                        datatoExcel.Add(new ProjectData { Project_Name = distinctProjectName[projectCount].Trim(), Source_Branch = sourceBranch.Trim(), Destination_Branch = "-", Status = (string)jsonForCreateBranch["message"], Protect = "Not protected" });
                                        projectUrlList.Add(item.Web_Url);
                                        branchUrlList.Add("empty");
                                    }
                                    else if ((string)jsonForCreateBranch["message"] == "Branch already exists")
                                    {
                                        jsonstringForProtectBranch = this.BranchProtection(item.Id, destinationBranch);
                                        JObject jsonForProtectBranch = JObject.Parse(jsonstringForProtectBranch);
                                        string protectBranchName = (string)jsonForProtectBranch["name"];
                                        if (protectBranchName.Equals(destinationBranch))
                                        {
                                            datatoExcel.Add(new ProjectData { Project_Name = distinctProjectName[projectCount].Trim(), Source_Branch = sourceBranch.Trim(), Destination_Branch = destinationBranch.Trim(), Status = (string)jsonForCreateBranch["message"], Protect = "Protected" });
                                            projectUrlList.Add(item.Web_Url);
                                            branchUrlList.Add(branchUrl);
                                        }
                                        else
                                        {
                                            datatoExcel.Add(new ProjectData { Project_Name = distinctProjectName[projectCount].Trim(), Source_Branch = sourceBranch.Trim(), Destination_Branch = destinationBranch.Trim(), Status = (string)jsonForCreateBranch["message"], Protect = "Not protected" });
                                            projectUrlList.Add(item.Web_Url);
                                            branchUrlList.Add(branchUrl);
                                        }
                                    }
                                    else
                                    {
                                        datatoExcel.Add(new ProjectData { Project_Name = distinctProjectName[projectCount].Trim(), Source_Branch = sourceBranch.Trim(), Destination_Branch = destinationBranch.Trim(), Status = (string)jsonForCreateBranch["message"], Protect = "Not protected" });
                                        projectUrlList.Add(item.Web_Url);
                                        branchUrlList.Add(branchUrl);
                                    }
                                }
                            }
                            else
                            {
                                datatoExcel.Add(new ProjectData { Project_Name = distinctProjectName[projectCount].Trim(), Source_Branch = sourceBranch.Trim(), Destination_Branch = "-", Status = "Given source branch not exist", Protect = "-" });
                                projectUrlList.Add(item.Web_Url);
                                branchUrlList.Add("empty");
                            }
                        }
                        else
                        {
                            datatoExcel.Add(new ProjectData { Project_Name = distinctProjectName[projectCount].Trim(), Source_Branch = sourceBranch.Trim(), Destination_Branch = "-", Status = "Project not exist in the GitLab instances. Check the spelling of the project", Protect = "-" });
                            projectUrlList.Add("empty");
                            branchUrlList.Add("empty");
                        }
                    }

                    this.excel.GenerateExcelSheet(datatoExcel, projectUrlList, branchUrlList, "Branch Creation");
                }
                else
                {
                    datatoExcel.Add(new ProjectData { Project_Name = "-", Source_Branch = "-", Destination_Branch = "-", Status = "Destination branch must contain branch type \"/\" before vx.x.x (ex: hotfix/v1.0.0)", Protect = "-" });
                    projectUrlList.Add("empty");
                    branchUrlList.Add("empty");
                    this.excel.GenerateExcelSheet(datatoExcel, projectUrlList, branchUrlList, "Branch Creation");
                }

                return datatoExcel;
            }
            catch (WebException ex)
            {
                datatoExcel.Add(new ProjectData { Project_Name = "-", Source_Branch = "-", Destination_Branch = "-", Status = "Exception in branch creation: " + ex, Protect = "-" });
                projectUrlList.Add("empty");
                branchUrlList.Add("empty");
                this.excel.GenerateExcelSheet(datatoExcel, projectUrlList, branchUrlList, "Branch Creation");

                return datatoExcel;
            }
        }

        /// <summary>
        /// Branch creation curl operation
        /// </summary>
        /// <param name="projectId">Project ID</param>
        /// <param name="destinationBranch">Destination branch name</param>
        /// <param name="sourceBranch">Source branch name</param>
        /// <returns>Return the output of the curl comment</returns>
        protected string BranchCreation(string projectId, string destinationBranch, string sourceBranch)
        {
            if (string.IsNullOrEmpty(projectId) == true)
            {
                throw new ArgumentNullException(nameof(projectId));
            }

            if (string.IsNullOrEmpty(destinationBranch) == true)
            {
                throw new ArgumentNullException(nameof(destinationBranch));
            }

            if (string.IsNullOrEmpty(sourceBranch) == true)
            {
                throw new ArgumentNullException(nameof(sourceBranch));
            }

            return this.CurlOperation("--request POST --header \"PRIVATE-TOKEN:" + Token + "\" \"https://gitlab.syncfusion.com/api/v4/projects/" + projectId + "/repository/branches?branch=" + destinationBranch.Trim() + "&ref=" + sourceBranch.Trim() + "\"");
        }

        /// <summary>
        /// Tag creation curl operation
        /// </summary>
        /// <param name="projectId">Project ID</param>
        /// <param name="tagName">Tag name</param>
        /// <param name="sourceBranch">Source branch name</param>
        /// <returns>Return the output of the curl comment</returns>
        protected string TagCreation(string projectId, string tagName, string sourceBranch)
        {
            if (string.IsNullOrEmpty(projectId) == true)
            {
                throw new ArgumentNullException(nameof(projectId));
            }

            if (string.IsNullOrEmpty(tagName) == true)
            {
                throw new ArgumentNullException(nameof(tagName));
            }

            if (string.IsNullOrEmpty(sourceBranch) == true)
            {
                throw new ArgumentNullException(nameof(sourceBranch));
            }

            return this.CurlOperation("--header \"PRIVATE-TOKEN:" + Token + "\" -X POST \"https://gitlab.syncfusion.com/api/v4/projects/" + projectId + "/repository/tags\" --data \"tag_name=" + tagName.Trim() + "\" --data \"ref=" + sourceBranch.Trim() + "\"");
        }

        /// <summary>
        /// Tag deletion CURL operation
        /// </summary>
        /// <param name="projectId">ID of the project</param>
        /// <param name="tagsName">Name of the Tag</param>
        /// <returns> ruturn the Json string</returns>
        protected string TagDeletion(string projectId, string tagsName)
        {
            if (string.IsNullOrEmpty(projectId) == true)
            {
                throw new ArgumentNullException(nameof(projectId));
            }

            if (string.IsNullOrEmpty(tagsName) == true)
            {
                throw new ArgumentNullException(nameof(tagsName));
            }

            return this.CurlOperation("--request DELETE --header \"PRIVATE-TOKEN:" + Token + "\" \"https://gitlab.syncfusion.com/api/v4/projects/" + projectId + "/repository/tags/" + tagsName);
        }

        /// <summary>
        /// Branch protection cURL operation 
        /// </summary>
        /// <param name="projectId">ID of the project</param>
        /// <param name="branchName">Name of the project</param>
        /// <returns>Returns the protected status in JSON</returns>
        protected string BranchProtection(string projectId, string branchName)
        {
            if (string.IsNullOrEmpty(projectId) == true)
            {
                throw new ArgumentNullException(nameof(projectId));
            }

            if (string.IsNullOrEmpty(branchName) == true)
            {
                throw new ArgumentNullException(nameof(branchName));
            }

            return this.CurlOperation("curl --request PUT --header \"PRIVATE-TOKEN:" + Token + "\" https://gitlab.syncfusion.com/api/v4/projects/" + projectId + "/repository/branches/" + branchName + "/protect");
        }

        /// <summary>
        /// Delete the branch for common method
        /// </summary>
        /// <param name="groupName">Name of the group</param>
        /// <param name="projectName">Project name list</param>
        /// <param name="branchName">branch name which is to be deleted</param>
        /// <returns>Returns the delete branch data</returns>
        protected Collection<ProjectData> BranchDeletionExcelDocument(Group groupName, List<string> projectName, string branchName)
        {
            List<ProjectDetails> projectDetails = this.GetProjectList((int)groupName);
            Collection<ProjectData> datatoExcel = new Collection<ProjectData>();
            Collection<string> tagsUrlList = new Collection<string>();
            Collection<string> projectUrlList = new Collection<string>();

            if (projectName == null || projectName.Count == 0)
            {
                throw new ArgumentNullException(nameof(projectName));
            }

            if (string.IsNullOrEmpty(branchName) == true)
            {
                throw new ArgumentNullException(nameof(branchName));
            }

            try
            {
                string url;
                string output;
                List<string> distinctProjectName = new List<string>();
                projectName = projectName.Distinct().ToList();
                distinctProjectName = projectName.OrderBy(data => data).ToList();
                for (int projectCount = 0; projectCount < distinctProjectName.Count; projectCount++)
                {
                    var data = projectDetails.FirstOrDefault(x => x.Name == distinctProjectName[projectCount].Trim());
                    if (data != null)
                    {
                        url = data.Web_Url;
                        List<string> branchesList = this.GetBranchList(data.Id).Select(branch => branch.Name).ToList();
                        if (branchesList.Contains(branchName.Trim()))
                        {
                            output = this.BranchDeletion(data.Id, branchName);

                            if (string.IsNullOrEmpty(output))
                            {
                                datatoExcel.Add(new ProjectData { Project_Name = distinctProjectName[projectCount].Trim(), Source_Branch = branchName.Trim(), Status = "Branch successfully deleted" });
                                projectUrlList.Add(url);
                            }
                            else
                            {
                                JObject json = JObject.Parse(output);
                                datatoExcel.Add(new ProjectData { Project_Name = distinctProjectName[projectCount].Trim(), Source_Branch = branchName.Trim(), Status = (string)json["message"] });
                                projectUrlList.Add(url);
                            }
                        }
                        else
                        {
                            datatoExcel.Add(new ProjectData { Project_Name = distinctProjectName[projectCount].Trim(), Source_Branch = branchName.Trim(), Status = "Given branch does not exist" });
                            projectUrlList.Add(url);
                        }
                    }
                    else
                    {
                        datatoExcel.Add(new ProjectData { Project_Name = distinctProjectName[projectCount].Trim(), Source_Branch = branchName.Trim(), Status = "Project not exist in the GitLab instances. Check the spelling of the project" });
                        projectUrlList.Add("empty");
                    }
                }

                this.excel.GenerateExcelSheet(datatoExcel, projectUrlList, tagsUrlList, "Branch Deletion");
            }
            catch (WebException ex)
            {
                datatoExcel.Add(new ProjectData { Project_Name = "-", Source_Branch = "-", Destination_Branch = "-", Status = "Exception in branch deletion. " + ex });
                projectUrlList.Add("empty");
                this.excel.GenerateExcelSheet(datatoExcel, projectUrlList, tagsUrlList, "Branch Deletion");
            }

            return datatoExcel;
        }

        /// <summary>
        /// Deleting operation of branch
        /// </summary>
        /// <param name="projectId">Project ID of project</param>
        /// <param name="branchName">Branch name which is to be deleted</param>
        /// <returns>Returns the output line of the cmd screen</returns>
        protected string BranchDeletion(string projectId, string branchName)
        {
            if (string.IsNullOrEmpty(projectId) == true)
            {
                throw new ArgumentNullException(nameof(projectId));
            }

            if (string.IsNullOrEmpty(branchName) == true)
            {
                throw new ArgumentNullException(nameof(branchName));
            }

            return this.CurlOperation("--request DELETE --header \"PRIVATE-TOKEN:" + Token + "\" \"https://gitlab.syncfusion.com/api/v4/projects/" + projectId + "/repository/branches/" + branchName);
        }

        /// <summary>
        /// Get the branch list to check the given source branch exist or not
        /// </summary>
        /// <param name="projectId">Project ID</param>
        /// <returns>return the branch details of the corresponding project</returns>
        protected List<BranchDetails> GetBranchList(string projectId)
        {
            List<BranchDetails> branchDetails = new List<BranchDetails>();
            branchDetails = JsonConvert.DeserializeObject<List<BranchDetails>>(this.HttpOperation("https://gitlab.syncfusion.com/api/v3/projects/" + projectId + "/repository/branches?private_token=" + Token));            
            return branchDetails.OrderBy(x => x.Name).ToList();
        }

        /// <summary>
        /// Get the tags list
        /// </summary>
        /// <param name="projectId">ID of the project</param>
        /// <returns>returns the project`s tags list</returns>
        protected List<string> GetTagsList(string projectId)
        {
            List<string> tagsDetails = new List<string>();
            dynamic json = JsonConvert.DeserializeObject(this.HttpOperation("https://gitlab.syncfusion.com/api/v3/projects/" + projectId + "/repository/tags?private_token=" + Token));
            foreach (var item in json)
            {
                string branchName = item.name;
                tagsDetails.Add(branchName);
            }

            tagsDetails.Sort();
            return tagsDetails;
        }       

        /// <summary>
        /// Get the project list by using Enum
        /// </summary>
        /// <param name="groupId">Id of the group</param>
        /// <returns>Returns the list of project</returns>
        protected List<ProjectDetails> GetProjectList(int groupId)
        {
            bool check = true;
            List<ProjectDetails> projectDetails = new List<ProjectDetails>();
            List<ProjectDetails> projectData = new List<ProjectDetails>();

            int pageCount = 0;
            while (check)
            {
                pageCount = pageCount + 1;
                var jSONString = this.HttpOperation("http://gitlab.syncfusion.com/api/v4/groups/" + groupId + "/projects?private_token=" + Token + "&per_page=100&page=" + pageCount);
                projectData = JsonConvert.DeserializeObject<List<ProjectDetails>>(jSONString);
                if (jSONString == "[]")
                {
                    goto exitLoop;
                }

                projectDetails.AddRange(projectData);
            }

            exitLoop: ;

            var sortedList = projectDetails.OrderBy(a => a.Name);
            sortedList.ToList();
            return sortedList.ToList();
        }

        /// <summary>
        /// Create merge request via cURL comment
        /// </summary>
        /// <param name="projectId">ID of the project</param>
        /// <param name="sourceBranch">Source branch name</param>
        /// <param name="targetBranch">Target branch name</param>
        /// <param name="mergeRequestTitle">Merge request title</param>
        /// <returns>returns the merge request creation JSON</returns>
        protected string CreateMergeRequest(string projectId, string sourceBranch, string targetBranch, string mergeRequestTitle)
        {
            if (string.IsNullOrEmpty(projectId) == true)
            {
                throw new ArgumentNullException(nameof(projectId));
            }

            if (string.IsNullOrEmpty(sourceBranch) == true)
            {
                throw new ArgumentNullException(nameof(sourceBranch));
            }

            if (string.IsNullOrEmpty(targetBranch) == true)
            {
                throw new ArgumentNullException(nameof(targetBranch));
            }

            if (string.IsNullOrEmpty(mergeRequestTitle) == true)
            {
                throw new ArgumentNullException(nameof(mergeRequestTitle));
            }

            return this.CurlOperation("--request POST --header \"PRIVATE-TOKEN:" + Token + "\" --data \"source_branch=" + sourceBranch + "&target_branch=" + targetBranch + "&title=" + mergeRequestTitle + "-Build\" \"https://gitlab.syncfusion.com/api/v4/projects/" + projectId + "/merge_requests");
        }

        /// <summary>
        /// Accepting the merge request via cURL
        /// </summary>
        /// <param name="projectId">ID of the project</param>
        /// <param name="mergeRequestIid">IID if the merge request</param>
        /// <returns>Returns the accept merge request JSON</returns>
        protected string AcceptMergeRequest(string projectId, string mergeRequestIid)
        {
            if (string.IsNullOrEmpty(projectId) == true)
            {
                throw new ArgumentNullException(nameof(projectId));
            }

            if (string.IsNullOrEmpty(mergeRequestIid) == true)
            {
                throw new ArgumentNullException(nameof(mergeRequestIid));
            }

            return this.CurlOperation("--request PUT --header \"PRIVATE-TOKEN:" + Token + "\" --data \"should_remove_source_branch=true\" \"https://gitlab.syncfusion.com/api/v4/projects/" + projectId + "/merge_requests/" + mergeRequestIid + "/merge\"");
        }

        /// <summary>
        /// Checking the difference between two branch
        /// </summary>
        /// <param name="projectId">ID of the project</param>
        /// <param name="mainBranch">The branch which is the target branch</param>
        /// <param name="checkingBranch">Source branch which is to be checked</param>
        /// <returns>Returns true, if the branches code have diff</returns>
        protected bool BranchDiff(string projectId, string mainBranch, string checkingBranch)
        {
            try
            {
                if (string.IsNullOrEmpty(projectId) == true)
                {
                    throw new ArgumentNullException(nameof(projectId));
                }

                if (string.IsNullOrEmpty(mainBranch) == true)
                {
                    throw new ArgumentNullException(nameof(mainBranch));
                }

                if (string.IsNullOrEmpty(checkingBranch) == true)
                {
                    throw new ArgumentNullException(nameof(checkingBranch));
                }

                string difference = this.GetDiffJson(projectId, mainBranch, checkingBranch);
                if (difference.StartsWith("{\"commit\":null,\"commits\":[],\"diffs\":[]", StringComparison.OrdinalIgnoreCase) == true)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (WebException)
            {
                return false;
            }
        }

        /// <summary>
        /// Get the single Merge request details 
        /// </summary>
        /// <param name="projectId">ID of the project</param>
        /// <param name="mergeRequestIid">IID of the merge request</param>
        /// <returns>Returns the single Merge request JSON</returns>
        protected string GetSingleMergeRequestJson(string projectId, string mergeRequestIid)
        {
            return this.HttpOperation("https://gitlab.syncfusion.com/api/v4/projects/" + projectId + "/merge_requests/" + mergeRequestIid + "?private_token=" + Token);
        }

        /// <summary>
        /// To Get the merge request details in given project
        /// </summary>
        /// <param name="projectId">ID of the project</param>
        /// <param name="state">state of the project which is to be checked</param>
        /// <returns>Returns the Given project`s merge request list which matches the given state </returns>
        protected List<MergeRequestDetails> GetMergeRequestDetails(string projectId, string state)
        {
            List<MergeRequestDetails> mergeRequestDetails = new List<MergeRequestDetails>();
            bool check = true;
            int pageCount = 1;
            while (check)
            {
                var jSONStringForId = this.HttpOperation("https://gitlab.syncfusion.com/api/v4/projects/" + projectId + "/merge_requests?" + state + "&per_page=99&page=" + pageCount + "&private_token=" + Token);
                dynamic jsonForId = JsonConvert.DeserializeObject(jSONStringForId);
                if (jSONStringForId != "[]")
                {
                    foreach (var item in jsonForId)
                    {
                        mergeRequestDetails.Add(new MergeRequestDetails { Iid = item.iid, State = item.state, TargetBranch = item.target_branch, Sha = item.sha, Title = item.title });
                    }
                }
                else
                {
                    check = false;
                }

                pageCount++;
            }

            return mergeRequestDetails;
        }

        /// <summary>
        /// To get the project members details in list form
        /// </summary>
        /// <param name="projectId">Id of project</param>
        /// <returns>Returns the User`s details list</returns>
        protected List<UserDetails> GetGroupMembersList(string projectId)
        {
            List<UserPermissions> userDetails = new List<UserPermissions>();
            int pageCount = 1;
            bool check = true;

            ////Get the project members 
            while (check)
            {
                var jSONStringForId = this.HttpOperation("https://gitlab.syncfusion.com/api/v4/groups/" + projectId + "/members?per_page=99&page=" + pageCount + "&private_token=" + Token);
                dynamic jsonForId = JsonConvert.DeserializeObject(jSONStringForId);
                if (jSONStringForId != "[]")
                {
                    foreach (var item in jsonForId)
                    {
                        userDetails.Add(new UserPermissions { UserId = item.id, AccessLevel = item.access_level });
                    }
                }
                else
                {
                    check = false;
                }

                pageCount++;
            }
            ////Get full details of the project members
            return this.GetUserDetails(userDetails);
        }

        /// <summary>
        /// To get the project members details in list form
        /// </summary>
        /// <param name="projectId">Id of project</param>
        /// <returns>Returns the User`s details list</returns>
        protected List<UserDetails> GetMembersList(string projectId)
        {
            List<UserPermissions> userDetails = new List<UserPermissions>();
            int pageCount = 1;
            bool check = true;

            ////Get the project members 
            while (check)
            {
                var jSONStringForId = this.HttpOperation("https://gitlab.syncfusion.com/api/v4/projects/" + projectId + "/members?per_page=99&page=" + pageCount + "&private_token=" + Token);
                dynamic jsonForId = JsonConvert.DeserializeObject(jSONStringForId);
                if (jSONStringForId != "[]")
                {
                    foreach (var item in jsonForId)
                    {
                        userDetails.Add(new UserPermissions { UserId = item.id, AccessLevel = item.access_level });
                    }
                }
                else
                {
                    check = false;
                }

                pageCount++;
            }
            ////Get full details of the project members
            return this.GetUserDetails(userDetails);
        }

        /// <summary>
        /// To get the full details of user by user ID
        /// </summary>
        /// <param name="userIdList">User ID List</param>
        /// <returns>Returns the list of given user ID`s details</returns>
        protected List<UserDetails> GetUserDetails(List<UserPermissions> userIdList)
        {
            List<UserDetails> userDetails = new List<UserDetails>();
            foreach (var userId in userIdList)
            {
                dynamic userDetail = JsonConvert.DeserializeObject(this.HttpOperation("https://gitlab.syncfusion.com/api/v4/users/" + userId.UserId + "?private_token=" + Token));
                userDetails.Add(new UserDetails { Id = userDetail.id, UserName = userDetail.username, Email = userDetail.email, State = userId.AccessLevel });
            }

            return userDetails;
        }

        /// <summary>
        /// Get the target branch`s commit details
        /// </summary>
        /// <param name="projectId">project`s project ID</param>
        /// <param name="sha">sha details in commit</param>
        /// <returns>returns the targets branch commit details in JSON</returns>
        protected string GetCommitDetails(string projectId, string sha)
        {
            return this.HttpOperation("https://gitlab.syncfusion.com/api/v4/projects/" + projectId + "/repository/commits/" + sha + "?private_token=" + Token);
        }

        /// <summary>
        /// Get the Diff JSON
        /// </summary>
        /// <param name="projectId">ID of the projecct</param>
        /// <param name="mainBranch">Destination branch</param>
        /// <param name="checkingBranch">Source branch</param>
        /// <returns>Returns the JSON of the two branch diff</returns>
        protected string GetDiffJson(string projectId, string mainBranch, string checkingBranch)
        {
            return this.HttpOperation("https://gitlab.syncfusion.com/api/v4/projects/" + projectId + "/repository/compare?from=" + mainBranch + "&to=" + checkingBranch + "&private_token=" + Token);
        }

        /// <summary>
        /// Changing the access level for a user in single project
        /// </summary>
        /// <param name="projectId">Id of the project where going to changing the access level</param>
        /// <param name="userId">Id of the user whose acccess level to be changed</param>
        /// <param name="accessLevel">Access level developer(30)/master(40)</param>
        /// <returns>Returns the bool result of permission changing</returns>
        protected bool ChangeAccessLevelInProject(string projectId, string userId, string accessLevel)
        {
            if (string.IsNullOrEmpty(projectId) == true)
            {
                throw new ArgumentNullException(nameof(projectId));
            }

            if (string.IsNullOrEmpty(accessLevel) == true)
            {
                throw new ArgumentNullException(nameof(accessLevel));
            }

            if (string.IsNullOrEmpty(userId) == true)
            {
                throw new ArgumentNullException(nameof(userId));
            }

            string addUserJson = string.Empty, deleteUserJson = string.Empty;
            addUserJson = this.AddUserInProject(projectId, userId, accessLevel);
            if (addUserJson.Contains("Member already exists"))
            {
                deleteUserJson = this.DeleteUserInProject(projectId, userId);
                if (string.IsNullOrEmpty(deleteUserJson))
                {
                    addUserJson = this.AddUserInProject(projectId, userId, accessLevel);
                    if (addUserJson.StartsWith("{\"name\":", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return true;
            }

        }

        /// <summary>
        /// Add a single user in specific project
        /// </summary>
        /// <param name="projectId">Id of the projecct</param>
        /// <param name="userId">Id of the user</param>
        /// <param name="accessLevel">Access level of the user</param>
        /// <returns>Returns the JSON of add user cURL</returns>
        protected string AddUserInProject(string projectId, string userId, string accessLevel)
        {
            if (string.IsNullOrEmpty(projectId) == true)
            {
                throw new ArgumentNullException(nameof(projectId));
            }

            if (string.IsNullOrEmpty(userId) == true)
            {
                throw new ArgumentNullException(nameof(userId));
            }

            if (string.IsNullOrEmpty(accessLevel) == true)
            {
                throw new ArgumentNullException(nameof(accessLevel));
            }

            return this.CurlOperation("--request POST --header \"PRIVATE-TOKEN:" + Token + "\" --data \"user_id=" + userId + "&access_level=" + accessLevel + "\" https://gitlab.syncfusion.com/api/v4/projects/" + projectId + "/members");
        }

        /// <summary>
        /// Delete the user in single project
        /// </summary>
        /// <param name="projectId">ID of the project where user have to remove</param>
        /// <param name="userId">ID of the user who have to remove from project</param>
        /// <returns>Returns the JSON of delete user cURL</returns>
        protected string DeleteUserInProject(string projectId, string userId)
        {
            if (string.IsNullOrEmpty(projectId) == true)
            {
                throw new ArgumentNullException(nameof(projectId));
            }

            if (string.IsNullOrEmpty(userId) == true)
            {
                throw new ArgumentNullException(nameof(userId));
            }

            return this.CurlOperation("--request DELETE --header \"PRIVATE-TOKEN:" + Token + "\" https://gitlab.syncfusion.com/api/v4/projects/" + projectId + "/members/" + userId);
        }

        /// <summary>
        /// Find the working days between two days
        /// </summary>
        /// <param name="startingDate">Starting date</param>
        /// <param name="endingDate">End date</param>
        /// <returns>Returns the number of working days</returns>
        protected int GetWeekdaysBetween(DateTime startingDate, DateTime endingDate)
        {
            if (startingDate > endingDate)
            {
                DateTime temp = startingDate;
                startingDate = endingDate;
                endingDate = temp;
            }

            DateTime increasedDate = startingDate.Date;
            int weekEndDays = 0;
            var totalDays = (endingDate.Date - startingDate.Date).Days;
            while (endingDate.Date != increasedDate)
            {
                if (increasedDate.DayOfWeek == DayOfWeek.Saturday || increasedDate.DayOfWeek == DayOfWeek.Sunday)
                {
                    weekEndDays++;
                }

                increasedDate = increasedDate.AddDays(1);
            }

            return totalDays - weekEndDays;
        }

        /// <summary>
        /// cURL operation 
        /// </summary>
        /// <param name="arguments">Arguments of the process</param>
        /// <returns>returns the JSON of given arguments</returns>
        protected string CurlOperation(string arguments)
        {
            if (string.IsNullOrEmpty(arguments) == true)
            {
                throw new ArgumentNullException(nameof(arguments));
            }

            Process commandProcess = new Process();
            try
            {
                commandProcess.StartInfo.UseShellExecute = false;
                if (File.Exists(@"C:\Program Files\cURL\bin\curl.exe"))
                {
                    commandProcess.StartInfo.FileName = @"C:\Program Files\cURL\bin\curl.exe"; //// this is the path of curl where it is installed;    
                }
                else if (File.Exists(@"C:\Program Files\Git\usr\bin\curl.exe"))
                {
                    commandProcess.StartInfo.FileName = @"C:\Program Files\Git\usr\bin\curl.exe"; //// this is the path of curl where it is installed;    
                }
                else
                {
                    commandProcess.StartInfo.FileName = @"C:\Program Files\Git\mingw64\bin\curl.exe"; //// this is the path of curl where it is installed;    
                }

                commandProcess.StartInfo.Arguments = arguments;
                commandProcess.StartInfo.CreateNoWindow = true;
                commandProcess.StartInfo.RedirectStandardInput = true;
                commandProcess.StartInfo.RedirectStandardOutput = true;
                commandProcess.StartInfo.RedirectStandardError = true;
                commandProcess.Start();
                commandProcess.WaitForExit();
                string output = commandProcess.StandardOutput.ReadToEnd();
                return output;
            }
            finally
            {
                if (commandProcess != null)
                {
                    commandProcess.Dispose();
                }
            }
        }

        /// <summary>
        /// HTTP Operation for API URI
        /// </summary>
        /// <param name="apiUrl">Web URL for get the JSON Data</param>
        /// <returns>Returns the JSON as string</returns>
        protected string HttpOperation(string apiUrl)
        {
            if (string.IsNullOrEmpty(apiUrl) == true)
            {
                throw new ArgumentNullException(nameof(apiUrl));
            }

            Uri url = new Uri(apiUrl);
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            HttpWebResponse response = request.GetResponse() as HttpWebResponse;
            using (Stream responseStream = response.GetResponseStream())
            {
                StreamReader reader = new StreamReader(responseStream, Encoding.UTF8);
                return reader.ReadToEnd();
            }
        }

        /// <summary>
        /// This class is for returning the status of the operation 
        /// </summary>
        public class OperationResult
        {
            /// <summary>
            /// Gets or sets the project name
            /// </summary>
            public string ProjectName { get; set; }

            /// <summary>
            /// Gets or sets the Comments
            /// </summary>
            public string Comments { get; set; }

            /// <summary>
            /// Gets or sets a value indicating whether the operation is successfull or not
            /// </summary>
            public bool Status { get; set; }
        }

        /// <summary>
        /// Merge request details 
        /// </summary>
        public class MergeRequestData
        {
            /// <summary>
            /// Gets or sets the Project name 
            /// </summary>
            public string ProjectName { get; set; }

            /// <summary>
            /// Gets or sets the status 
            /// </summary>
            public string Status { get; set; }

            /// <summary>
            /// Gets or sets the merge IID 
            /// </summary>
            public string MergeIid { get; set; }

            /// <summary>
            /// Gets or sets the commit ID 
            /// </summary>
            public string CommitId { get; set; }
        }

        /// <summary>
        /// This class is for storing the project user details 
        /// </summary>
        public class UserPermissions
        {
            /// <summary>
            /// Gets or sets the project name
            /// </summary>
            public string UserId { get; set; }

            /// <summary>
            /// Gets or sets a value indicating whether the operation is successfull or not
            /// </summary>
            public string AccessLevel { get; set; }
        }

        /// <summary>
        /// Merge request details 
        /// </summary>
        protected class MergeRequestDetails
        {
            /// <summary>
            /// Gets or sets the iid 
            /// </summary>
            public string Iid { get; set; }

            /// <summary>
            /// Gets or sets the state 
            /// </summary>
            public string State { get; set; }

            /// <summary>
            /// Gets or sets the target branch 
            /// </summary>
            public string TargetBranch { get; set; }

            /// <summary>
            /// Gets or sets the sha ID 
            /// </summary>
            public string Sha { get; set; }

            /// <summary>
            /// Gets or sets the title 
            /// </summary>
            public string Title { get; set; }
        }

        protected class ExcelOperations
        {
            /// <summary>
            /// Generating the excel sheet
            /// </summary>
            /// <param name="finalExcelData"> the data to be filled in excel data</param>
            /// <param name="projectUrl">The URL of the project</param>
            /// <param name="tagsOrBranchUrl">The URL of the tags o branches</param>
            /// <param name="operationType">describe the operation</param>
            public void GenerateExcelSheet(List<OperationResult> finalExcelData, List<string> projectUrl, List<string> tagsOrBranchUrl, string operationType)
            {
                if (projectUrl == null)
                {
                    throw new ArgumentNullException(nameof(projectUrl));
                }

                if (tagsOrBranchUrl == null)
                {
                    throw new ArgumentNullException(nameof(tagsOrBranchUrl));
                }

                if (finalExcelData == null)
                {
                    throw new ArgumentNullException(nameof(finalExcelData));
                }

                Collection<string> headingforExcel = new Collection<string>();
                switch (operationType)
                {
                    case "Merge Protect":
                        headingforExcel.Add("Project name");
                        headingforExcel.Add("Comments");
                        headingforExcel.Add("Status");
                        string errorMessageforBranch = "Projects were not frozen";
                        this.GenerateExcel(finalExcelData, projectUrl, tagsOrBranchUrl, 3, headingforExcel, errorMessageforBranch, "Project frozen");
                        break;
                }
            }

            /// <summary>
            /// Generating the Excel document
            /// </summary>
            /// <param name="finalExcelData">The date to be filled in excel</param>
            /// <param name="projectUrl"> the URL of the projects</param>
            /// <param name="tagsOrBranchUrl">The Url of the tags or branches</param>
            /// <param name="columnCounts">Count of the column for style</param>
            /// <param name="heading">Heading of the excel sheet</param>
            /// <param name="errorMessage">Error message for fauilure cases</param>
            /// <param name="operation">defines the kind of operation</param>
            public void GenerateExcel(List<OperationResult> finalExcelData, List<string> projectUrl, List<string> tagsOrBranchUrl, int columnCounts, Collection<string> heading, string errorMessage, string operation)
            {
                foreach (var a in finalExcelData)
                {
                    Console.WriteLine(a.ProjectName + "*****" + a.Status + "*****" + a.Comments + "*****");
                }
                if (projectUrl == null)
                {
                    throw new ArgumentNullException(nameof(projectUrl));
                }

                if (tagsOrBranchUrl == null)
                {
                    throw new ArgumentNullException(nameof(tagsOrBranchUrl));
                }

                if (finalExcelData == null)
                {
                    throw new ArgumentNullException(nameof(finalExcelData));
                }

                if (heading == null)
                {
                    throw new ArgumentNullException(nameof(heading));
                }

                if (operation == null)
                {
                    throw new ArgumentNullException(nameof(heading));
                }

                ExcelEngine excelEngine = new ExcelEngine();
                try
                {
                    ////deleting the existing excel file
                    if (File.Exists(@"GitLab_Log_file.xlsx"))
                    {
                        File.Delete(@"GitLab_Log_file.xlsx");
                    }

                    IApplication application = excelEngine.Excel;
                    int columnCount = columnCounts;
                    int rowCount = finalExcelData.Count;
                    application.DefaultVersion = ExcelVersion.Excel2013;
                    IWorkbook workbook = application.Workbooks.Create(1);
                    IWorksheet workSheet = workbook.Worksheets[0];
                    if (finalExcelData.Count == 0)
                    {
                        workSheet.Range["A1"].Text = errorMessage;
                    }
                    else
                    {
                        int count = 0;
                        for (char alphbets = 'A'; alphbets <= 'Z'; alphbets++)
                        {
                            if (count < columnCount)
                            {
                                workSheet.Range[alphbets + "1"].Text = heading[count];
                                count++;
                            }
                            else
                            {
                                break;
                            }
                        }

                        if (operation == "Project frozen")
                        {
                            workSheet.ImportData(finalExcelData, 2, 1, false);
                        }
                    }
                    ////Add Project URL
                    if (projectUrl != null)
                    {
                        for (int projectUrlCount = 0; projectUrlCount < projectUrl.Count; projectUrlCount++)
                        {
                            if (projectUrl[projectUrlCount] != "empty")
                            {
                                IHyperLink hyperlink = workSheet.HyperLinks.Add(workSheet.Range["A" + (projectUrlCount + 2)]);
                                hyperlink.Type = ExcelHyperLinkType.Url;
                                hyperlink.Address = projectUrl[projectUrlCount];
                            }
                        }
                    }

                    ////Add tags or branch URL
                    if (tagsOrBranchUrl != null)
                    {
                        if (operation.Equals("Branch Protection"))
                        {
                            for (int tagsOrBranchUrlCount = 0; tagsOrBranchUrlCount < tagsOrBranchUrl.Count; tagsOrBranchUrlCount++)
                            {
                                if (tagsOrBranchUrl[tagsOrBranchUrlCount] != "empty")
                                {
                                    IHyperLink hyperlink = workSheet.HyperLinks.Add(workSheet.Range["B" + (tagsOrBranchUrlCount + 2)]);
                                    hyperlink.Type = ExcelHyperLinkType.Url;
                                    hyperlink.Address = tagsOrBranchUrl[tagsOrBranchUrlCount];
                                }
                            }
                        }
                        else
                        {
                            for (int tagsOrBranchUrlCount = 0; tagsOrBranchUrlCount < tagsOrBranchUrl.Count; tagsOrBranchUrlCount++)
                            {
                                if (tagsOrBranchUrl[tagsOrBranchUrlCount] != "empty")
                                {
                                    IHyperLink hyperlink = workSheet.HyperLinks.Add(workSheet.Range["C" + (tagsOrBranchUrlCount + 2)]);
                                    hyperlink.Type = ExcelHyperLinkType.Url;
                                    hyperlink.Address = tagsOrBranchUrl[tagsOrBranchUrlCount];
                                }
                            }
                        }
                    }

                    ////Excel document style
                    IRange xlRangeHeader = workSheet.Range[1, 1, 1, columnCount];
                    IRange xlRangeResult = workSheet.Range[2, 1, rowCount + 1, columnCount];
                    workSheet[1, 1, rowCount + 1, columnCount].Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
                    workSheet[1, 1, rowCount + 1, columnCount].BorderInside(ExcelLineStyle.Thin);
                    xlRangeHeader.BorderAround();
                    if (columnCount > 1)
                    {
                        xlRangeHeader.BorderInside();
                    }

                    xlRangeHeader.CellStyle.Font.Bold = true;
                    workSheet.UsedRange.AutofitColumns();
                    xlRangeHeader.CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                    xlRangeResult.CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;
                    xlRangeResult.BorderAround(ExcelLineStyle.Thin, ExcelKnownColors.Black);
                    workSheet.Range[2, 1].FreezePanes();
                    workbook.SaveAs("GitLab_Log_file.xlsx");
                    workbook.Close();
                }
                finally
                {
                    ////excelEngine.Dispose();
                    if (excelEngine != null)
                    {
                        excelEngine.Dispose();
                    }
                }
            }
        }

        /// <summary>
        /// Excel operation
        /// </summary>
        protected class ExcelOperation
        {
            /// <summary>
            /// Generating the excel sheet
            /// </summary>
            /// <param name="finalExcelData"> the data to be filled in excel data</param>
            /// <param name="projectUrl">The URL of the project</param>
            /// <param name="tagsOrBranchUrl">The URL of the tags o branches</param>
            /// <param name="operationType">describe the operation</param>
            public void GenerateExcelSheet(Collection<ProjectData> finalExcelData, Collection<string> projectUrl, Collection<string> tagsOrBranchUrl, string operationType)
            {
                if (projectUrl == null)
                {
                    throw new ArgumentNullException(nameof(projectUrl));
                }

                if (tagsOrBranchUrl == null)
                {
                    throw new ArgumentNullException(nameof(tagsOrBranchUrl));
                }

                if (finalExcelData == null)
                {
                    throw new ArgumentNullException(nameof(finalExcelData));
                }

                Collection<string> headingforExcel = new Collection<string>();
                switch (operationType)
                {
                    case "Branch Creation":
                        headingforExcel.Add("Project name");
                        headingforExcel.Add("Source branch");
                        headingforExcel.Add("Destination branch");
                        headingforExcel.Add("Protect");
                        headingforExcel.Add("Status");
                        string errorMessageforBranch = "Branches were not created";
                        this.GenerateExcel(finalExcelData, projectUrl, tagsOrBranchUrl, 5, headingforExcel, errorMessageforBranch, "Branch Creation");
                        break;
                    case "Tag Creation":
                        headingforExcel.Add("Project name");
                        headingforExcel.Add("Source branch");
                        headingforExcel.Add("Destination tag");
                        headingforExcel.Add("Status");
                        string errorMessageforTags = "Tags were not created";
                        this.GenerateExcel(finalExcelData, projectUrl, tagsOrBranchUrl, 4, headingforExcel, errorMessageforTags, "Tag Creation");
                        break;
                    case "Tag Deletion":
                        headingforExcel.Add("Project name");
                        headingforExcel.Add("Tag name");
                        headingforExcel.Add("Status");
                        string errorMessageforTag = "Tags were not deleted";
                        this.GenerateExcel(finalExcelData, projectUrl, tagsOrBranchUrl, 3, headingforExcel, errorMessageforTag, "Tag Deletion");
                        break;
                    case "Branch Deletion":
                        headingforExcel.Add("Project name");
                        headingforExcel.Add("Branch name");
                        headingforExcel.Add("Status");
                        string errorMessageforBranchDeletion = "Branches were not deleted";
                        this.GenerateExcel(finalExcelData, projectUrl, tagsOrBranchUrl, 3, headingforExcel, errorMessageforBranchDeletion, "Tag Deletion");
                        break;
                    case "Branch Protection":
                        headingforExcel.Add("Project name");
                        headingforExcel.Add("Branch name");
                        headingforExcel.Add("Status");
                        string errorMessageforBranchProtection = "Branches were not protected";
                        this.GenerateExcel(finalExcelData, projectUrl, tagsOrBranchUrl, 3, headingforExcel, errorMessageforBranchProtection, "Branch Protection");
                        break;
                }
            }

            /// <summary>
            /// Generating the Excel document
            /// </summary>
            /// <param name="finalExcelData">The date to be filled in excel</param>
            /// <param name="projectUrl"> the URL of the projects</param>
            /// <param name="tagsOrBranchUrl">The Url of the tags or branches</param>
            /// <param name="columnCounts">Count of the column for style</param>
            /// <param name="heading">Heading of the excel sheet</param>
            /// <param name="errorMessage">Error message for fauilure cases</param>
            /// <param name="operation">defines the kind of operation</param>
            public void GenerateExcel(Collection<ProjectData> finalExcelData, Collection<string> projectUrl, Collection<string> tagsOrBranchUrl, int columnCounts, Collection<string> heading, string errorMessage, string operation)
            {
                if (projectUrl == null)
                {
                    throw new ArgumentNullException(nameof(projectUrl));
                }

                if (tagsOrBranchUrl == null)
                {
                    throw new ArgumentNullException(nameof(tagsOrBranchUrl));
                }

                if (finalExcelData == null)
                {
                    throw new ArgumentNullException(nameof(finalExcelData));
                }

                if (heading == null)
                {
                    throw new ArgumentNullException(nameof(heading));
                }

                if (operation == null)
                {
                    throw new ArgumentNullException(nameof(heading));
                }

                ExcelEngine excelEngine = new ExcelEngine();
                try
                {
                    ////deleting the existing excel file
                    if (File.Exists(@"GitLab_Log_file.xlsx"))
                    {
                        File.Delete(@"GitLab_Log_file.xlsx");
                    }

                    IApplication application = excelEngine.Excel;
                    int columnCount = columnCounts;
                    int rowCount = finalExcelData.Count;
                    application.DefaultVersion = ExcelVersion.Excel2013;
                    IWorkbook workbook = application.Workbooks.Create(1);
                    IWorksheet workSheet = workbook.Worksheets[0];
                    if (finalExcelData.Count == 0)
                    {
                        workSheet.Range["A1"].Text = errorMessage;
                    }
                    else
                    {
                        int count = 0;
                        for (char alphbets = 'A'; alphbets <= 'Z'; alphbets++)
                        {
                            if (count < columnCount)
                            {
                                workSheet.Range[alphbets + "1"].Text = heading[count];
                                count++;
                            }
                            else
                            {
                                break;
                            }
                        }

                        if (operation == "Branch Creation" || operation == "Tag Creation")
                        {
                            workSheet.ImportData(finalExcelData, 2, 1, false);
                        }
                        else if (operation == "Tag Deletion" || operation == "Branch Deletion")
                        {
                            var reducedList = finalExcelData.Select(data => new { data.Project_Name, data.Source_Branch, data.Status }).ToList();
                            workSheet.ImportData(reducedList, 2, 1, false);
                        }
                        else if (operation == "Branch Protection")
                        {
                            var reducedList = finalExcelData.Select(data => new { data.Project_Name, data.Source_Branch, data.Protect }).ToList();
                            workSheet.ImportData(reducedList, 2, 1, false);
                        }
                    }
                    ////Add Project URL
                    if (projectUrl != null)
                    {
                        for (int projectUrlCount = 0; projectUrlCount < projectUrl.Count; projectUrlCount++)
                        {
                            if (projectUrl[projectUrlCount] != "empty")
                            {
                                IHyperLink hyperlink = workSheet.HyperLinks.Add(workSheet.Range["A" + (projectUrlCount + 2)]);
                                hyperlink.Type = ExcelHyperLinkType.Url;
                                hyperlink.Address = projectUrl[projectUrlCount];
                            }
                        }
                    }

                    ////Add tags or branch URL
                    if (tagsOrBranchUrl != null)
                    {
                        if (operation.Equals("Branch Protection"))
                        {
                            for (int tagsOrBranchUrlCount = 0; tagsOrBranchUrlCount < tagsOrBranchUrl.Count; tagsOrBranchUrlCount++)
                            {
                                if (tagsOrBranchUrl[tagsOrBranchUrlCount] != "empty")
                                {
                                    IHyperLink hyperlink = workSheet.HyperLinks.Add(workSheet.Range["B" + (tagsOrBranchUrlCount + 2)]);
                                    hyperlink.Type = ExcelHyperLinkType.Url;
                                    hyperlink.Address = tagsOrBranchUrl[tagsOrBranchUrlCount];
                                }
                            }
                        }
                        else
                        {
                            for (int tagsOrBranchUrlCount = 0; tagsOrBranchUrlCount < tagsOrBranchUrl.Count; tagsOrBranchUrlCount++)
                            {
                                if (tagsOrBranchUrl[tagsOrBranchUrlCount] != "empty")
                                {
                                    IHyperLink hyperlink = workSheet.HyperLinks.Add(workSheet.Range["C" + (tagsOrBranchUrlCount + 2)]);
                                    hyperlink.Type = ExcelHyperLinkType.Url;
                                    hyperlink.Address = tagsOrBranchUrl[tagsOrBranchUrlCount];
                                }
                            }
                        }
                    }

                    ////Excel document style
                    IRange xlRangeHeader = workSheet.Range[1, 1, 1, columnCount];
                    IRange xlRangeResult = workSheet.Range[2, 1, rowCount + 1, columnCount];
                    workSheet[1, 1, rowCount + 1, columnCount].Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
                    workSheet[1, 1, rowCount + 1, columnCount].BorderInside(ExcelLineStyle.Thin);
                    xlRangeHeader.BorderAround();
                    if (columnCount > 1)
                    {
                        xlRangeHeader.BorderInside();
                    }

                    xlRangeHeader.CellStyle.Font.Bold = true;
                    workSheet.UsedRange.AutofitColumns();
                    xlRangeHeader.CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                    xlRangeResult.CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;
                    xlRangeResult.BorderAround(ExcelLineStyle.Thin, ExcelKnownColors.Black);
                    workSheet.Range[2, 1].FreezePanes();
                    workbook.SaveAs("GitLab_Log_file.xlsx");
                    workbook.Close();
                }
                finally
                {
                    ////excelEngine.Dispose();
                    if (excelEngine != null)
                    {
                        excelEngine.Dispose();
                    }
                }
            }
        }

        /// <summary>
        /// Mail notificatio class
        /// </summary>
        protected class MailNotification
        {
            /// <summary>
            /// The mailing part of the service
            /// </summary>
            /// <param name="subject">subject of the mail</param>
            /// <param name="body">body of the mail</param>
            /// <param name="attachmentFileName">attachment of the mail</param>
            /// <param name="emailId">To address for mail</param>
            /// <returns>Return the bool value of mail delivery</returns>
            public bool MailNotifications(string subject, string body, string attachmentFileName, string emailId)
            {
                if (string.IsNullOrEmpty(subject) == true)
                {
                    throw new ArgumentNullException(nameof(subject));
                }

                //if (string.IsNullOrEmpty(body) == true)
                //{
                //    throw new ArgumentNullException(nameof(body));
                //}

                if (string.IsNullOrEmpty(emailId) == true)
                {
                    throw new ArgumentNullException(nameof(emailId));
                }

                this.Equals(0);
                Collection<string> toAddress = new Collection<string>();
                string[] mailId = emailId.Split(',');
                foreach (string mail in mailId)
                {
                    toAddress.Add(mail);
                }

                Collection<string> ccAddress = new Collection<string>();
                //ccAddress.Add("gitlabteam@syncfusion.com ");
                Collection<string> bccAddress = new Collection<string>();
                string fromAddress = "gitlab@syncfusion.com";
                bool mailResult = MailService.SendMessage(toAddress, ccAddress, bccAddress, fromAddress, subject, body, attachmentFileName);
                return mailResult;
            }

            /// <summary>
            /// Mail service
            /// </summary>
            protected static class MailService
            {
                /// <summary>
                /// Mail details 
                /// </summary>
                /// <param name="emailAddress">TO mail ID</param>
                /// <param name="ccEmailAddress">To mail ID</param>
                /// <param name="bccEmailAddress">BCC mail ID</param>
                /// <param name="fromAddress">from mail ID</param>
                /// <param name="subject">Subject of the mail</param>
                /// <param name="body">Body of the mail</param>
                /// <param name="attachmentFileName">Attachment of the mail</param>
                /// <returns>return the bool result of mail sending</returns>
                public static bool SendMessage(Collection<string> emailAddress, Collection<string> ccEmailAddress, Collection<string> bccEmailAddress, string fromAddress, string subject, string body, string attachmentFileName)
                {
                    EmailSenderThread est = new EmailSenderThread();
                    try
                    {
                        est.EmailSenderThreadD(emailAddress, ccEmailAddress, bccEmailAddress, fromAddress, subject, body, attachmentFileName);
                    }
                    finally
                    {
                        if (est != null)
                        {
                            est.Dispose();
                        }
                    }

                    return true;
                }
            }

            /// <summary>
            /// Convert into mail message
            /// </summary>
            protected static class ConvertMailMessage
            {
                /// <summary>
                /// Converting the mail message to memory stream
                /// </summary>
                /// <param name="message"> mail message</param>
                /// <returns>returns the mail stream</returns>
                public static MemoryStream ConvertMailMessageToMemoryStream(MailMessage message)
                {
                    MemoryStream fileStream = new MemoryStream();
                    try
                    {
                        Assembly assembly = typeof(SmtpClient).Assembly;
                        Type mailWriterType = assembly.GetType("System.Net.Mail.MailWriter");
                        ConstructorInfo mailWriterContructor = mailWriterType.GetConstructor(BindingFlags.Instance | BindingFlags.NonPublic, null, new[] { typeof(Stream) }, null);
                        object mailWriter = mailWriterContructor.Invoke(new object[] { fileStream });
                        MethodInfo sendMethod = typeof(MailMessage).GetMethod("Send", BindingFlags.Instance | BindingFlags.NonPublic);
                        sendMethod.Invoke(message, BindingFlags.Instance | BindingFlags.NonPublic, null, new[] { mailWriter, true, true }, null); // have to set additional parameter when manually run the tool                
                        MethodInfo closeMethod = mailWriter.GetType().GetMethod("Close", BindingFlags.Instance | BindingFlags.NonPublic);
                        closeMethod.Invoke(mailWriter, BindingFlags.Instance | BindingFlags.NonPublic, null, new object[] { }, null);
                        return fileStream;
                    }
                    finally
                    {
                        if (fileStream != null)
                        {
                            fileStream.Close();
                        }
                    }
                }
            }

            /// <summary>
            /// Email configuration
            /// </summary>
            protected class EmailConfig : IDisposable
            {
                /// <summary>
                /// Set the disposed value
                /// </summary>
                private bool disposedValue = false; // To detect redundant calls

                /// <summary>
                /// Amazon user name
                /// </summary>
                private string amazonUserName = "AKIAJMNE62WIFKEIFNVQ";

                /// <summary>
                /// Amazon password
                /// </summary>
                private string amazonPassword = "wylPqE2YWr11cAZgbeY/bOVsKgVRoH7PP3qFieDC";

                /// <summary>
                /// Initializes a new instance of the <see cref="EmailConfig"/> class.
                /// </summary>
                public EmailConfig()
                {
                    this.AmazonClient = new AmazonSimpleEmailServiceClient(this.amazonUserName, this.amazonPassword);
                }

                /// <summary>
                /// Gets the mail TO list
                /// </summary>
                public ICollection<string> MailNotifications { get; } = new List<string>();

                /// <summary>
                /// Gets or sets Mail message 
                /// </summary>
                protected MailMessage MailMessage { get; set; } = new MailMessage();

                /// <summary>
                /// Gets or sets the raw message content
                /// </summary>
                protected RawMessage RawMessage { get; set; } = new RawMessage();

                /// <summary>
                /// Gets or sets mail request for mail send
                /// </summary>
                protected SendRawEmailRequest Request { get; set; } = new SendRawEmailRequest();

                /// <summary>
                /// Gets the mail CC list
                /// </summary>
                protected ICollection<string> AdditionalNotifications { get; } = new List<string>();

                /// <summary>
                /// Gets  the mail BCC list
                /// </summary>
                protected ICollection<string> AdditionalNotificationsBcc { get; } = new List<string>();

                /// <summary>
                /// Gets or sets the Amazon mail service
                /// </summary>
                protected AmazonSimpleEmailServiceClient AmazonClient { get; set; }

                #region IDisposable Support

                /// <summary>
                /// Disposable method
                /// </summary>
                public void Dispose()
                {
                    // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
                    this.Dispose(true);
                    //// TODO: uncomment the following line if the finalizer is overridden above.
                    GC.SuppressFinalize(this);
                }

                /// <summary>
                ///  dispose method
                /// </summary>
                /// <param name="disposing"> passing the disposing parameter</param>
                protected virtual void Dispose(bool disposing)
                {
                    if (!this.disposedValue)
                    {
                        this.disposedValue = true;

                        if (disposing)
                        {
                            this.Dispose(true);
                            GC.SuppressFinalize(this);
                            //// TODO: dispose managed state (managed objects).
                        }

                        //// TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                        //// TODO: set large fields to null.
                    }
                }

                //// TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
                //// This code added to correctly implement the disposable pattern.

                #endregion
            }

            /// <summary>
            /// Email sender thread
            /// </summary>
            protected class EmailSenderThread : EmailConfig
            {
                /// <summary>
                /// thread for the mail message
                /// </summary>
                private Thread msgThread;

                /// <summary>
                /// Initializes a new instance of the <see cref="EmailSenderThreadD"/> class.
                /// </summary>
                /// <param name="emailAddresses">To address</param>
                /// <param name="ccEmailAddress">CC address</param>
                /// <param name="bccEmailAddress">BCC address</param>
                /// <param name="from">from address</param>
                /// <param name="subject">subject of the mail</param>
                /// <param name="body">Body of the mail</param>
                /// <param name="attachmentFileName">Attachment of the mail</param>
                public void EmailSenderThreadD(Collection<string> emailAddresses, Collection<string> ccEmailAddress, Collection<string> bccEmailAddress, string from, string subject, string body, string attachmentFileName)
                {
                    AlternateView htmlView = AlternateView.CreateAlternateViewFromString(body, Encoding.UTF8, "text/html");
                    try
                    {
                        this.msgThread = new Thread(new ThreadStart(this.MailSender));
                        MailMessage.From = new MailAddress(string.IsNullOrEmpty(from) ? "gitlabteam@syncfusion.com" : from);
                        if (emailAddresses != null)
                        {
                            var tomails = emailAddresses;
                            foreach (string tomail in tomails)
                            {
                                if (!string.IsNullOrEmpty(tomail))
                                {
                                    MailMessage.To.Add(new MailAddress(tomail));
                                    MailNotifications.Add(tomail);
                                }
                            }
                        }

                        if (ccEmailAddress != null)
                        {
                            var ccemails = ccEmailAddress;
                            foreach (string ccmail in ccemails)
                            {
                                if (!string.IsNullOrEmpty(ccmail))
                                {
                                    MailMessage.CC.Add(new MailAddress(ccmail));
                                    AdditionalNotifications.Add(ccmail);
                                }
                            }
                        }

                        if (bccEmailAddress != null)
                        {
                            var bccemails = bccEmailAddress;
                            foreach (string bccmail in bccemails)
                            {
                                if (!string.IsNullOrEmpty(bccmail))
                                {
                                    MailMessage.Bcc.Add(new MailAddress(bccmail));
                                    AdditionalNotificationsBcc.Add(bccmail);
                                }
                            }
                        }

                        MailMessage.Subject = subject;
                        if (body != null)
                        {
                            MailMessage.AlternateViews.Add(htmlView);
                        }

                        if (!string.IsNullOrEmpty(attachmentFileName))
                        {
                            var attachment = new Attachment(attachmentFileName);
                            MailMessage.Attachments.Add(attachment);
                        }

                        MemoryStream memoryStream = ConvertMailMessage.ConvertMailMessageToMemoryStream(MailMessage);
                        RawMessage.WithData(memoryStream);
                        Request.WithRawMessage(this.RawMessage);
                        Request.WithDestinations(this.MailNotifications);
                        Request.WithDestinations(this.AdditionalNotifications);
                        Request.WithDestinations(this.AdditionalNotificationsBcc);
                        Request.WithSource(from);
                        this.msgThread.Start();
                    }
                    catch (IOException ex)
                    {
                        Console.WriteLine(ex);
                    }
                    finally
                    {
                        if (htmlView != null)
                        {
                            {
                                htmlView.Dispose();
                            }
                        }
                    }
                }

                /// <summary>
                /// Mail sender
                /// </summary>
                public void MailSender()
                {
                    try
                    {
                        AmazonClient.SendRawEmail(this.Request);
                    }
                    catch (FileNotFoundException ex)
                    {
                        Console.WriteLine(ex);
                    }
                }
            }
        }

        /// <summary>
        /// Project data which is used to get the data from JSON
        /// </summary>
        protected class ProjectDetails
        {
            /// <summary>
            /// Gets or sets the project name as per GitLab
            /// </summary>
            public string Name { get; set; }

            /// <summary>
            /// Gets or sets the ID of the project
            /// </summary>
            public string Id { get; set; }

            /// <summary>
            /// Gets or sets the Web URL of the project
            /// </summary>
            public string Web_Url { get; set; }

            /// <summary>
            /// Gets or sets the HTTP URL of the project
            /// </summary>
            public string Http_Url_To_Repo { get; set; }

            /// <summary>
            /// Gets or sets the path of project with group
            /// </summary>
            public string Path_With_Namespace { get; set; }
        }

        /// <summary>
        /// Project details which is using for generate the excel document data
        /// </summary>
        protected class ProjectData
        {
            /// <summary>
            /// Gets or sets the Project name 
            /// </summary>
            public string Project_Name { get; set; }

            /// <summary>
            /// Gets or sets theSource branch name
            /// </summary>
            public string Source_Branch { get; set; }

            /// <summary>
            /// Gets or sets the destination branch name
            /// </summary>
            public string Destination_Branch { get; set; }

            /// <summary>
            /// Gets or sets the status 
            /// </summary>
            public string Status { get; set; }

            /// <summary>
            /// Gets or sets the Protect
            /// </summary>
            public string Protect { get; set; }
        }

        /// <summary>
        /// User details 
        /// </summary>
        public class UserDetails
        {
            /// <summary>
            /// Gets or sets the id 
            /// </summary>
            public string Id { get; set; }

            /// <summary>
            /// Gets or sets the username 
            /// </summary>
            public string UserName { get; set; }

            /// <summary>
            /// Gets or sets the email 
            /// </summary>
            public string Email { get; set; }

            /// <summary>
            /// Gets or sets the state 
            /// </summary>
            public string State { get; set; }
        }

        /// <summary>
        /// Branch details
        /// </summary>
        protected class BranchDetails
        {
            /// <summary>
            /// Gets or sets the Project name 
            /// </summary>
            public string Name { get; set; }

            /// <summary>
            /// Gets or sets a value indicating whether the merged state
            /// </summary>
            public bool Merged { get; set; }

            /// <summary>
            /// Gets or sets a value indicating whether the Project name
            /// </summary>
            public bool Protected { get; set; }

            /// <summary>
            /// Gets or sets the commit date 
            /// </summary>
            public CommitData Commit { get; set; }
        }

        /// <summary>
        /// To get the created date in JSON
        /// </summary>
        protected class CommitData
        {
            /// <summary>
            /// Gets or sets the created date
            /// </summary>
            [JsonProperty(PropertyName = "authored_date")]
            public DateTime AuthoredDate { get; set; }
        }

        /// <summary>
        /// Group details
        /// </summary>
        protected class GroupDetails
        {
            /// <summary>
            /// Gets or sets the id 
            /// </summary>
            public string Id { get; set; }

            /// <summary>
            /// Gets or sets the name 
            /// </summary>
            public string Name { get; set; }

            /// <summary>
            /// Gets or sets the visibility 
            /// </summary>
            public string Visibility { get; set; }

            /// <summary>
            /// Gets or sets the web url 
            /// </summary>
            public string Web_Url { get; set; }
        }
    }
}
