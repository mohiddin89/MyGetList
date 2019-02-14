using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.SharePoint;
using System.IO;

namespace GetListUK
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            List<string> listColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(txtInput.Text));
            while (!sr.EndOfStream)
            {
                try
                {
                    listColl.Add(sr.ReadLine().Trim());
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            StreamWriter ExcelwriterScoringMatrixNew = null;
            ExcelwriterScoringMatrixNew = System.IO.File.CreateText(txtReport.Text + "\\" + "ScoringMatrix" + "_" + DateTime.Now.ToString("dd-mm-yyyy-hh-mm-ss") + ".csv");
            ExcelwriterScoringMatrixNew.WriteLine("SiteURL" +","+ "PageName" +","+ "PageUrl" +","+ "PageId" +","+ "ListType" +","+ "Type" +","+ "webpart" +","+ "url");
            ExcelwriterScoringMatrixNew.Flush();

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        ListCollection _Lists = clientcontext.Web.Lists;
                        clientcontext.Load(_Lists);
                        clientcontext.ExecuteQuery();

                        foreach (List list in _Lists)
                        {
                            clientcontext.Load(list);
                            clientcontext.ExecuteQuery();

                            string listName = list.Title;

                            try
                            {
                                //bool _dListExist = clientcontext.Web.Lists.Cast<List>().Any(xlist => string.Equals(xlist.Title, listName));

                                //if (_dListExist)
                                {
                                    if (listName == "Status")
                                    {
                                        // try
                                        {
                                            List Pagelist = clientcontext.Web.Lists.GetByTitle(listName);
                                            clientcontext.Load(Pagelist);
                                            clientcontext.ExecuteQuery();

                                            ViewCollection ViewColl = Pagelist.Views;
                                            clientcontext.Load(ViewColl);
                                            clientcontext.ExecuteQuery();

                                            Microsoft.SharePoint.Client.View v = ViewColl[0];
                                            clientcontext.Load(v);
                                            clientcontext.ExecuteQuery();

                                            //v.DeleteObject();
                                            //clientcontext.ExecuteQuery();

                                            v.ViewFields.RemoveAll();
                                            v.Update();
                                            clientcontext.ExecuteQuery();

                                            v.ViewFields.Add("StatusDescription");
                                            v.Update();
                                            clientcontext.ExecuteQuery();
                                        }
                                    }

                                    if ((list.BaseTemplate.ToString() == "109") && (listName != "Photos" || listName == "Images"))
                                    {
                                        #region Commented

                                        FieldCollection FldColl = list.Fields;
                                        clientcontext.Load(FldColl);
                                        clientcontext.ExecuteQuery();
                                        bool TagCateExist = false;

                                        foreach (Field tagField in FldColl)
                                        {
                                            clientcontext.Load(tagField);
                                            clientcontext.ExecuteQuery();

                                            if (tagField.Title.ToLower() == "tags" || tagField.Title.ToLower() == "categorization")
                                            {
                                                TagCateExist = true;
                                                break;
                                            }
                                        }

                                        #endregion

                                        if (TagCateExist)
                                        {
                                            List Pagelist = clientcontext.Web.Lists.GetByTitle(listName);
                                            clientcontext.Load(Pagelist);
                                            clientcontext.ExecuteQuery();

                                            ViewCollection ViewColl = Pagelist.Views;
                                            clientcontext.Load(ViewColl);
                                            clientcontext.ExecuteQuery();

                                            Microsoft.SharePoint.Client.View v = ViewColl[0];
                                            clientcontext.Load(v);
                                            clientcontext.ExecuteQuery();

                                            v.ViewFields.RemoveAll();
                                            v.Update();
                                            clientcontext.ExecuteQuery();

                                            v.ViewFields.Add("DocIcon");
                                            v.ViewFields.Add("Title");
                                            v.ViewFields.Add("LinkFilename");
                                            v.ViewFields.Add("Created");
                                            v.ViewFields.Add("Created By");
                                            v.ViewFields.Add("Modified");
                                            v.ViewFields.Add("Modified By");
                                            v.ViewFields.Add("Tags");
                                            v.ViewFields.Add("Categorization");
                                            v.Update();
                                            clientcontext.ExecuteQuery();
                                        }
                                    }

                                    #region Commented

                                    //else
                                    //{
                                    //    List Pagelist = clientcontext.Web.Lists.GetByTitle(listName);
                                    //    clientcontext.Load(Pagelist);
                                    //    clientcontext.ExecuteQuery();

                                    //    ViewCollection ViewColl = Pagelist.Views;
                                    //    clientcontext.Load(ViewColl);
                                    //    clientcontext.ExecuteQuery();

                                    //    Microsoft.SharePoint.Client.View v = ViewColl[0];
                                    //    clientcontext.Load(v);
                                    //    clientcontext.ExecuteQuery();

                                    //    v.ViewFields.RemoveAll();
                                    //    v.Update();
                                    //    clientcontext.ExecuteQuery();

                                    //    v.ViewFields.Add("DocIcon");
                                    //    v.ViewFields.Add("Title");
                                    //    v.ViewFields.Add("LinkFilename");
                                    //    v.ViewFields.Add("Created");
                                    //    v.ViewFields.Add("Created By");
                                    //    v.ViewFields.Add("Modified");
                                    //    v.ViewFields.Add("Modified By");
                                    //    v.ViewFields.Add("Tags");
                                    //    v.ViewFields.Add("Categorization");
                                    //    v.Update();
                                    //    clientcontext.ExecuteQuery();
                                    //}

                                    //Pagelist.ContentTypesEnabled = true;
                                    //Pagelist.Update();
                                    //clientcontext.ExecuteQuery(); 

                                    #endregion
                                }
                            }
                            catch (Exception ex)
                            {
                                continue;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
                #region OLD
                //    this.Text = lstSiteColl[j] + "  Processing...";

                //    string startingTime = DateTime.Now.ToString();

                //    try
                //    {
                //        siteTitle = string.Empty;
                //        AuthenticationManager authManager = new AuthenticationManager();

                //        List<string> SPoExist = new List<string>();

                //        SPoExist.Add("spo.admin.verinon@agilent.onmicrosoft.com");
                //        SPoExist.Add("spo.admin2.verinon@agilent.onmicrosoft.com");
                //        SPoExist.Add("spo.admin3.verinon@agilent.onmicrosoft.com");
                //        SPoExist.Add("spo.admin4.verinon@agilent.onmicrosoft.com");
                //        SPoExist.Add("spo.admin5.verinon@agilent.onmicrosoft.com");

                //        string actualSPO = string.Empty;

                //        foreach (string sp in SPoExist)
                //        {
                //            try
                //            {
                //                using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), sp, "Lot62215"))
                //                {
                //                    clientcontext.Load(clientcontext.Web);
                //                    clientcontext.ExecuteQuery();

                //                    actualSPO = sp;

                //                    break;
                //                }
                //            }
                //            catch (Exception ex)
                //            {
                //                continue;
                //            }
                //        }

                //        //using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), "adam.a@VerinonTechnology.onmicrosoft.com", "Lot62215##"))
                //        using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), actualSPO, "Lot62215"))
                //        {


                //            ListCollection _Lists = clientcontext.Web.Lists;
                //            clientcontext.Load(_Lists);
                //            clientcontext.ExecuteQuery();

                //            bool _dListExist = clientcontext.Web.Lists.Cast<List>().Any(list => string.Equals(list.Title, "Site Assets"));

                //            if (_dListExist)
                //            {
                //                List Pagelist = clientcontext.Web.Lists.GetByTitle("Site Assets");
                //                clientcontext.Load(Pagelist);
                //                clientcontext.Load(Pagelist.RootFolder);
                //                clientcontext.ExecuteQuery();

                //                Pagelist.ContentTypesEnabled = true;
                //                Pagelist.Update();
                //                clientcontext.ExecuteQuery();
                //            }

                //            string admins = string.Empty;
                //            List<UserEntity> adminsColl = clientcontext.Site.RootWeb.GetAdministrators();

                //            foreach (UserEntity admin in adminsColl)
                //            {//SPO Admin 

                //                //User adUser = clientcontext.Site.RootWeb.SiteUsers.GetByLoginName(admin.LoginName);
                //                //adUser.is

                //                if (admin.Title != "FUN-SPO-SITECOLL-ADMINS" && (!admin.Title.ToLower().Contains("spo admin")) && (!admin.Email.ToLower().Contains("spo.admin@agilent.onmicrosoft.com")))
                //                {
                //                    if (!string.IsNullOrEmpty(admin.Email))
                //                    {
                //                        admins += admin.Email + ";";
                //                    }
                //                    else
                //                    {
                //                        admins += admin.Title + ";";
                //                    }
                //                }
                //            }

                //            Web oWebcurr = clientcontext.Site.RootWeb;
                //            clientcontext.Load(oWebcurr);
                //            clientcontext.ExecuteQuery();

                //            BuiltinGroups.Clear();
                //            ADGroups.Clear();

                //            siteTitle = oWebcurr.Title;

                //            string siteCollName = siteTitle.Replace(" ", "_");

                //            siteCollName = siteCollName.Replace("//", "_");

                //            string siteCollNameFileName = string.Empty;

                //            StreamWriter excelWriterScoringNew = null;

                //            if (!string.IsNullOrEmpty(siteTitle))
                //            {
                //                siteCollNameFileName = siteCollName;
                //                excelWriterScoringNew = System.IO.File.CreateText(textBox2.Text + "\\" + siteCollName + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
                //            }
                //            else
                //            {
                //                string[] siteCollNameFileNameXX = lstSiteColl[j].ToString().Trim().Split(new char[] { '/' });

                //                string actName = siteCollNameFileNameXX[siteCollNameFileNameXX.Length - 1];

                //                actName = actName.Replace(" ", "_");
                //                actName = actName.Replace("\\", "_");

                //                siteCollNameFileName = actName;
                //                excelWriterScoringNew = System.IO.File.CreateText(textBox2.Text + "\\" + actName + "_Report_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
                //            }

                //            excelWriterScoringNew.WriteLine("Site Coll Owners" + "," + admins + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "");
                //            excelWriterScoringNew.Flush();


                //            excelWriterScoringNew.WriteLine("Object Type" + "," + "URL" + "," + "Group" + "," + "Given though" + "," + "Folders" + "," + "Files" + "," + "Design" + "," + "Contribute" + "," + "Read" + "," + "Full Control" + "," + "Edit" + "," + "View Only" + "," + "Approve" + "," + "Contribute Limited" + "," + "OtherPermissions");
                //            excelWriterScoringNew.Flush();

                //            //////excelWriterScoringNew.WriteLine("Object Type" + "," + "URL" + "," + "SiteCollection Owners" + "," + "AD Group/Everyone granted directly" + "," + "Granted directly/added inside SP-Group" + "," + "Total number of Folders" + "," + "Total number of Files" + "," + "Design" + "," + "Contribute" + "," + "Design" + "," + "Design" + "," + "Design" + "," + "Design" + "," + "Design" + "," + "Design" + "," + "Design");
                //            //excelWriterScoringNew.WriteLine("Object Type" + "," + "URL" + "," + "SiteCollection Owners" + "," + "AD Group/Everyone granted directly" + "," + "Granted directly/added inside SP-Group" + "," + "Total number of Folders" + "," + "Total number of Files" + "," + "Design" + "," + "Contribute" + "," + "Read" + "," + "Full Control" + "," + "Edit" + "," + "View Only" + "," + "Approve" + "," + "Contribute Limited" + "," + "OtherPermissions");
                //            //excelWriterScoringNew.Flush();

                //            #region Site Coll

                //            RoleAssignmentCollection webRoleAssignments = null;
                //            GroupCollection webGroups = null;

                //            try
                //            {
                //                webRoleAssignments = clientcontext.Web.RoleAssignments;
                //                clientcontext.Load(webRoleAssignments);
                //                clientcontext.ExecuteQuery();

                //                clientcontext.Load(clientcontext.Web);
                //                clientcontext.ExecuteQuery();

                //                webGroups = clientcontext.Web.SiteGroups;
                //                clientcontext.Load(webGroups);
                //                clientcontext.ExecuteQuery();

                //                bool foundatSiteLevel = false;

                //                string AdGroupsinGroup = string.Empty;
                //                string AdGroupsatSite = string.Empty;

                //                foreach (RoleAssignment member1 in webRoleAssignments)
                //                { //c:0u.c|tenant|                             

                //                    try
                //                    {
                //                        //if (!foundatSiteLevel)
                //                        //{
                //                        clientcontext.Load(member1.Member);
                //                        clientcontext.ExecuteQuery();

                //                        if (member1.Member.Title.Contains("c:0u.c|tenant|"))
                //                        {
                //                            continue;
                //                        }

                //                        #region Role Definations

                //                        RoleDefinitionBindingCollection rdefColl = member1.RoleDefinitionBindings;
                //                        clientcontext.Load(rdefColl);
                //                        clientcontext.ExecuteQuery();

                //                        string Design = string.Empty;
                //                        string Contribute = string.Empty;
                //                        string Read = string.Empty;
                //                        string FullControl = string.Empty;
                //                        string Edit = string.Empty;
                //                        string ViewOnly = string.Empty;
                //                        string Approve = string.Empty;
                //                        string ContributeLimited = string.Empty;
                //                        string OtherPermissions = string.Empty;

                //                        foreach (RoleDefinition rdef in rdefColl)
                //                        {
                //                            clientcontext.Load(rdef);
                //                            clientcontext.ExecuteQuery();

                //                            switch (rdef.Name)
                //                            {
                //                                case "Design":
                //                                    Design = "Yes";
                //                                    break;

                //                                case "Contribute":
                //                                    Contribute = "Yes";
                //                                    break;

                //                                case "Read":
                //                                    Read = "Yes";
                //                                    break;

                //                                case "Full Control":
                //                                    FullControl = "Yes";
                //                                    break;

                //                                case "Edit":
                //                                    Edit = "Yes";
                //                                    break;

                //                                case "View Only":
                //                                    ViewOnly = "Yes";
                //                                    break;

                //                                case "Contribute Limited":
                //                                    ContributeLimited = "Yes";
                //                                    break;

                //                                case "Approve":
                //                                    Approve = "Yes";
                //                                    break;

                //                                default:
                //                                    OtherPermissions = rdef.Name;
                //                                    break;
                //                            }
                //                        }

                //                        #endregion

                //                        if (member1.Member.PrincipalType == PrincipalType.SharePointGroup)
                //                        {
                //                            Group ouserGroup = (Group)member1.Member.TypedObject;
                //                            clientcontext.Load(ouserGroup);
                //                            clientcontext.ExecuteQuery();

                //                            UserCollection userColl = ouserGroup.Users;
                //                            clientcontext.Load(userColl);
                //                            clientcontext.ExecuteQuery();

                //                            foreach (User xUser in userColl)
                //                            {
                //                                if (xUser.PrincipalType == PrincipalType.SecurityGroup)
                //                                {
                //                                    //excelWriterScoringNew.WriteLine("\"" + "SiteCollection" + "\"" + "," + "\"" + clientcontext.Web.Url.ToString() + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + "--" + "\"" + "," + "\"" + ouserGroup.Title + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"");
                //                                    //AdGroupsinGroup += ouserGroup.Title + "; ";
                //                                    //break;
                //                                    if (lstADGroupsColl.Contains(xUser.Title.ToString().Trim().ToLower()))
                //                                    {
                //                                        if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                //                                        {
                //                                            if (xUser.Title.ToString().ToLower().Contains("everyone"))
                //                                            {
                //                                                //if (!BuiltinGroups.Contains(xUser.Title))
                //                                                //{
                //                                                //    BuiltinGroups.Add(xUser.Title);
                //                                                //}

                //                                                if (BuiltinGroups.ContainsKey(xUser.Title))
                //                                                {
                //                                                    BuiltinGroups[xUser.Title]++;
                //                                                }
                //                                                else
                //                                                {
                //                                                    BuiltinGroups.Add(xUser.Title, 1);
                //                                                }
                //                                            }
                //                                            else
                //                                            {
                //                                                //if (!ADGroups.Contains(xUser.Title))
                //                                                //{
                //                                                //    ADGroups.Add(xUser.Title);
                //                                                //}

                //                                                if (ADGroups.ContainsKey(xUser.Title))
                //                                                {
                //                                                    ADGroups[xUser.Title]++;
                //                                                }
                //                                                else
                //                                                {
                //                                                    ADGroups.Add(xUser.Title, 1);
                //                                                }
                //                                            }

                //                                            excelWriterScoringNew.WriteLine("\"" + "SiteCollection" + "\"" + "," + "\"" + clientcontext.Web.Url.ToString() + "\"" + "," + "\"" + xUser.Title + "\"" + "," + "\"" + ouserGroup.Title + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                //                                            excelWriterScoringNew.Flush();
                //                                        }
                //                                    }
                //                                    //foundatSiteLevel = true;
                //                                    //break;
                //                                }

                //                                //if (xUser.Title == "Everyone except external users")
                //                                //{
                //                                //    excelWriterScoringNew.WriteLine("\"" + "SiteCollection" + "\"" + "," + "\"" + clientcontext.Web.Url.ToString() + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + "Yes" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"");

                //                                //    foundatSiteLevel = true;
                //                                //    break;
                //                                //}
                //                            }
                //                        }
                //                        if (member1.Member.PrincipalType == PrincipalType.SecurityGroup)
                //                        {
                //                            //if (member1.Member.Title == "Everyone except external users")
                //                            //{
                //                            //excelWriterScoringNew.WriteLine("\"" + "SiteCollection" + "\"" + "," + "\"" + clientcontext.Web.Url.ToString() + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + member1.Member.Title + "\"" + "," + "\"" + "--" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"");
                //                            //AdGroupsatSite += member1.Member.Title + "; ";

                //                            if (lstADGroupsColl.Contains(member1.Member.Title.ToString().Trim().ToLower()))
                //                            {
                //                                if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                //                                {
                //                                    if (member1.Member.Title.ToString().ToLower().Contains("everyone"))
                //                                    {
                //                                        //if (!BuiltinGroups.Contains(member1.Member.Title))
                //                                        //{
                //                                        //    BuiltinGroups.Add(member1.Member.Title);
                //                                        //}

                //                                        if (BuiltinGroups.ContainsKey(member1.Member.Title))
                //                                        {
                //                                            BuiltinGroups[member1.Member.Title]++;
                //                                        }
                //                                        else
                //                                        {
                //                                            BuiltinGroups.Add(member1.Member.Title, 1);
                //                                        }
                //                                    }
                //                                    else
                //                                    {
                //                                        //if (!ADGroups.Contains(member1.Member.Title))
                //                                        //{
                //                                        //    ADGroups.Add(member1.Member.Title);
                //                                        //}
                //                                        if (ADGroups.ContainsKey(member1.Member.Title))
                //                                        {
                //                                            ADGroups[member1.Member.Title]++;
                //                                        }
                //                                        else
                //                                        {
                //                                            ADGroups.Add(member1.Member.Title, 1);
                //                                        }
                //                                    }

                //                                    excelWriterScoringNew.WriteLine("\"" + "SiteCollection" + "\"" + "," + "\"" + clientcontext.Web.Url.ToString() + "\"" + "," + "\"" + member1.Member.Title + "\"" + "," + "\"" + "Directly Assigned" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                //                                    excelWriterScoringNew.Flush();
                //                                }
                //                            }
                //                            //foundatSiteLevel = true;
                //                            //break;
                //                            //}
                //                        }

                //                        #region Commented Is Uesr

                //                        //if (member1.Member.PrincipalType == PrincipalType.User)
                //                        //{
                //                        //    if (member1.Member.Title == "Everyone except external users")
                //                        //    {
                //                        //        excelWriterScoringNew.WriteLine("\"" + "Site" + "\"" + "," + "\"" + clientcontext.Web.Url.ToString() + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + "Yes" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"");
                //                        //        foundatSiteLevel = true;
                //                        //        break;
                //                        //    }
                //                        //} 

                //                        #endregion
                //                        //}
                //                        //else
                //                        //{
                //                        //    break;
                //                        //}
                //                    }
                //                    catch (Exception ex)
                //                    {
                //                        continue;
                //                    }
                //                }
                //                //excelWriterScoringNew.WriteLine("\"" + "SiteCollection" + "\"" + "," + "\"" + clientcontext.Web.Url.ToString() + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + AdGroupsatSite + "\"" + "," + "\"" + AdGroupsinGroup + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"");
                //                //excelWriterScoringNew.Flush();
                //            }
                //            catch (Exception ex)
                //            {
                //            }

                //            #endregion

                //            #region Lists

                //            ListCollection olistColl = clientcontext.Web.Lists;
                //            clientcontext.Load(olistColl);
                //            clientcontext.ExecuteQuery();

                //            foreach (List oList in olistColl)
                //            {
                //                bool foundatListLevel = false;

                //                clientcontext.Load(oList);
                //                clientcontext.Load(oList, li => li.HasUniqueRoleAssignments);
                //                clientcontext.ExecuteQuery();

                //                if (oList.BaseType == BaseType.DocumentLibrary)
                //                {
                //                    if (oList.Title == "Documents")
                //                    {
                //                        bool foXXXXundatListLevel = false;
                //                    }

                //                    if ((oList.Title != "Form Templates" && oList.Title != "Site Assets" && oList.Title != "SitePages" && oList.Title != "Style Library" && oList.Hidden == false && oList.IsCatalog == false && oList.BaseTemplate == 101) || oList.BaseTemplate == 700)
                //                    {
                //                        string UniqueRoles = string.Empty;

                //                        #region Commented Test

                //                        //if (oList.Title == "Documents")
                //                        //{
                //                        //    clientcontext.Load(oList.RootFolder);
                //                        //    clientcontext.ExecuteQuery();

                //                        //    clientcontext.Load(clientcontext.Web);
                //                        //    clientcontext.ExecuteQuery();

                //                        //    GetCounts(oList.RootFolder, clientcontext);
                //                        //}

                //                        #endregion

                //                        //if (oList.HasUniqueRoleAssignments)
                //                        //{
                //                        //    UniqueRoles = "Unique Permissions";
                //                        //}
                //                        //else
                //                        //{
                //                        //    UniqueRoles = "Inherit from Parent";
                //                        //}

                //                        if (oList.HasUniqueRoleAssignments)
                //                        {
                //                            ListFoldCount = 0;
                //                            ListFileCount = 0;

                //                            GetCountsatListLevel(oList.RootFolder, clientcontext);

                //                            string AdGroupsinSileCollListGroup = string.Empty;
                //                            string AdGroupsatSileCollListSite = string.Empty;

                //                            #region SiteColl Lists Permission check

                //                            RoleAssignmentCollection roles = oList.RoleAssignments;
                //                            clientcontext.Load(roles);
                //                            clientcontext.ExecuteQuery();

                //                            Web oWebx = clientcontext.Web;
                //                            clientcontext.Load(oWebx);
                //                            clientcontext.ExecuteQuery();

                //                            foreach (RoleAssignment rAssignment in roles)
                //                            {


                //                                #region Role Definations

                //                                RoleDefinitionBindingCollection rdefColl = rAssignment.RoleDefinitionBindings;
                //                                clientcontext.Load(rdefColl);
                //                                clientcontext.ExecuteQuery();

                //                                string Design = string.Empty;
                //                                string Contribute = string.Empty;
                //                                string Read = string.Empty;
                //                                string FullControl = string.Empty;
                //                                string Edit = string.Empty;
                //                                string ViewOnly = string.Empty;
                //                                string Approve = string.Empty;
                //                                string ContributeLimited = string.Empty;
                //                                string OtherPermissions = string.Empty;

                //                                foreach (RoleDefinition rdef in rdefColl)
                //                                {
                //                                    clientcontext.Load(rdef);
                //                                    clientcontext.ExecuteQuery();

                //                                    switch (rdef.Name)
                //                                    {
                //                                        case "Design":
                //                                            Design = "Yes";
                //                                            break;

                //                                        case "Contribute":
                //                                            Contribute = "Yes";
                //                                            break;

                //                                        case "Read":
                //                                            Read = "Yes";
                //                                            break;

                //                                        case "Full Control":
                //                                            FullControl = "Yes";
                //                                            break;

                //                                        case "Edit":
                //                                            Edit = "Yes";
                //                                            break;

                //                                        case "View Only":
                //                                            ViewOnly = "Yes";
                //                                            break;

                //                                        case "Contribute Limited":
                //                                            ContributeLimited = "Yes";
                //                                            break;

                //                                        case "Approve":
                //                                            Approve = "Yes";
                //                                            break;

                //                                        default:
                //                                            OtherPermissions = rdef.Name;
                //                                            break;
                //                                    }
                //                                }

                //                                #endregion

                //                                try
                //                                {
                //                                    //if (!foundatListLevel)
                //                                    //{
                //                                    clientcontext.Load(rAssignment.Member);
                //                                    clientcontext.ExecuteQuery();

                //                                    if (rAssignment.Member.Title.Contains("c:0u.c|tenant|"))
                //                                    {
                //                                        continue;
                //                                    }

                //                                    if (rAssignment.Member.PrincipalType == PrincipalType.SharePointGroup)
                //                                    {
                //                                        Group ouserGroup = (Group)rAssignment.Member.TypedObject;
                //                                        clientcontext.Load(ouserGroup);
                //                                        clientcontext.ExecuteQuery();

                //                                        UserCollection userColl = ouserGroup.Users;
                //                                        clientcontext.Load(userColl);
                //                                        clientcontext.ExecuteQuery();

                //                                        foreach (User xUser in userColl)
                //                                        {
                //                                            if (xUser.PrincipalType == PrincipalType.SecurityGroup)
                //                                            {

                //                                                //if (xUser.Title == "Everyone except external users")
                //                                                //{
                //                                                //clientcontext.Load(oList.RootFolder);
                //                                                //clientcontext.ExecuteQuery();   

                //                                                //AdGroupsinSileCollListGroup += ouserGroup.Title + ";";
                //                                                //foundatListLevel = true;
                //                                                //break;

                //                                                if (lstADGroupsColl.Contains(xUser.Title.ToString().Trim().ToLower()))
                //                                                {
                //                                                    if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                //                                                    {
                //                                                        if (xUser.Title.ToString().ToLower().Contains("everyone"))
                //                                                        {
                //                                                            //if (!BuiltinGroups.Contains(xUser.Title))
                //                                                            //{
                //                                                            //    BuiltinGroups.Add(xUser.Title);
                //                                                            //}

                //                                                            if (BuiltinGroups.ContainsKey(xUser.Title))
                //                                                            {
                //                                                                BuiltinGroups[xUser.Title]++;
                //                                                            }
                //                                                            else
                //                                                            {
                //                                                                BuiltinGroups.Add(xUser.Title, 1);
                //                                                            }
                //                                                        }
                //                                                        else
                //                                                        {
                //                                                            //if (!ADGroups.Contains(xUser.Title))
                //                                                            //{
                //                                                            //    ADGroups.Add(xUser.Title);
                //                                                            //}
                //                                                            if (ADGroups.ContainsKey(xUser.Title))
                //                                                            {
                //                                                                ADGroups[xUser.Title]++;
                //                                                            }
                //                                                            else
                //                                                            {
                //                                                                ADGroups.Add(xUser.Title, 1);
                //                                                            }
                //                                                        }

                //                                                        excelWriterScoringNew.WriteLine("\"" + "Doc Library" + "\"" + "," + "\"" + oWebx.Url.Replace(oWebx.ServerRelativeUrl, "") + oList.RootFolder.ServerRelativeUrl + "\"" + "," + "\"" + xUser.Title + "\"" + "," + "\"" + ouserGroup.Title + "\"" + "," + "\"" + ListFoldCount.ToString() + "\"" + "," + "\"" + ListFileCount.ToString() + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                //                                                        excelWriterScoringNew.Flush();
                //                                                        //excelWriterScoringNew.WriteLine("\"" + "Doc Library" + "\"" + "," + "\"" + oWebx.Url.Replace(oWebx.ServerRelativeUrl, "") + oList.RootFolder.ServerRelativeUrl + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + "--" + "\"" + "\"" + "," + "\"" + ouserGroup.Title + "\"" + "," + "\"" + UniqueRoles + "\"" + "," + "\"" + FoldCount.ToString() + "\"" + "," + "\"" + FileCount.ToString() + "\"");
                //                                                    }
                //                                                }

                //                                                //break;
                //                                            }
                //                                        }
                //                                    }
                //                                    if (rAssignment.Member.PrincipalType == PrincipalType.SecurityGroup)
                //                                    {
                //                                        //if (rAssignment.Member.Title == "Everyone except external users")
                //                                        //{
                //                                        //clientcontext.Load(oList.RootFolder);
                //                                        //clientcontext.ExecuteQuery();                                                  

                //                                        //AdGroupsatSileCollListSite += rAssignment.Member.Title + ";";
                //                                        //foundatListLevel = true;
                //                                        if (lstADGroupsColl.Contains(rAssignment.Member.Title.ToString().Trim().ToLower()))
                //                                        {


                //                                            if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                //                                            {
                //                                                if (rAssignment.Member.Title.ToString().ToLower().Contains("everyone"))
                //                                                {
                //                                                    //if (!BuiltinGroups.Contains(rAssignment.Member.Title))
                //                                                    //{
                //                                                    //    BuiltinGroups.Add(rAssignment.Member.Title);
                //                                                    //}

                //                                                    if (BuiltinGroups.ContainsKey(rAssignment.Member.Title))
                //                                                    {
                //                                                        BuiltinGroups[rAssignment.Member.Title]++;
                //                                                    }
                //                                                    else
                //                                                    {
                //                                                        BuiltinGroups.Add(rAssignment.Member.Title, 1);
                //                                                    }
                //                                                }
                //                                                else
                //                                                {
                //                                                    //if (!ADGroups.Contains(rAssignment.Member.Title))
                //                                                    //{
                //                                                    //    ADGroups.Add(rAssignment.Member.Title);
                //                                                    //}

                //                                                    if (ADGroups.ContainsKey(rAssignment.Member.Title))
                //                                                    {
                //                                                        ADGroups[rAssignment.Member.Title]++;
                //                                                    }
                //                                                    else
                //                                                    {
                //                                                        ADGroups.Add(rAssignment.Member.Title, 1);
                //                                                    }
                //                                                }

                //                                                excelWriterScoringNew.WriteLine("\"" + "Doc Library" + "\"" + "," + "\"" + oWebx.Url.Replace(oWebx.ServerRelativeUrl, "") + oList.RootFolder.ServerRelativeUrl + "\"" + "," + "\"" + rAssignment.Member.Title + "\"" + "," + "\"" + "Directly Assigned" + "\"" + "," + "\"" + ListFoldCount.ToString() + "\"" + "," + "\"" + ListFileCount.ToString() + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                //                                                excelWriterScoringNew.Flush();
                //                                            }
                //                                        }
                //                                        //excelWriterScoringNew.WriteLine("\"" + "Doc Library" + "\"" + "," + "\"" + oWebx.Url.Replace(oWebx.ServerRelativeUrl, "") + oList.RootFolder.ServerRelativeUrl + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + "Yes" + "\"" + "," + "\"" + "--" + "\"" + "\"" + "," + "\"" + UniqueRoles + "\"" + "," + "\"" + FoldCount.ToString() + "\"" + "," + "\"" + FileCount.ToString() + "\"");


                //                                        //break;
                //                                        //}
                //                                    }
                //                                    //}
                //                                    //else
                //                                    //{
                //                                    //    break;
                //                                    //}
                //                                }
                //                                catch (Exception ex)
                //                                {
                //                                    continue;
                //                                }
                //                            }

                //                            //if (foundatListLevel)
                //                            //{
                //                            //    ListFoldCount = 0;
                //                            //    ListFileCount = 0;

                //                            //    GetCountsatListLevel(oList.RootFolder, clientcontext);

                //                            //    excelWriterScoringNew.WriteLine("\"" + "Doc Library" + "\"" + "," + "\"" + oWebx.Url.Replace(oWebx.ServerRelativeUrl, "") + oList.RootFolder.ServerRelativeUrl + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + AdGroupsatSileCollListSite + "\"" + "," + "\"" + AdGroupsinSileCollListGroup + "\"" + "," + "\"" + ListFoldCount.ToString() + "\"" + "," + "\"" + ListFileCount.ToString() + "\"");
                //                            //    excelWriterScoringNew.Flush();
                //                            //}

                //                            #endregion
                //                        }

                //                        clientcontext.Load(oList.RootFolder.Folders);
                //                        clientcontext.ExecuteQuery();

                //                        foreach (Folder sFolder in oList.RootFolder.Folders)
                //                        {
                //                            clientcontext.Load(sFolder);
                //                            clientcontext.ExecuteQuery();

                //                            if (sFolder.Name != "Forms")
                //                            {
                //                                GetCounts(sFolder, clientcontext, excelWriterScoringNew);
                //                            }
                //                        }
                //                    }
                //                }
                //            }

                //            #endregion

                //            #region SubSites

                //            WebCollection oWebs = clientcontext.Web.Webs;
                //            clientcontext.Load(oWebs);
                //            clientcontext.ExecuteQuery();

                //            foreach (Web oWeb in oWebs)
                //            {
                //                try
                //                {
                //                    clientcontext.Load(oWeb);
                //                    clientcontext.ExecuteQuery();
                //                    this.Text = oWeb.Url + "  Processing...";
                //                    getWeb(oWeb.Url, excelWriterScoringNew);
                //                }
                //                catch (Exception ex)
                //                {
                //                    continue;
                //                }
                //            }

                //            #endregion


                //            excelWriterScoringNew.Flush();
                //            excelWriterScoringNew.Close();

                //            string bGroups = string.Empty;
                //            string AdsGroups = string.Empty;

                //            foreach (KeyValuePair<string, int> kp in BuiltinGroups)
                //            {
                //                if (kp.Key != "FUN-SPO-SITECOLL-ADMINS" && (!kp.Key.ToLower().Contains("spo admin")))
                //                {
                //                    bGroups += kp.Key.ToString().Trim() + "(" + kp.Value.ToString() + ")" + "; ";
                //                }
                //            }

                //            foreach (KeyValuePair<string, int> kp in ADGroups)
                //            {
                //                if (kp.Key != "FUN-SPO-SITECOLL-ADMINS" && (!kp.Key.ToLower().Contains("spo admin")))
                //                {
                //                    AdsGroups += kp.Key.ToString().Trim() + "(" + kp.Value.ToString() + ")" + "; ";
                //                }
                //            }

                //            //foreach (string gp in BuiltinGroups)
                //            //{
                //            //    bGroups += gp + "; ";
                //            //}

                //            //foreach (string ap in ADGroups)
                //            //{
                //            //    AdsGroups += ap + "; ";
                //            //}

                //            excelWriterScoringMatrixNew.WriteLine("\"" + siteCollNameFileName + ".xlsx" + "\"" + "," + "\"" + clientcontext.Web.Url.ToString() + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + bGroups + "\"" + "," + "\"" + AdsGroups + "\"" + "," + "\"" + startingTime + "\"" + "," + "\"" + DateTime.Now.ToString() + "\"" + "," + "\"" + "" + "\"");
                //            excelWriterScoringMatrixNew.Flush();

                //            //excelWriterScoringMatrixNew.WriteLine(siteCollNameFileName +".xlsx" + "," + clientcontext.Web.Url.ToString() + "," + admins + "," + bGroups + "," + AdsGroups + "," + DateTime.Now.ToString());
                //            //excelWriterScoringMatrixNew.Flush();
                //        }
                //    }
                //    catch (Exception ex)
                //    {
                //        excelWriterScoringMatrixNew.WriteLine("\"" + "--" + "\"" + "," + "\"" + lstSiteColl[j] + "\"" + "," + "\"" + "" + "\"" + "," + "\"" + "" + "\"" + "," + "\"" + "" + "\"" + "," + "\"" + startingTime + "\"" + "," + "\"" + DateTime.Now.ToString() + "\"" + "," + "\"" + ex.Message + "\"");
                //        excelWriterScoringMatrixNew.Flush();

                //        continue;
                //    }
                //}

                //excelWriterScoringMatrixNew.Flush();
                //excelWriterScoringMatrixNew.Close();

                //this.Text = "Process completed successfully.";
                //MessageBox.Show("Process Completed"); 
                #endregion
            }
            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");
        }

        private void btninputCSV_Click(object sender, EventArgs e)
        {
            
            OpenFileDialog inputfile = new OpenFileDialog();

            if(inputfile.ShowDialog()==DialogResult.OK)
            {
                txtInput.Text = inputfile.FileName;
            }
            #region Commented
            //inputfile.Multiselect = true;
            //inputfile.ShowDialog();
            //inputfile.Filter = "allfiles|*.xls";
            //txtInput.Text = inputfile.FileName;
            //int count = 0;
            //string[] Fname;
            //foreach(string s in inputfile.FileNames)
            //{
            //    Fname = s.Split('\\');
            //    File.Copy(s,  Fname[Fname.Length - 1]);
            //    count++;
            //}
            //MessageBox.Show(Convert.ToString(count) + " File(s) copied"); 
            #endregion
        }

        private void btninputReport_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog reportFolder = new FolderBrowserDialog();

            if(reportFolder.ShowDialog()==DialogResult.OK)
            {
                txtReport.Text = reportFolder.SelectedPath;
            }
            #region Commented Code
            //reportFolder.ShowDialog();
            //txtReport.Text = reportFolder.SelectedPath;            
            //string path = txtReport.Text;
            //StreamWriter sw = new StreamWriter(path);
            //System.IO.File.WriteAllLines(path,"Reports.csv",) 
            #endregion

        }
    }
}
