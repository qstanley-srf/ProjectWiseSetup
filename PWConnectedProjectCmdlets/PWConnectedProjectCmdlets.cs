using System;
using System.Data;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.IO;
using System.Collections;
using System.Collections.Generic;

namespace PWConnectedProjectCmdlets
{
    [Cmdlet(VerbsCommon.Add, "DisciplineFolders")]
    public class AddDisciplineFolders : PSCmdlet
    {
        [Parameter]
        public string Template
        {
            get { return template; }
            set { template = value; }
        }
        private string template;

        [Parameter]
        public string TopLevel
        {
            get { return topLevel; }
            set { topLevel = value; }
        }
        private string topLevel;

        [Parameter]
        public string ProjectName
        {
            get { return projectName; }
            set { projectName = value; }
        }
        private string projectName;

        [Parameter]
        public string[] FunctionalGroups
        {
            get { return functionalGroups; }
            set { functionalGroups = value; }
        }
        private string[] functionalGroups;

        protected override void EndProcessing()
        {
            List<string> NDDisciplines = new List<string>();
            NDDisciplines.Add("NorthDakotaAsBuilt");
            NDDisciplines.Add("NorthDakotaDesign");
            NDDisciplines.Add("NorthDakotaBridge");
            NDDisciplines.Add("NorthDakotaConstruction");
            NDDisciplines.Add("NorthDakotaGIS");
            NDDisciplines.Add("NorthDakotaROW");
            foreach (string fullDiscipline in functionalGroups)
            {
                string discipline = fullDiscipline;
                //Set source folder
                string sourceFolder = "Templates\\Additional_Pilot_Project_Folders\\" + discipline;
                if (NDDisciplines.Contains(discipline))
                {
                    discipline = discipline.Substring(11);
                    sourceFolder = "Templates\\Additional_Pilot_Project_Folders\\NorthDakota Additional Folders\\" + discipline;
                }
                //Set target folder
                string targetFolder;
                if (template == "SRF_ND_Pilot_Project")
                {
                    targetFolder = topLevel + "\\" + projectName + "\\NDProjectNo\\TechData\\";
                    if (discipline == "WR_CAD")
                    {
                        targetFolder = targetFolder + "WaterResources\\";
                    }
                }
                else
                {
                    targetFolder = topLevel + "\\" + projectName + "\\TechData\\";
                    if (discipline == "WR_CAD")
                    {
                        targetFolder = targetFolder + "WaterResources\\";
                    }
                }
                //Resulting folder path
                string resultFolder = targetFolder + discipline;
                string resultFolderCAD = targetFolder + "CADDesign";
                //Parameters for copy command
                Hashtable copyParameters = new Hashtable();
                copyParameters.Add("FolderPath", targetFolder);
                copyParameters.Add("FolderToCopy", sourceFolder);
                copyParameters.Add("IncludeSubFolders", true);
                copyParameters.Add("IncludeDocuments", true);
                copyParameters.Add("IncludeDocumentAttributes", true);
                copyParameters.Add("IncludeAccessControl", true);
                using (var ps = PowerShell.Create(RunspaceMode.CurrentRunspace))
                {
                    //Copy and rename folders
                    ps.AddCommand("Copy-PWFolder").AddParameters(copyParameters);
                    ps.AddCommand("Out-Null");
                    ps.Invoke();
                }
                if (discipline.Contains("CADDesign"))
                {
                    using (var ps1 = PowerShell.Create(RunspaceMode.CurrentRunspace))
                    {
                        ps1.AddStatement().AddCommand("Update-PWFolderNameProps").AddParameter("FolderPath", resultFolder).AddParameter("NewName", "CADDesign");
                        ps1.AddCommand("Out-Null");
                        resultFolder = resultFolderCAD;
                        ps1.Invoke();
                    }
                }
                using (var ps = PowerShell.Create(RunspaceMode.CurrentRunspace))
                {
                    //Set views
                    ps.AddScript("$thisFolder = Get-PWFolders -FolderPath \"" + sourceFolder + "\" -JustOne").Invoke();
                    ps.AddScript("$thatFolder = Get-PWFolders -FolderPath \"" + resultFolder + "\" -JustOne").Invoke();
                    ps.AddScript("Copy-PWFolderViewsToFolders -SourceFolder $thisFolder -TargetFolder $thatFolder").Invoke();
                }
            }
        }
    }

    [Cmdlet(VerbsCommon.Add, "PhaseDocuments")]
    public class AddPhaseDocuments : Cmdlet
    {
        [Parameter]
        public DataRow[] Phases
        {
            get { return phases; }
            set { phases = value; }
        }
        private DataRow[] phases;

        [Parameter]
        public DataRow[] DefaultPhases
        {
            get { return defaultPhases; }
            set { defaultPhases = value; }
        }
        private DataRow[] defaultPhases;

        [Parameter]
        public string PhaseFolderPath
        {
            get { return phaseFolderPath; }
            set { phaseFolderPath = value; }
        }
        private string phaseFolderPath;

        protected override void ProcessRecord()
        {
            List<string> phaseList = new List<string>();
            //DataRow[] submittedPhases = phases.BaseObject;
            //DataRow[] defaultPhaseList = defaultPhases.BaseObject;
            foreach (DataRow phase in phases)
            {
                string phaseName = "";
                if (phase.IsNull("CustPhaseName") || phase["CustPhaseName"].ToString() == "")
                {
                    foreach (DataRow item in defaultPhases)
                    {
                        if (item["WBS2Code"].Equals(phase["PhaseCode"]))
                        {
                            phaseName = item["WBS2Name"].ToString();
                        }
                    }
                }
                else
                {
                    phaseName = phase["CustPhaseName"].ToString();
                    if (phaseName.Contains("_"))
                    {
                        phaseName = phaseName.Substring(0, phaseName.IndexOf("_"));
                    }
                }
                phaseList.Add(phaseName);
            }
            foreach (String phase in phaseList)
            {
                using (var ps = PowerShell.Create(RunspaceMode.CurrentRunspace))
                {
                    Hashtable attributes = new Hashtable();
                    attributes.Add("PHASE", phase);
                    ps.AddCommand("New-PWDocumentAbstract").AddParameter("FolderPath", phaseFolderPath);
                    ps.AddCommand("Update-PWDocumentAttributes").AddParameter("Attributes", attributes).Invoke();
                }
            }
        }
    }

    [Cmdlet(VerbsData.Edit, "CADConfigFile")]
    public class EditCADConfigFile : PSCmdlet
    {
        // Declare the parameters for the cmdlet.
        [Parameter]
        public string TopLevel
        {
            get { return topLevel; }
            set { topLevel = value; }
        }
        private string topLevel;
        [Parameter]
        public string TopLevelReplacement
        {
            get { return topLevelReplacement; }
            set { topLevelReplacement = value; }
        }
        private string topLevelReplacement;
        [Parameter]
        public string Template
        {
            get { return template; }
            set { template = value; }
        }
        private string template;
        [Parameter]
        public string ProjectName
        {
            get { return projectName; }
            set { projectName = value; }
        }
        private string projectName;
        [Parameter(Mandatory = false)]
        public string NDProjectNumber
        {
            get { return ndProjectNumber; }
            set { ndProjectNumber = value; }
        }
        private string ndProjectNumber;


        protected override void EndProcessing()
        {
            string projectPath = topLevel + "\\" + projectName;
            string configPath = projectPath + "\\_PWSetup\\Worksets\\";
            string firstTwo = projectName.Substring(0, 2);
            string tempLocation = "H:\\" + topLevel + "\\" + projectName + "\\temp\\";
            if (!(Directory.Exists(tempLocation)))
            {
                Directory.CreateDirectory(tempLocation);
            }
            //Changing names of 
            using (var ps = PowerShell.Create(RunspaceMode.CurrentRunspace))
            {
                ps.AddCommand("Get-PWDocumentsBySearch").AddParameter("FolderPath", configPath).AddParameter("DocumentName", "ProjNumber.cfg");
                ps.AddCommand("CheckOut-PWDocuments").AddParameter("Export", true).AddParameter("ExportFolder", tempLocation);
                ps.Invoke();
            }
            //Change things in the file
            string workAreaPath = "pw://srf-pw.bentley.com:srf-pw/Documents/" + topLevelReplacement + "/" + projectName + "/";
            string tempFile = tempLocation + "ProjNumber.cfg";
            string text = File.ReadAllText(tempFile);
            text = text.Replace("INSERTWORKAREAPATHHERE", workAreaPath);
            text = text.Replace("XXXXX", projectName);
            if (template == "SRF_ND_Pilot_Project")
            {
                text = text.Replace("INSERTCLIENTNUMBERHERE", ndProjectNumber);
            }
            File.WriteAllText(tempFile, text);

            using (var ps = PowerShell.Create(RunspaceMode.CurrentRunspace))
            {
                //Reimport file to PW
                ps.AddCommand("Get-PWDocumentsBySearch").AddParameter("FolderPath", configPath).AddParameter("DocumentName", "ProjNumber.cfg");
                ps.AddCommand("CheckIn-PWDocumentsOrFree").Invoke();
            }
            //Rename files
            using (var ps = PowerShell.Create(RunspaceMode.CurrentRunspace))
            {

                //ProjNumber.cfg
                ps.AddCommand("Get-PWDocumentsBySearch").AddParameter("FolderPath", configPath).AddParameter("DocumentName", "ProjNumber.cfg");
                ps.AddCommand("Rename-PWDocument").AddParameter("DocumentNewName", projectName + ".cfg").AddParameter("RenameFile", true).Invoke();
            }
            using (var ps = PowerShell.Create(RunspaceMode.CurrentRunspace))
            {
                //ProjNumber.dgnws
                ps.AddCommand("Get-PWDocumentsBySearch").AddParameter("FolderPath", configPath).AddParameter("DocumentName", "ProjNumber.dgnws");
                ps.AddCommand("Rename-PWDocument").AddParameter("DocumentNewName", projectName + ".dgnws").AddParameter("RenameFile", true).Invoke();
            }
            Directory.Delete(tempLocation, true);
        }
    }

    [Cmdlet(VerbsData.Edit, "ManagementFiles")]
    public class EditManagementFiles : PSCmdlet
    {
        // Declare the parameters for the cmdlet.
        [Parameter]
        public string ManagementFolder
        {
            get { return managementFolder; }
            set { managementFolder = value; }
        }
        private string managementFolder;
        [Parameter]
        public string PursuitFolder
        {
            get { return pursuitFolder; }
            set { pursuitFolder = value; }
        }
        private string pursuitFolder;
        [Parameter]
        public string ProjectName
        {
            get { return projectName; }
            set { projectName = value; }
        }
        private string projectName;

        protected override void EndProcessing()
        {
            int currentYear = DateTime.Now.Year;
            //PW paths
            string meetings = managementFolder + "\\Meetings(year)";
            string emails = managementFolder + "\\Emails(year)";
            string pursuitEmails = pursuitFolder + "\\Emails(year)";
            //replacement strings
            string meetingReplacement = "Meetings(" + currentYear + ")";
            string emailReplacement = "Emails(" + currentYear + ")";
            //Parameters
            IDictionary meetingsParameters = new Dictionary<String, String>();
            meetingsParameters.Add("FolderPath", meetings);
            meetingsParameters.Add("NewName", meetingReplacement);
            IDictionary emailsParameters = new Dictionary<String, String>();
            emailsParameters.Add("FolderPath", emails);
            emailsParameters.Add("NewName", emailReplacement);
            IDictionary pursuitParameters = new Dictionary<String, String>();
            pursuitParameters.Add("FolderPath", pursuitEmails);
            pursuitParameters.Add("NewName", emailReplacement);
            //Changing names of 
            using (var ps = PowerShell.Create(RunspaceMode.CurrentRunspace))
            {
                ps.AddCommand("Update-PWFolderNameProps").AddParameters(meetingsParameters);
                ps.AddStatement().AddCommand("Update-PWFolderNameProps").AddParameters(emailsParameters);
                ps.AddStatement().AddCommand("Update-PWFolderNameProps").AddParameters(pursuitParameters);
                ps.Invoke();
            }
            //Export config file to 
            string firstTwo = projectName.Substring(0, 2);
            string tempLocation = "H:\\Projects\\" + firstTwo + "000\\" + projectName + "\\temp\\";
            Directory.CreateDirectory(tempLocation);
            using (var ps = PowerShell.Create(RunspaceMode.CurrentRunspace))
            {
                ps.AddCommand("Get-PWDocumentsBySearch").AddParameter("FolderPath", managementFolder).AddParameter("FileName", "SRF_Project_Config.cfg");
                ps.AddCommand("Export-PWDocumentsSimple").AddParameter("TargetFolder", tempLocation);
                ps.Invoke();
            }
            //C:\Temp\configs\Projects\09909\Management
            //Change things in the file
            string tempFile = tempLocation + "SRF_Project_Config.cfg";
            string text = File.ReadAllText(tempFile);
            text = text.Replace("CHANGEME", projectName);
            File.WriteAllText(tempFile, text);
            //Reimport file to PW
            using (var ps = PowerShell.Create(RunspaceMode.CurrentRunspace))
            {
                ps.AddCommand("Import-PWDocuments").AddParameter("ProjectWiseFolder", managementFolder).AddParameter("InputFolder", tempLocation).AddParameter("ExcludeSourceDirectoryFromTargetPath");
                ps.Invoke();
            }
            Directory.Delete(tempLocation, true);
        }
    }

    [Cmdlet(VerbsCommon.Get, "FunctionalGroups")]
    public class GetFunctionalGroups : Cmdlet
    {
        [Parameter]
        public string Folders
        {
            get { return folders; }
            set { folders = value; }
        }
        private string folders;

        [Parameter]
        public string CADStandard
        {
            get { return cadstandard; }
            set { cadstandard = value; }
        }
        private string cadstandard;

        [Parameter]
        public string CADSoftware
        {
            get { return cadsoftware; }
            set { cadsoftware = value; }
        }
        private string cadsoftware;

        [Parameter]
        public string State
        {
            get { return state; }
            set { state = value; }
        }
        private string state;

        [Parameter]
        public bool GIS
        {
            get { return gis; }
            set { gis = value; }
        }
        private bool gis;

        public string[] functionalGroups;
        protected override void ProcessRecord()
        {
            List<string> functionalGroupsList = new List<string>();
            bool ndProject = (state.Contains("NorthDakota") || cadstandard.Contains("NorthDakota"));
            string[] ndFolders = { "Construction", "RealEstate", "Structural" };
            foreach (string folder in ndFolders)
            {
                if (folders.Contains(folder))
                {
                    if (!ndProject)
                    {
                        functionalGroupsList.Add(folder);
                    }
                    if (ndProject && folder == "Construction")
                    {
                        functionalGroupsList.Add("NorthDakotaConstruction");
                    }
                    if (ndProject && folder == "RealEstate")
                    {
                        functionalGroupsList.Add("NorthDakotaROW");
                    }
                    if (ndProject && folder == "Structural")
                    {
                        functionalGroupsList.Add("NorthDakotaBridge");
                    }
                }
            }
            if (gis && ndProject)
            {
                functionalGroupsList.Add("NorthDakotaGIS");
            }
            if (folders.Contains("RealEstate"))
            {
                if (ndProject)
                {
                    functionalGroupsList.Add("NorthDakotaROW");
                }
                else
                {
                    functionalGroupsList.Add("RealEstate");
                }
            }
            if (folders.Contains("Structural"))
            {
                if (ndProject)
                {
                    functionalGroupsList.Add("NorthDakotaBridge");
                }
                else
                {
                    functionalGroupsList.Add("Structural");
                }
            }
            string[] basicFolders = { "Electrical", "Environmental", "ITSCAV", "Landscape", "Planning", "ProjectCtrls", "PublicEngagement", "SiteDevelopment", "TrafficStudies", "TrafficEng", "WaterResources", "UtilityCoord" };
            foreach (string folder in basicFolders)
            {
                if (folders.Contains(folder))
                {
                    functionalGroupsList.Add(folder);
                }
            }
            if (cadstandard.Contains("MnDOT"))
            {
                if (cadsoftware.Contains("ORD"))
                {
                    functionalGroupsList.Add("CADDesignMNORD");
                }
                else if (cadsoftware.Contains("SS10"))
                {
                    functionalGroupsList.Add("CADDesignMNSS10");
                }
            }
            if (cadstandard.Contains("IlDOT"))
            {
                if (cadsoftware.Contains("ORD"))
                {
                    functionalGroupsList.Add("CADDesignIlORD");
                }
                else if (cadsoftware.Contains("SS10"))
                {
                    functionalGroupsList.Add("CADDesignMNSS10");
                }
            }
            if (cadstandard.Contains("NdDOT"))
            {
                functionalGroupsList.Add("NorthDakotaDesign");
            }
            if (cadstandard.ToLower().Contains("sd"))
            {
                if (cadsoftware.Contains("ORD"))
                {
                    functionalGroupsList.Add("CADDesignSD2023ORD");
                }
            }
            if (cadstandard.Contains("Hennepin"))
            {
                if (cadsoftware.Contains("ORD"))
                {
                    functionalGroupsList.Add("CADDesignHennCntyORD");
                }
            }
            functionalGroups = functionalGroupsList.ToArray();
            //I don't know why this is here, but if we remove, lots of errors happen.
            WriteObject(functionalGroups);
        }
    }

    [Cmdlet(VerbsCommon.Get, "Submittals")]
    public class GetSubmittals : PSCmdlet
    {
        // Declare the parameters for the cmdlet.
        [Parameter]
        public string FolderPath
        {
            get { return folderPath; }
            set { folderPath = value; }
        }
        private string folderPath;
        [Parameter]
        public List<bool> Submittals
        {
            get { return submittals; }
            set { submittals = value; }
        }
        private List<bool> submittals;

        protected override void EndProcessing()
        {
            string[] percentages = { "15%", "30%", "40%", "50%", "60%", "70%", "80%", "90%", "95%", "100%" };
            int index = 0;
            bool submittalsExist = false;
            foreach (bool submittal in submittals)
            {
                if (submittal)
                {
                    submittalsExist = true;
                }
            }
            if (submittalsExist)
            {
                //remove existing submittals
                using (var ps = PowerShell.Create(RunspaceMode.CurrentRunspace))
                {
                    //ps.AddCommand("Get-PWDocumentsBySearch").AddParameter("FolderPath", folderPath);
                    //ps.AddStatement().AddCommand("Remove-PWDocuments");

                    ps.AddCommand("Get-PWDocumentsBySearch").AddParameter("FolderPath", folderPath);
                    ps.AddCommand("Remove-PWDocuments");
                    ps.Invoke();
                    if (ps.HadErrors)
                    {
                        WriteError(new ErrorRecord(ps.Streams.Error[0].Exception,
                            "Error removing documents", ErrorCategory.NotSpecified, null));
                    }
                }
                //add new submittals

                foreach (bool submittal in submittals)
                {
                    if (submittal)
                    {
                        using (var ps = PowerShell.Create(RunspaceMode.CurrentRunspace))
                        {
                            Hashtable attributePair = new Hashtable();
                            attributePair.Add("MilestonePercent", percentages[index]);
                            ps.AddCommand("New-PWDocumentAbstract").AddParameter("FolderPath", folderPath);
                            //ps.Invoke();
                            ps.AddCommand("Update-PWDocumentAttributes").AddParameter("Attributes", attributePair);
                            ps.Invoke();
                            if (ps.HadErrors)
                            {
                                WriteError(new ErrorRecord(ps.Streams.Error[0].Exception,
                                "Error Adding document for" + percentages[index], ErrorCategory.NotSpecified, null));
                            }
                        }
                        index++;
                    }
                }
            }
        }
    }

    [Cmdlet(VerbsCommon.New, "HDriveFolders")]
    public class NewHDriveFolders : Cmdlet
    {
        [Parameter]
        public string TopLevel
        {
            get { return topLevel; }
            set { topLevel = value; }
        }
        private string topLevel;

        [Parameter]
        public string ProjectName
        {
            get { return projectName; }
            set { projectName = value; }
        }
        private string projectName;

        [Parameter]
        public string[] FunctionalGroups
        {
            get { return functionalGroups; }
            set { functionalGroups = value; }
        }
        private string[] functionalGroups;

        protected override void ProcessRecord()
        {
            string targetPath = "H:\\" + topLevel + "\\" + projectName;
            if (!Directory.Exists(targetPath))
            {
                Directory.CreateDirectory(targetPath);
                foreach (string groupName in FunctionalGroups)
                {
                    if (groupName == "TrafficStudies")
                    {
                        Directory.CreateDirectory(targetPath + "\\TrafficStudies");
                    }
                    if (groupName == "TrafficEng")
                    {
                        Directory.CreateDirectory(targetPath + "\\TrafficEng");
                    }
                }
            }
            else
            {
                WriteObject("H drive folder already exists");
            }
        }
    }

    [Cmdlet(VerbsCommon.New, "PWWebProject")]
    public class NewPWWebProject : Cmdlet
    {
        [Parameter]
        public string URN
        {
            get { return urn; }
            set { urn = value; }
        }
        private string urn;

        [Parameter]
        public Object OIDC
        {
            get { return oidc; }
            set { oidc = value; }
        }
        private Object oidc;

        [Parameter]
        public string ProjectName
        {
            get { return projectName; }
            set { projectName = value; }
        }
        private string projectName;

        [Parameter]
        public string ProjectDescription
        {
            get { return projectDescription; }
            set { projectDescription = value; }
        }
        private string projectDescription;

        protected override void EndProcessing()
        {
            Hashtable attributes = new Hashtable();
            attributes.Add("CloudProjectName", projectDescription);
            attributes.Add("CloudProjectNumber", projectName);
            attributes.Add("OIDCToken", oidc);
            attributes.Add("TemplateName", "Lower Savannah Council of Governments Best Friend Express (BFE) Transit Improvement Study");
            attributes.Add("TimeZone", "Central Standard Time");
            attributes.Add("WSGURL", "https://srf-pw-ws.bentley.com/ws");
            using (var ps = PowerShell.Create(RunspaceMode.CurrentRunspace))
            {
                ps.AddCommand("Get-PWFolderByURN").AddParameter("URN", urn);
                ps.AddCommand("Add-CloudProject").AddParameters(attributes).AddParameter("SetDriveSync").AddParameter("SetUserSync").AddParameter("SetAutoUserSync")
                    .AddParameter("UseOrganizationTemplates");
                ps.AddCommand("Out-Null")
                //ps.AddCommand("Open-CloudProject");
                ps.Invoke();
            }
        }
    }

    [Cmdlet(VerbsCommon.Rename, "PlotFiles")]
    public class RenamePlotFiles : PSCmdlet
    {
        // Declare the parameters for the cmdlet.
        [Parameter]
        public string TopLevel
        {
            get { return topLevel; }
            set { topLevel = value; }
        }
        private string topLevel;
        [Parameter]
        public string Template
        {
            get { return template; }
            set { template = value; }
        }
        private string template;
        [Parameter]
        public string ProjectName
        {
            get { return projectName; }
            set { projectName = value; }
        }
        private string projectName;
        [Parameter]
        public string CadSoftware
        {
            get { return cadSoftware; }
            set { cadSoftware = value; }
        }
        private string cadSoftware;
        [Parameter(Mandatory = false)]
        public string NDProjectNumber
        {
            get { return ndProjectNumber; }
            set { ndProjectNumber = value; }
        }
        private string ndProjectNumber;


        protected override void EndProcessing()
        {
            if (template.Equals("SRF_Pilot_Project") && (cadSoftware.Contains("SS10") || cadSoftware.Contains("ORD")))
            {
                string plotFilesPath = topLevel + "\\" + projectName + "\\TechData\\CADDesign\\CADMgmt\\PlotFiles\\";
                string plotSetCategoriesPath = plotFilesPath + "PlotSetCategories\\";
                string[] fileNames = { "xxxxx-plot.tbl", "xxxxx-plot-BW.tbl", "xxxxx-plot-color.tbl", "xxxxx Bridge Plans Sheets_Print_Organizer.xlsx", "xxxxxPlan-Sheets_Print_Organizer.xlsx" };
                foreach (string file in fileNames)
                {
                    string newFile = file.Replace("xxxxx", projectName);
                    using (var ps = PowerShell.Create(RunspaceMode.CurrentRunspace))
                    {
                        ps.AddCommand("Get-PWDocumentsBySearch").AddParameter("FolderPath", plotFilesPath).AddParameter("DocumentName", file);
                        ps.AddCommand("Rename-PWDocument").AddParameter("DocumentNewName", newFile).AddParameter("RenameFile", true).Invoke();
                    }
                }
                Hashtable attributes = new Hashtable();
                attributes.Add("PW_Filter2", projectName);
                using (var ps = PowerShell.Create(RunspaceMode.CurrentRunspace))
                {
                    ps.AddCommand("Get-PWDocumentsBySearch").AddParameter("FolderPath", plotSetCategoriesPath);
                    ps.AddCommand("Update-PWDocumentAttributes").AddParameter("Attributes", attributes).Invoke();
                }
            }
            else if (template.Equals("SRF_ND_Pilot_Project") && (cadSoftware.Contains("SS10") || cadSoftware.Contains("ORD")))
            {
                string plotFilesPath = topLevel + "\\" + projectName + "\\" + ndProjectNumber + "\\TechData\\Design\\CADMgmt\\PlotFiles\\";
                string plotSetCategoriesPath = plotFilesPath + "PlotSetCategories\\";
                string[] fileNames = { "xxxxx-plot.tbl", "xxxxx-plot-BW.tbl", "xxxxx-plot-color.tbl", "xxxxx Bridge Plans Sheets_Print_Organizer.xlsx", "xxxxxPlan-Sheets_Print_Organizer.xlsx" };
                foreach (string file in fileNames)
                {
                    string newFile = file.Replace("xxxxx", projectName);
                    using (var ps = PowerShell.Create(RunspaceMode.CurrentRunspace))
                    {
                        ps.AddCommand("Get-PWDocumentsBySearch").AddParameter("FolderPath", plotFilesPath).AddParameter("DocumentName", file);
                        ps.AddCommand("Rename-PWDocument").AddParameter("DocumentNewName", newFile).AddParameter("RenameFile", true).Invoke();
                    }
                }
                Hashtable attributes = new Hashtable();
                attributes.Add("PW_Filter2", projectName);
                using (var ps = PowerShell.Create(RunspaceMode.CurrentRunspace))
                {
                    ps.AddCommand("Get-PWDocumentsBySearch").AddParameter("FolderPath", plotSetCategoriesPath);
                    ps.AddCommand("Update-PWDocumentAttributes").AddParameter("Attributes", attributes).Invoke();
                }
            }
        }
    }
}
