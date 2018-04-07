//------------------------------------------------------------------------------
// <copyright file="BuildOneFile.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------


using EnvDTE;
using Microsoft.VisualStudio.VCProjectEngine;
using System;
using System.ComponentModel.Design;
using System.Collections.Generic;

using System.Globalization;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace VSNinja
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class BuildOneFile
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("2ec731ea-206b-4bf9-8472-0e3e8e650084");
        public static readonly Guid ItemContextMenuGuid = new Guid("9f67a0bd-ee0a-47e3-b656-5efb12e3c770");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        private static ErrorListProvider errorListProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="BuildOneFile"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private BuildOneFile(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        private void beforeQueryStatus(object sender, EventArgs e)
        {
            var command = sender as OleMenuCommand;
            if (command == null)
                return;

            command.Enabled = true;
            command.Visible = true;
            // @todo: hide/show menu item
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static BuildOneFile Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new BuildOneFile(package);
            errorListProvider = new ErrorListProvider(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            //var command = sender as OleMenuCommand;
            //if (command == null)
            //    return; // @todo: more commands
            var Dte = this.ServiceProvider.GetService(typeof(DTE)) as DTE;           
          
            var selectedFiles = GetSelectedFiles(Dte);
            var selPro = GetSelectedProject(Dte, selectedFiles[0]);
            RunNinja(selectedFiles[0], selPro);
        }

        public static bool RunNinja(VCFile vcFile, EnvDTE.Project pro)
        {
            if (vcFile == null || pro == null)
                return false;
                       
            PaneMessage(pro.DTE, "--- (ninja) file: " + vcFile.FullPath);
            var vcProject = (pro.Object as VCProject);
            string pdir = vcProject.ProjectDirectory;
            string pfile = vcProject.ProjectFile;
            string pname = vcProject.Name;
            string pnameFull = pro.FullName;
            VCConfiguration activeConfig = vcProject.ActiveConfiguration; // @todo: it doesn't allow access properties like outputDirectory...
            string activeConfigName = activeConfig.ConfigurationName;
            //string buildDir = activeConfig.OutputDirectory; // @todo:
            string objname = Path.ChangeExtension(vcFile.Name, ".obj");           
            string objnameFull = pname + ".dir\\" + activeConfigName + "\\" + objname;
            if (File.Exists(pdir + objnameFull))
            {
                File.Delete(pdir + objnameFull); // remove object file before compilation.
            }
            var tools = (IVCCollection)activeConfig.Tools;


            foreach (var tool in tools)
            {
                var nmakeTool = tool as VCNMakeTool;
                if (nmakeTool != null)
                {
                    string originalLine = nmakeTool.BuildCommandLine;
                    string ninjaBin = originalLine.Split(' ')[0];
                    nmakeTool.BuildCommandLine = ninjaBin + " " + objnameFull;
                    try
                    {
                        pro.DTE.Solution.SolutionBuild.BuildProject(activeConfigName, pnameFull, true);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    nmakeTool.BuildCommandLine = originalLine;
                    pro.Save();
                    return true;
                }
            }
           
            return false;
        }

        public static Project GetSelectedProject(DTE dteObject, VCFile vcFile)
        {
            if (dteObject == null)
                return null;
            Projects prjs = null;
            try
            {
                prjs = dteObject.Solution.Projects;
            }
            catch
            {
            }
            if (prjs == null || prjs.Count < 1)
                return null;

            foreach (Project proj in prjs)
            {
                var vcProject = (proj.Object as VCProject);
                if (vcProject == (vcFile.project as VCProject))
                    return proj;
            }

            return null;
        }


        public static VCFile[] GetSelectedFiles(DTE dteObject)
        {
            if (dteObject.SelectedItems.Count <= 0)
                return null;

            var items = dteObject.SelectedItems;

            var files = new VCFile[items.Count + 1];
            for (var i = 1; i <= items.Count; ++i)
            {
                var item = items.Item(i);
                if (item.ProjectItem == null)
                    continue;

                VCProjectItem vcitem;
                try
                {
                    vcitem = (VCProjectItem)item.ProjectItem.Object;
                }
                catch (Exception)
                {
                    return null;
                }

                if (vcitem.Kind == "VCFile")
                    files[i - 1] = (VCFile)vcitem;
            }
            files[items.Count] = null;
            return files;
        }

        private static OutputWindowPane wndp;

        private static OutputWindowPane GetBuildPane(OutputWindow outputWindow)
        {
            foreach (OutputWindowPane owp in outputWindow.OutputWindowPanes)
            {
                if (owp.Guid == "{1BD8A850-02D1-11D1-BEE7-00A0C913D1F8}")
                    return owp;
            }
            return null;
        }
        public static void PaneMessage(DTE dte, string str)
        {
            var wnd = (OutputWindow)dte.Windows.Item(EnvDTE.Constants.vsWindowKindOutput).Object;
            if (wndp == null)
                wndp = wnd.OutputWindowPanes.Add("VSNinja");

            wndp.OutputString(str + "\r\n");
            var buildPane = GetBuildPane(wnd);
            // show buildPane if a build is in progress
            if (/*dte.Solution.SolutionBuild.BuildState == vsBuildState.vsBuildStateInProgress &&*/ buildPane != null)
                buildPane.Activate();
        }
    }
}
