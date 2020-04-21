//------------------------------------------------------------------------------
// <copyright file="OpenBinFolderPackage.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;
using EnvDTE80;
using EnvDTE;
using System.IO;
using OpenFolderExtension;
using Microsoft.VisualStudio.Shell.Interop;

namespace OpenBinFolder
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "1.0")] // Info on this package for Help/About
    [Guid(OpenBinFolderPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    public sealed class OpenBinFolderPackage : AsyncPackage
    {
        /// <summary>
        /// OpenBinFolderPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "21572e2d-a591-4cd1-b073-c4ae5e3f6be6";

        /// <summary>
        /// Initializes a new instance of the <see cref="OpenBinFolderPackage"/> class.
        /// </summary>
        public OpenBinFolderPackage()
        {

            if (this.GetService(typeof(IMenuCommandService)) is OleMenuCommandService commandService)
            {
                var menuCommandID = new CommandID(Guid.Parse("{02AB237F-F580-4278-A02B-8DA88483528E}"), int.Parse("3B9ACA01", System.Globalization.NumberStyles.HexNumber));
                var menuItem = new MenuCommand(OpenBinFolderWithFileExplorer, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        private void OpenBinFolderWithFileExplorer(object sender, EventArgs e)
        {
            var dte =(DTE2) this.GetService(typeof(SDTE));
            Array _activeProjects = (Array)dte.ActiveSolutionProjects;
            var folders = new Folders();
            foreach (var _activeProject in _activeProjects)
            {
                var path = folders.GetOutputPath((Project)_activeProject);
                if (string.IsNullOrWhiteSpace(path))
                {
                    return;
                }

                if (Directory.Exists(path) == false)
                {
                    Directory.CreateDirectory(path);
                }

                System.Diagnostics.Process.Start("explorer.exe", "\"" + path + "\"");
            }
        }
    }
}
