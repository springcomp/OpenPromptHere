using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using EnvDTE;
using EnvDTE80;
using Microsoft.Build.Evaluation;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using OpenPromptHere.Utils;
using Project = EnvDTE.Project;

namespace OpenPromptHere.Commands
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class OpenPromptCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("c340ebab-4672-4655-bbfa-5282d5be37b0");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package_;

        /// <summary>
        /// Initializes a new instance of the <see cref="OpenPromptCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private OpenPromptCommand(Package package)
        {
            if (package == null)
                throw new ArgumentNullException(nameof(package));

            package_ = package;

            var commandService = ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandId = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(MenuItemCallback, menuCommandId);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static OpenPromptCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider => this.package_;

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new OpenPromptCommand(package);
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
            // retrieve currently selected project

            var project = SolutionExplorer.GetSelectedProject();
            var path = project?.FullName;
            if (String.IsNullOrEmpty(path))
                return;

            // retrieve current project configuration (Debug|Release) and platform (x86|x64|AnyCPU)

            var configuration = project.ConfigurationManager.ActiveConfiguration;

            // retrieve current project target path, based upon the currently selected configuration

            var configurationName = configuration.ConfigurationName;
            var platformName = configuration.PlatformName;

            var targetFolder = GetTargetFolder(path, configurationName, platformName);

            // remove space from Visual Studio's "Any CPU" platform
            // so as to be mapped to MSBuild's "AnyCPU" platform.

            if (targetFolder.Length == 0 && platformName == "Any CPU")
                targetFolder = GetTargetFolder(path, configurationName, "AnyCPU");

            // run PowerShell prompt at the target location

            const string program = @"%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe";
            const string arguments = "/nologo /noexit /encodedCommand {0}";

            var command = $"Set-Location \"{targetFolder}\"";
            var encodedCommand = Convert.ToBase64String(Encoding.Unicode.GetBytes(command));

            var prog = Environment.ExpandEnvironmentVariables(program);
            var options = String.Format(arguments, encodedCommand);

            var process = new System.Diagnostics.Process();
            process.StartInfo.FileName = prog;
            process.StartInfo.Arguments = options;
            process.StartInfo.WorkingDirectory = targetFolder;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.CreateNoWindow = false;
            process.Start();
        }

        private static string GetTargetFolder(string path, string configurationName, string platformName)
        {
            var targetPath = MsBuild.GetTargetPath(path, configurationName, platformName);
            var targetFolder = Path.GetDirectoryName(targetPath);
            return targetFolder;
        }
    }
}