using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using OpenPromptHere.Utils;
using System.Text;
using System.IO;
using EnvDTE;
using EnvDTE80;

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
        public const int CommandId = PackageIds.OpenPromptCommandId;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = PackageGuids.guidOpenPromptCommandPackageCmdSet;

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
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(MenuItemCallback, menuCommandID);
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
        private IServiceProvider ServiceProvider => package_;

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
            var project = SolutionExplorer.GetSelectedProject();
            var path = project?.FullName;
            if (String.IsNullOrEmpty(path))
                return;

            // retrieve current project configuration (Debug|Release) and platform (x86|x64|AnyCPU)

            var configuration = project.ConfigurationManager.ActiveConfiguration;

            // retrieve current project target path, based upon the currently selected configuration

            var configurationName = configuration.ConfigurationName;
            var platformName = configuration.PlatformName;

            string targetFolder;
            if (project.IsNetFrameworkProject())
                targetFolder = GetNetFrameworkProjectPath(path, configurationName, platformName);
            else if (project.IsNetCoreProject())
                targetFolder = GetNetCoreProjectPath(path, configurationName, platformName);
            else
                throw new NotSupportedException("Unsupported Visual Studio project type.");

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

        private static string GetNetCoreProjectPath(string path, string configurationName, string platformName)
        {
            var targetFrameworks = MsBuild.GetTargetFrameworks(path, configurationName);
            // WARNING: currently supporting only the "first" target framework

            if (targetFrameworks == null)
                targetFrameworks = "netcoreapp2.1";

            var frameworkFragments = targetFrameworks.Split(';');
            var targetFramework = frameworkFragments.Length > 0
                    ? frameworkFragments[0]
                    : targetFrameworks
                ;

            var runtimeIdentifiers = MsBuild.GetRuntimeIdentifiers(path, configurationName);
            if (runtimeIdentifiers != null)
            {
                // WARNING: currently supporting only the "first" runtime identifier
                var fragments = runtimeIdentifiers.Split(';');
                var runtime = fragments.Length > 0
                        ? fragments[0]
                        : runtimeIdentifiers
                    ;
            }

            if (platformName == "Any CPU")
                platformName = "AnyCPU";

            var outputPath = MsBuild.GetOutputPath(path, configurationName, targetFramework, platformName);
            var directory = Path.GetDirectoryName(path);

            var targetFolder = Path.Combine(
                directory,
                outputPath
            );

            return targetFolder;
        }


        private static string GetNetFrameworkProjectPath(string path, string configurationName, string platformName)
        {
            var targetFolder = GetTargetFolder(path, configurationName, platformName);

            // remove space from Visual Studio's "Any CPU" platform
            // so as to be mapped to MSBuild's "AnyCPU" platform.

            if (targetFolder.Length == 0 && platformName == "Any CPU")
                targetFolder = GetTargetFolder(path, configurationName, "AnyCPU");
            return targetFolder;
        }

        private static string GetTargetFolder(string path, string configurationName, string platformName)
        {
            var targetPath = MsBuild.GetTargetPath(path, configurationName, platformName);
            var targetFolder = Path.GetDirectoryName(targetPath);
            return targetFolder;
        }
    }

    internal static class EnvDteProjectExtensions
    {
        public static bool IsNetFrameworkProject(this Project project)
        {
            return IsVisualStudioProjectKind(project, ProjectKind.NetFrameworkProject);
        }

        public static bool IsNetCoreProject(this Project project)
        {
            return IsVisualStudioProjectKind(project, ProjectKind.NetCoreProject);
        }

        public static bool IsVisualStudioProjectKind(this Project project, string projectKind)
        {
            return project.Kind.ToLowerInvariant().Contains(projectKind.ToLowerInvariant());
        }
    }

    internal static class ProjectKind
    {
        public const string NetFrameworkProject = "{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}";
        public const string NetCoreProject = "{9A19103F-16F7-4668-BE54-9A1E7A4F7556}";
    }
}
