using System.Collections.Generic;
using System.Linq;
using Microsoft.Build.Evaluation;

namespace OpenPromptHere.Utils
{
    public static class MsBuild
    {
        public static string GetTargetPath(string csproj)
        {
            const string propertyName = "TargetPath";
            return GetProperty(csproj, propertyName);
        }

        public static string GetProperty(string csproj, string propertyName)
        {
            var properties = new Dictionary<string, string>();

            // load MSBuild project and evaluate its TargetPath
            // based upon the selected configuration

            using (var collection = new ProjectCollection())
            {
                var project = new Project(
                    csproj
                    , properties
                    , null
                    , collection
                    , ProjectLoadSettings.Default
                    );

                var property = project.Properties
                    .Where(p => p.Name == propertyName)
                    .Select(p => p.EvaluatedValue)
                    .SingleOrDefault()
                    ;

                return property;
            }
        }
    }
}