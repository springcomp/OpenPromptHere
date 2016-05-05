namespace OpenPromptHere
{
    using System;
    
    /// <summary>
    /// Helper class that exposes all GUIDs used across VS Package.
    /// </summary>
    internal sealed partial class PackageGuids
    {
        public const string guidOpenPromptCommandPackageString = "b07bb768-fef0-4c48-b3e7-16a953c0e70e";
        public const string guidOpenPromptCommandPackageCmdSetString = "c340ebab-4672-4655-bbfa-5282d5be37b0";
        public const string guidImagesString = "9d84bc07-21e6-4516-ba45-f907eb209f46";
        public static Guid guidOpenPromptCommandPackage = new Guid(guidOpenPromptCommandPackageString);
        public static Guid guidOpenPromptCommandPackageCmdSet = new Guid(guidOpenPromptCommandPackageCmdSetString);
        public static Guid guidImages = new Guid(guidImagesString);
    }
    /// <summary>
    /// Helper class that encapsulates all CommandIDs uses across VS Package.
    /// </summary>
    internal sealed partial class PackageIds
    {
        public const int MyMenuGroup = 0x1020;
        public const int OpenPromptCommandId = 0x0100;
        public const int openPromptCommand = 0x0001;
    }
}
