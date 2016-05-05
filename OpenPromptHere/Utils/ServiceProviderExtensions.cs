using System;
using Microsoft.VisualStudio.Shell;

namespace OpenPromptHere.Utils
{
    public static class ServiceProviderExtensions
    {
        /// <summary>
        /// Return an instance of the specified service type.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="provider"></param>
        /// <returns></returns>
        public static T GetService<T>(this IServiceProvider provider) where T : class
        {
            return provider.GetService(typeof (T)) as T;
        }
    }

    public static class PackageExtensions
    {
        public static T GetGlobalService<T>() where T : class
        {
            return GetGlobalService<T, T>();
        }

        public static TInterface GetGlobalService<TInterface, TObject>() where TInterface: class where TObject: class
        {
            return Package.GetGlobalService(typeof (TObject)) as TInterface;
        }
    }
}