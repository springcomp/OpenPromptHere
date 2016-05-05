using System;
using System.Runtime.InteropServices;

namespace OpenPromptHere.Utils
{
    public sealed class TypedObjectForIUnknown<T> : IDisposable where T : class
    {
        private readonly IntPtr pointer_;

        public static T Get(IntPtr pointer)
        {
            return Marshal.GetTypedObjectForIUnknown(pointer, typeof(T)) as T;
        }

        private TypedObjectForIUnknown(IntPtr pointer)
        {
            pointer_ = pointer;
            Instance = Get(pointer_);
        }

        public static TypedObjectForIUnknown<T> Attach(IntPtr pointer)
        {
            return new TypedObjectForIUnknown<T>(pointer);
        }

        public T Instance { get; }

        public void Dispose()
        {
            GC.SuppressFinalize(this);
            Dispose(true);
        }

        private void Dispose(bool disposing)
        {
            if (disposing)
                Marshal.Release(pointer_);
        }
    }
}