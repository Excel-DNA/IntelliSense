using System;
using System.Runtime.InteropServices;

namespace ExcelDna.IntelliSense
{
    // Also note LDR_MODULE info here: http://stackoverflow.com/questions/4242469/detect-when-a-module-dll-is-unloaded
    class LoaderNotification : IDisposable
    {
        public enum Reason : uint
        {
            Loaded = 1,
            Unloaded = 2
        }

        public class NotificationEventArgs : EventArgs
        {
            public Reason Reason;
            public string FullDllName;
        }

        public event EventHandler<NotificationEventArgs> LoadNotification;

        #region PInvoke details
        // Helper for UNICODE_STRING type - couldn't figure out how to do it simply with Marshaling
        static class UnicodeString
        {
            // Layout is:

            // ushort Length;        // Bytes
            // ushort MaximumLength; // Bytes
            // IntPtr buffer;

            public static string ToString(IntPtr pUnicodeString)
            {
                 short length = (short)Marshal.PtrToStructure(pUnicodeString, typeof(short));
                 IntPtr buffer = Marshal.ReadIntPtr(pUnicodeString, 4);
                 return Marshal.PtrToStringUni(buffer, length / 2);
            }
        }

        // At the moment, _LDR_DLL_LOADED_NOTIFICATION_DATA  and _LDR_DLL_UNLOADED_NOTIFICATION_DATA
        // are the same, so we don't have to bother with the union.
        // TODO: Check 64-bit packing
        [StructLayout(LayoutKind.Sequential)]
        struct Data
        {
            public uint Flags;            // Reserved.
            public IntPtr FullDllName;    // PCUNICODE_STRING  // The full path name of the DLL module.
            public IntPtr BaseDllName;    // PCUNICODE_STRING  // The base file name of the DLL module.
            public IntPtr DllBase;         // A pointer to the base address for the DLL in memory.
            public uint SizeOfImage;      // The size of the DLL image, in bytes.
        }

        enum NtStatus : uint
        {
            // Success
            Success = 0x00000000,
            DllNotFound = 0xc0000135,
            // Many, many others...
        }

        delegate void LdrNotification(Reason notificationReason, IntPtr pNotificationData, IntPtr context);

        // Registers for notification when a DLL is first loaded. This notification occurs before dynamic linking takes place.
        [DllImport("ntdll.dll")]
        static extern uint /*NtStatus*/ LdrRegisterDllNotification(
            uint flags, // This parameter must be zero.
            LdrNotification notificationFunction, 
            IntPtr context, 
            out IntPtr cookie); 

        [DllImport("ntdll.dll")]
        static extern uint /*NtStatus*/ LdrUnregisterDllNotification(IntPtr cookie);

        #endregion

        IntPtr _cookie;
        LdrNotification _notificationDelegate;

        public LoaderNotification()
        {
            IntPtr context = IntPtr.Zero; // new IntPtr(12345);
            _notificationDelegate = Notification;
            var status = LdrRegisterDllNotification(0, _notificationDelegate, context, out _cookie);
            if (status != 0)
            {
                throw new InvalidOperationException($"Error in LdrRegisterDlLNotification. Result: {status}");
            }
        }

        // WARNING! LoaderLock danger here
        void Notification(Reason notificationReason, IntPtr pNotificationData, IntPtr context)
        {
            IntPtr pFullDllName = Marshal.ReadIntPtr(pNotificationData, 4);
            string fullDllName = UnicodeString.ToString(pFullDllName);
            NotificationEventArgs args = new NotificationEventArgs { Reason = notificationReason, FullDllName = fullDllName };
            LoadNotification?.Invoke(this, args);

            // Debug.Print($"@@@@ LdrNotification: {notificationReason} - {fullDllName}");
        }

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects).
                }

                var status = LdrUnregisterDllNotification(_cookie);
                disposedValue = true;
            }
        }

        ~LoaderNotification()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(false);
        }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
             GC.SuppressFinalize(this);
        }
        #endregion


        /*
        [StructLayout(LayoutKind.Sequential)]
        struct UnicodeString
        {
            // Helper for UNICODE_STRING type, with layout
            ushort Length;        // Bytes
            ushort MaximumLength; // Bytes
            IntPtr buffer;

            public override string ToString()
            {
                return Marshal.PtrToStringUni(buffer, Length / 2);
            }
        }

        // At the moment, _LDR_DLL_LOADED_NOTIFICATION_DATA  and _LDR_DLL_UNLOADED_NOTIFICATION_DATA
        // are the same, so we don't have to bother with the union.
        // TODO: Check 64-bit packing
        [StructLayout(LayoutKind.Sequential)]
        struct Data
        {
            public uint Flags;            // Reserved.
     ????       [MarshalAs(UnmanagedType.LPStruct, MarshalTypeRef = typeof(UnicodeString))]
            public UnicodeString FullDllName;
     ????       [MarshalAs(UnmanagedType.LPStruct, MarshalTypeRef = typeof(UnicodeString))]
            public UnicodeString BaseDllName;
            //public IntPtr FullDllName;    // PCUNICODE_STRING  // The full path name of the DLL module.
            //public IntPtr BaseDllName;    // PCUNICODE_STRING  // The base file name of the DLL module.
            public IntPtr DllBase;         // A pointer to the base address for the DLL in memory.
            public uint SizeOfImage;      // The size of the DLL image, in bytes.
        }

        delegate void LdrNotification(Reason notificationReason, [In] ref Data notificationData, IntPtr context);
*/
    }
}
