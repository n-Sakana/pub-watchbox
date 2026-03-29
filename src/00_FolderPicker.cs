using System;
using System.Runtime.InteropServices;

namespace MailPull
{
    public static class FolderPicker
    {
        public static string Show(string initialPath)
        {
            try
            {
                var dlg = (IFileOpenDialog)new FileOpenDialog();
                dlg.SetOptions(FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM);

                if (!string.IsNullOrEmpty(initialPath))
                {
                    IShellItem folder;
                    if (SHCreateItemFromParsingName(initialPath, IntPtr.Zero,
                        typeof(IShellItem).GUID, out folder) == 0)
                        dlg.SetFolder(folder);
                }

                if (dlg.Show(IntPtr.Zero) != 0) return null;

                IShellItem result;
                dlg.GetResult(out result);
                string path;
                result.GetDisplayName(SIGDN_FILESYSPATH, out path);
                return path;
            }
            catch { return null; }
        }

        const uint FOS_PICKFOLDERS = 0x20;
        const uint FOS_FORCEFILESYSTEM = 0x40;
        const uint SIGDN_FILESYSPATH = 0x80058000;

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        static extern int SHCreateItemFromParsingName(
            string pszPath, IntPtr pbc, [In] Guid riid, out IShellItem ppv);

        [ComImport, Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7")]
        class FileOpenDialog { }

        [ComImport, Guid("42F85136-DB7E-439C-85F1-E4075D135FC8")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IFileOpenDialog
        {
            [PreserveSig] int Show(IntPtr hwnd);
            void SetFileTypes(); void SetFileTypeIndex(); void GetFileTypeIndex();
            void Advise(); void Unadvise();
            void SetOptions(uint fos); void GetOptions();
            void SetDefaultFolder(IShellItem psi);
            void SetFolder(IShellItem psi);
            void GetFolder(); void GetCurrentSelection();
            void SetFileName(); void GetFileName();
            void SetTitle(); void SetOkButtonLabel(); void SetFileNameLabel();
            void GetResult(out IShellItem ppsi);
        }

        [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IShellItem
        {
            void BindToHandler(); void GetParent();
            void GetDisplayName(uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
        }
    }
}
