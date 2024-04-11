using System;
using System.Collections.Generic;
using System.Runtime;
using System.Runtime.InteropServices;

using Shell32;

using Newtonsoft.Json;
using CommandLine;
using CommandLine.Text;
using SHDocVw;
using System.IO;
using Vanara.PInvoke;
using System.Diagnostics;
using System.Threading;
using System.Text;


namespace DesktopUtility
{
    class Win32
    {
        [DllImport("user32.dll", SetLastError=true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll", SetLastError=true)]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
        [DllImport("user32.dll", SetLastError=true)]
        public static extern IntPtr GetDesktopWindow();
        [DllImport("user32.dll", CharSet=CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError=true)]
        public static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

        [DllImport("user32.dll", SetLastError=true)]
        public static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);
        [DllImport("shell32.dll", SetLastError=true)]
        public static extern int SHChangeNotify(uint eventId, uint flags, IntPtr item1, IntPtr item2);
        [DllImport("user32.dll", SetLastError=true)]
        public static extern IntPtr SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, IntPtr flags);

        [DllImport("user32.dll", SetLastError=true)]
        public static extern IntPtr UpdateWindow(IntPtr hWnd);
        [DllImport("user32.dll", SetLastError=true)]
        public static extern IntPtr InvalidateRect(IntPtr hWnd, IntPtr lpRect, int bErase);

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32.dll", SetLastError=true)]
        public static extern bool GetWindowInfo(IntPtr hWnd, ref WINDOWINFO pwi);

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32.dll", SetLastError=true)]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT windowRect);

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32.dll", SetLastError=true)]
        public static extern bool GetWindowPlacement(IntPtr hWnd, out WINDOWPLACEMENT windowPlacement);

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32.dll", SetLastError=true)]
        public static extern bool SetWindowPlacement(IntPtr hWnd, WINDOWPLACEMENT windowPlacement);

        // [return: MarshalAs(UnmanagedType.Bool)]
        // [DllImport("user32.dll", SetLastError=true)]
        // public static extern bool ;


        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public long x;
            public long y;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct WINDOWINFO
        {
            public uint cbSize;
            public RECT rcWindow;
            public RECT rcClient;
            public uint dwStyle;
            public uint dwExStyle;
            public uint dwWindowStatus;
            public uint cxWindowBorders;
            public uint cyWindowBorders;
            public ushort atomWindowType;
            public ushort wCreatorVersion;

            public WINDOWINFO(Boolean? filler) : this()
            {
                cbSize = (UInt32)(Marshal.SizeOf(typeof(WINDOWINFO)));
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct WINDOWPLACEMENT {
            public uint length;
            public uint flags;
            public uint showCmd;
            public POINT ptMinPosition;
            public POINT ptMaxPosition;
            public RECT rcNormalPosition;
            public RECT rcDevice;
        }
    }

    public static class Constants
    {
        public const Int32 WM_COMMAND = 0x0111;
        public const Int32 WS_VISIBLE = 0x10000000;
        public const Int32 SW_SHOW = 5;
        public const Int32 GW_CHILD = 5;
        public const Int32 GWL_STYLE = -16;
        public static IntPtr NULL = (IntPtr)0x0;
        public static IntPtr TRUE = (IntPtr)1;
        public static IntPtr FALSE = (IntPtr)1;
    }

    public class Desktop
    {
        public static IntPtr GetHandle()
        {
            IntPtr hDesktopWin = Win32.GetDesktopWindow();
            IntPtr hProgman = Win32.FindWindow("Progman", "Program Manager");
            IntPtr hWorkerW = Constants.NULL;

            IntPtr hShellViewWin = Win32.FindWindowEx(hProgman, Constants.NULL, "SHELLDLL_DefView", "");
            if (hShellViewWin == Constants.NULL)
            {
                do
                {
                    hWorkerW = Win32.FindWindowEx(hDesktopWin, hWorkerW, "WorkerW", "");
                    hShellViewWin = Win32.FindWindowEx(hWorkerW, Constants.NULL, "SHELLDLL_DefView", "");
                } while (hShellViewWin == Constants.NULL && hWorkerW != Constants.NULL);
            }
            return hShellViewWin;
        }

        public static void Refresh()
        {
            IntPtr hWndDesktop = Desktop.GetHandle();
            IntPtr hWndChild = Win32.GetWindow(hWndDesktop, Constants.GW_CHILD);
            // Win32Functions.SetWindowPos(hWnd, Constants.NULL, 0, 0, 0, 0, (IntPtr)(0x0002 | 0x0200 | 0x0001 | 0x0004 | 0x0010));
            // Win32Functions.SetWindowPos(shelldll_hWnd, Constants.NULL, 0, 0, 0, 0, (IntPtr)(0x0002 | 0x0200 | 0x0001 | 0x0004 | 0x0010));

            // I'm not sure why, but I have to invalidate and update for both.
            Win32.InvalidateRect(hWndDesktop, Constants.NULL, (int)Constants.TRUE);
            Win32.UpdateWindow(hWndDesktop);
            Win32.InvalidateRect(hWndChild, Constants.NULL, (int)Constants.TRUE);
            Win32.UpdateWindow(hWndChild);
        }

        public static void ToggleDesktopIcons()
        {
            Win32.SendMessage(Desktop.GetHandle(), Constants.WM_COMMAND, (IntPtr)0x7402, Constants.NULL);
        }

        public static void EnableDesktopIcons()
        {
            IntPtr hWnd = Win32.GetWindow(GetHandle(), Constants.GW_CHILD);

            Win32.WINDOWINFO info = new Win32.WINDOWINFO();
            info.cbSize = (uint)Marshal.SizeOf(info);
            Win32.GetWindowInfo(hWnd, ref info);
            Win32.SetWindowLongPtr(hWnd, Constants.GWL_STYLE, (IntPtr)(info.dwStyle | Constants.WS_VISIBLE));

            Refresh();

        }

        public static void DisableDesktopIcons()
        {
            IntPtr hWnd = Win32.GetWindow(GetHandle(), Constants.GW_CHILD);

            Win32.WINDOWINFO info = new Win32.WINDOWINFO();
            info.cbSize = (uint)Marshal.SizeOf(info);
            Win32.GetWindowInfo(hWnd, ref info);
            Win32.SetWindowLongPtr(hWnd, Constants.GWL_STYLE, (IntPtr)(info.dwStyle & ~Constants.WS_VISIBLE));

            Refresh();
        }

        public static bool GetIconsVisibility()
        {
            // "FolderView" SysListView32 gets or removes WS_VISIBLE window style based on if the icons are hidden or not.
            IntPtr hWnd = Win32.GetWindow(GetHandle(), Constants.GW_CHILD);

            Win32.WINDOWINFO info = new Win32.WINDOWINFO();
            info.cbSize = (uint)Marshal.SizeOf(info);
            Win32.GetWindowInfo(hWnd, ref info);
            return (info.dwStyle & Constants.WS_VISIBLE) == Constants.WS_VISIBLE;
        }
    }
}

namespace Program
{
    [Serializable]
    public struct SerializedWindows
    {
        public List<SerializedWindow> windows;
    }

    [Serializable]
    public struct SerializedWindow
    {
        public string Name;
        public string Path;
        public int Top;
        public int Left;
        public int Width;
        public int Height;
    }

    public class Program
    {
        public class Options
        {
            [Option('v', "verbose", Required=false, HelpText="Set output verbosity.")]
            public bool Verbose { get; set; }

            [Verb("save", HelpText="Save the list of open windows.")]
            public class SaveOptions {
                [Option('o', "output", Required=false, Default="SavedWindows.json", HelpText="Set output filename.")]
                public string Output { get; set; }
            }

            [Verb("get", HelpText="Output list of open windows")]
            public class GetOptions { }

            [Verb("restore", HelpText="Restore the saved windows by opening and respositioning them.")]
            public class RestoreOptions {
                [Option('i', "input", Required=false, Default="SavedWindows.json", HelpText="Set input filename.")]
                public string Input { get; set; }
            }

            [Usage(ApplicationAlias="todo: application name")]
            public static IEnumerable<Options> Examples
            {
                get {
                    return new List<Options>() {
                    };
                }
            }
        }

        public static Object GetShell()
        {
            Type shellAppType = Type.GetTypeFromProgID("Shell.Application");
            Object shell = Activator.CreateInstance(shellAppType);
            return shell;
        }

        public static ShellWindows GetWindows()
        {
            var shell = new Shell();
            var windows = shell.Windows() as SHDocVw.ShellWindows;
            if (windows == null) {
                // TODO: Error
            }

            return windows;
        }
        
        public static void DumpObj(Object obj)
        {
            var type = obj.GetType();
            var props = type.GetFields();
            Console.WriteLine($"Type: {type}");
            foreach (var prop in props)
            {
                var field = type.GetField(prop.Name);
                var name = field.Name;
                var value = field.GetValue(obj);
                Console.WriteLine($"\t{name}: {value}");
            }
        }

        [STAThread]
        public static void Main(string[] args)
        {
            Parser.Default.ParseArguments<Options, Options.GetOptions, Options.RestoreOptions, Options.SaveOptions>(args)
                .WithParsed<Options.GetOptions>(RunOptions)
                .WithParsed<Options.SaveOptions>(RunOptions)
                .WithParsed<Options.RestoreOptions>(RunOptions)
                .WithNotParsed(HandleParseError);
        }

        static void RunOptions(Options.GetOptions opts)
        {
            var windows = GetWindows();

            for (int j = 0; j < windows.Count; j++)
            {
                var item = windows.Item(j) as SHDocVw.InternetExplorer;
                if (item == null)
                {
                    // TODO: Error
                }
                Console.WriteLine(item.Name);
                Console.WriteLine(item.LocationURL);
            }
        }

        static void RunOptions(Options.SaveOptions saveOptions)
        {
            var savedWindows = new SerializedWindows();
            savedWindows.windows = new List<SerializedWindow>();

            var windows = GetWindows();

            for (int j = 0; j < windows.Count; j++)
            {
                var item = windows.Item(j) as SHDocVw.InternetExplorer;
                if (item == null)
                {
                    // TODO: Error
                }

                var windowPlacement = new User32.WINDOWPLACEMENT();
                User32.GetWindowPlacement(new HWND((IntPtr)item.HWND), ref windowPlacement);

                Console.WriteLine(windowPlacement.length);
                Console.WriteLine(windowPlacement.flags);
                Console.WriteLine($"left: {item.Left}, top: {item.Top} width: {item.Width}, height: {item.Height}");
                DumpObj(windowPlacement);

                var serializedWindow = new SerializedWindow();
                serializedWindow.Name   = item.LocationName;
                serializedWindow.Path   = new Uri(item.LocationURL).LocalPath;
                serializedWindow.Top    = windowPlacement.rcNormalPosition.Top;
                serializedWindow.Left   = windowPlacement.rcNormalPosition.Left;
                serializedWindow.Width  = windowPlacement.rcNormalPosition.Width;
                serializedWindow.Height = windowPlacement.rcNormalPosition.Height;

                savedWindows.windows.Add(serializedWindow);
            }

            var filename = saveOptions.Output;

            var serialized = JsonConvert.SerializeObject(savedWindows, Formatting.Indented);
            File.WriteAllText(filename, serialized);
        }

        static void RunOptions(Options.RestoreOptions resOptions)
        {
            var filename = resOptions.Input;
            // TODO: Handle file errors.
            var savedWindows = JsonConvert.DeserializeObject<SerializedWindows>(File.ReadAllText(filename));
            foreach (var win in savedWindows.windows)
            {
                var pInfo = new Kernel32.SafePROCESS_INFORMATION();

                var sInfo = new Kernel32.STARTUPINFO
                {
                    dwX = (uint)win.Left,
                    dwY = (uint)win.Top,
                    dwXSize = (uint)win.Width,
                    dwYSize = (uint)win.Height,
                    wShowWindow = 10,
                    dwFlags = Kernel32.STARTF.STARTF_USEPOSITION | Kernel32.STARTF.STARTF_USESIZE | Kernel32.STARTF.STARTF_USESHOWWINDOW
                };
                var pSec = new SECURITY_ATTRIBUTES();
                var tSec = new SECURITY_ATTRIBUTES();
                pSec.nLength = Marshal.SizeOf(pSec);
                tSec.nLength = Marshal.SizeOf(tSec);


                var result = Kernel32.CreateProcess(
                    "C:\\Windows\\explorer.exe",
                    new StringBuilder($" \"{win.Path}\""),
                    pSec,
                    tSec,
                    false,
                    Kernel32.CREATE_PROCESS.NORMAL_PRIORITY_CLASS | Kernel32.CREATE_PROCESS.CREATE_UNICODE_ENVIRONMENT,
                    null,
                    null,
                    sInfo,
                    out pInfo
                );

                Console.WriteLine($"Result: {result}");

                Console.WriteLine($"{pInfo.hProcess.DangerousGetHandle()}, {pInfo.hThread.DangerousGetHandle()}");

                Thread.Sleep(1000);

                // TODO: Need to move this logic into a "WindowMover" class that handles finding and moving the windows until they have all been found.
                foreach (InternetExplorer window in new SHDocVw.ShellWindows())
                {
                    if (Path.GetFileNameWithoutExtension(window.FullName).ToLowerInvariant() == "explorer")
                    {
                        var path1 = new Uri(window.LocationURL).LocalPath.ToLowerInvariant();
                        if (path1 == win.Path.ToLowerInvariant())
                        {
                            window.Left = win.Left;
                            window.Top = win.Top;
                            window.Width = win.Width;
                            window.Height = win.Height;
                        }
                    }
                }
            }
        }

        static void HandleParseError(IEnumerable<Error> errors)
        {
        }
    }
}
