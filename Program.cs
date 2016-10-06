using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;

namespace CopyFindReferencesToClipboard
{
    // After compiling this code into an Exe, add a menu item (TOOLS -> Copy 'Find references' to clipboard) in Visual Studio by:
    // 1) TOOLS -> External Tools...
    //      (Note: in the 'Menu contents:' list, count which item the new item is, starting at base-1).
    //      Title: Copy 'Find references' to clipboard
    //      Command: C:\<Path>\CopyFindReferencesToClipboard.exe
    // 2) TOOLS -> Customize... -> Keyboard... (button)
    //      Show Commands Containing: tools.externalcommand
    //      Then select the n'th one, where n is the count from step 1).
    static class Program
    {
        [STAThread]
        static void Main(String[] args)
        {
            String className = "LiteTreeView32";

            DateTime startTime = DateTime.Now;
            Data data = new Data() { className = className };

            Thread t = new Thread((o) => { GetText((Data)o); });
            t.IsBackground = true;
            t.Start(data);

            lock (data)
            {
                Monitor.Wait(data);
            }

            if (data.p == null || data.p.MainWindowHandle == IntPtr.Zero)
            {
                System.Windows.Forms.MessageBox.Show("Cannot find Microsoft Visual Studio process.");
                return;
            }

            SimpleWindow owner = new SimpleWindow { Handle = data.MainWindowHandle };

            if (data.appRoot == null)
            {
                System.Windows.Forms.MessageBox.Show(owner, "Cannot find AutomationElement from process MainWindowHandle: " + data.MainWindowHandle);
                return;
            }

            String text = data.text;
            if (text.Length == 0)
            { // otherwise Clipboard.SetText throws exception
                return;
            }

            text = "Type\tName\tSub-node\tSub-node name\tLine numbers\tExtra info\n" + text;

            text = text.Replace("dynamics://", "\t");
            text = text.Replace(" - ", "\t");
            text = text.Replace(" : ", "\t");
            text = text.Replace("/", "\t");
            text = text.Replace("?DataField", "");

            System.Windows.Forms.Clipboard.SetText(text);

            String msg = "References have been copied to the clipboard.";
            var icon = System.Windows.Forms.MessageBoxIcon.None;

            if (data.lines != data.count)
            {
                msg = String.Format("Only {0} of {1} references have been copied to the clipboard.", data.lines, data.count);
                icon = System.Windows.Forms.MessageBoxIcon.Error;
            }

            System.Windows.Forms.MessageBox.Show(owner, msg, "", System.Windows.Forms.MessageBoxButtons.OK, icon);
        }

        private class SimpleWindow : System.Windows.Forms.IWin32Window
        {
            public IntPtr Handle { get; set; }
        }

        private const int TVM_GETCOUNT = 0x1100 + 5;

        [DllImport("user32.dll")]
        static extern int SendMessage(IntPtr hWnd, int msg, int wparam, int lparam);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int Width, int Height, bool Repaint);

        private class Data
        {
            public int lines = 0;
            public int count = 0;
            public IntPtr MainWindowHandle = IntPtr.Zero;
            public IntPtr TreeViewHandle = IntPtr.Zero;
            public Process p;
            public AutomationElement appRoot = null;
            public String text = null;
            public String className = null;
        }

        private static void GetText(Data data)
        {
            Process p = GetParentProcess();
            data.p = p;

            if (p == null || p.MainWindowHandle == IntPtr.Zero)
            {
                data.text = "";
                lock (data) { Monitor.Pulse(data); }
                return;
            }

            data.MainWindowHandle = p.MainWindowHandle;
            AutomationElement appRoot = AutomationElement.FromHandle(p.MainWindowHandle);
            data.appRoot = appRoot;

            if (appRoot == null)
            {
                data.text = "";
                lock (data) { Monitor.Pulse(data); }
                return;
            }

            AutomationElement treeView = appRoot.FindFirst(TreeScope.Subtree, new PropertyCondition(AutomationElement.ClassNameProperty, data.className));
            if (treeView == null)
            {
                data.text = "";
                lock (data) { Monitor.Pulse(data); }
                return;
            }

            data.TreeViewHandle = new IntPtr(treeView.Current.NativeWindowHandle);
            data.count = SendMessage(data.TreeViewHandle, TVM_GETCOUNT, 0, 0);

            // making the window really large makes it so less calls to FindAll are required
            MoveWindow(data.TreeViewHandle, 0, 0, 800, 32767, false);
            int TV_FIRST = 0x1100;
            int TVM_SELECTITEM = (TV_FIRST + 11);
            int TVGN_CARET = TVGN_CARET = 0x9;

            // if a vertical scrollbar is detected, then scroll to the top sending a TVM_SELECTITEM command
            var vbar = treeView.FindFirst(TreeScope.Subtree, new PropertyCondition(AutomationElement.NameProperty, "Vertical Scroll Bar"));
            if (vbar != null)
            {
                SendMessage(data.TreeViewHandle, TVM_SELECTITEM, TVGN_CARET, 0); // select the first item
            }

            StringBuilder sb = new StringBuilder();
            Hashtable ht = new Hashtable();

            int chunk = 0;
            while (true)
            {
                bool foundNew = false;

                var treeViewItems = treeView.FindAll(TreeScope.Subtree, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem));
                if (treeViewItems.Count == 0)
                    break;

                if (ht.Count == 0)
                {
                    chunk = treeViewItems.Count - 1;
                }

                foreach (AutomationElement ele in treeViewItems)
                {
                    try
                    {
                        String n = ele.Current.Name;
                        if (!ht.ContainsKey(n))
                        {
                            ht[n] = n;
                            foundNew = true;
                            data.lines++;
                            sb.AppendLine(n);
                        }
                    }
                    catch { }
                }

                if (!foundNew || data.lines == data.count)
                    break;

                int x = Math.Min(data.count - 1, data.lines + chunk);
                SendMessage(data.TreeViewHandle, TVM_SELECTITEM, TVGN_CARET, x);
            }

            data.text = sb.ToString();
            lock (data) { Monitor.Pulse(data); }
        }

        // this program expects to be launched from Visual Studio
        // alternative approach is to look for "Microsoft Visual Studio" in main window title
        // but there could be multiple instances running.
        private static Process GetParentProcess()
        {
            // from thread: http://stackoverflow.com/questions/2531837/how-can-i-get-the-pid-of-the-parent-process-of-my-application
            var myId = Process.GetCurrentProcess().Id;
            var query = string.Format("SELECT ParentProcessId FROM Win32_Process WHERE ProcessId = {0}", myId);
            var search = new ManagementObjectSearcher("root\\CIMV2", query);
            var results = search.Get().GetEnumerator();
            if (!results.MoveNext()) return null;
            var queryObj = results.Current;
            uint parentId = (uint)queryObj["ParentProcessId"];
            return Process.GetProcessById((int)parentId);
        }
    }
}


