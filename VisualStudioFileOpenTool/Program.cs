using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using EnvDTE;

namespace VisualStudioFileOpenTool
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 4)
            {
                return;
            }

            string vsPath = args[0];
            string solutionPath = args[1];
            string filePath = args[2];
            int fileLine;
            int.TryParse(args[3], out fileLine);

            try
            {
                var dte = FindRunningVSProWithOurSolution(solutionPath);
                if (dte == null)
                {
                    dte = CreateNewRunningVSProWithOurSolution(vsPath, solutionPath);
                }

                HaveRunningVSProOpenFile(dte, filePath, fileLine);
            }
            catch (Exception e)
            {
                Console.Write(e.Message);
            }
        }

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        static DTE FindRunningVSProWithOurSolution(string solutionPath)
        {
            DTE dte = null;
            object runningObject = null;
            IBindCtx bindCtx = null;
            IRunningObjectTable rot = null;
            IEnumMoniker enumMonikers = null;

            try
            {
                Marshal.ThrowExceptionForHR(CreateBindCtx(reserved: 0, ppbc: out bindCtx));
                bindCtx.GetRunningObjectTable(out rot);
                rot.EnumRunning(out enumMonikers);

                IMoniker[] moniker = new IMoniker[1];
                IntPtr numberFetched = IntPtr.Zero;
                while (enumMonikers.Next(1, moniker, numberFetched) == 0)
                {
                    IMoniker runningObjectMoniker = moniker[0];
                    Marshal.ThrowExceptionForHR(rot.GetObject(runningObjectMoniker, out runningObject));
                    var dte2 = runningObject as DTE;
                    if (dte2 != null)
                    {
                        if (dte2.Solution.FullName == solutionPath)
                        {
                            dte = dte2;
                            break;
                        }
                    }
                }
            }
            finally
            {
                if (enumMonikers != null)
                {
                    Marshal.ReleaseComObject(enumMonikers);
                }

                if (rot != null)
                {
                    Marshal.ReleaseComObject(rot);
                }

                if (bindCtx != null)
                {
                    Marshal.ReleaseComObject(bindCtx);
                }
            }

            return dte;
        }

        static DTE FindRunningVSProWithOurProcess(int processId)
        {
            string progId = ":" + processId.ToString();
            DTE dte = null;
            object runningObject = null;
            IBindCtx bindCtx = null;
            IRunningObjectTable rot = null;
            IEnumMoniker enumMonikers = null;

            try
            {
                Marshal.ThrowExceptionForHR(CreateBindCtx(reserved: 0, ppbc: out bindCtx));
                bindCtx.GetRunningObjectTable(out rot);
                rot.EnumRunning(out enumMonikers);

                IMoniker[] moniker = new IMoniker[1];
                IntPtr numberFetched = IntPtr.Zero;
                while (enumMonikers.Next(1, moniker, numberFetched) == 0)
                {
                    IMoniker runningObjectMoniker = moniker[0];
                    string name = null;

                    try
                    {
                        if (runningObjectMoniker != null)
                        {
                            runningObjectMoniker.GetDisplayName(bindCtx, null, out name);
                        }
                    }
                    catch (UnauthorizedAccessException)
                    {
                        // Do nothing, there is something in the ROT that we do not have access to.
                    }

                    if (!string.IsNullOrEmpty(name) && name.Contains(progId))
                    {
                        Marshal.ThrowExceptionForHR(rot.GetObject(runningObjectMoniker, out runningObject));
                        dte = runningObject as DTE;
                        if (dte != null)
                        {
                            break;
                        }
                    }
                }
            }
            finally
            {
                if (enumMonikers != null)
                {
                    Marshal.ReleaseComObject(enumMonikers);
                }

                if (rot != null)
                {
                    Marshal.ReleaseComObject(rot);
                }

                if (bindCtx != null)
                {
                    Marshal.ReleaseComObject(bindCtx);
                }
            }

            return dte;
        }

        static DTE CreateNewRunningVSProWithOurSolution(string vsPath, string solutionPath)
        {
            if (!File.Exists(vsPath))
            {
                return null;
            }

            var devenv = System.Diagnostics.Process.Start(vsPath, solutionPath);

            DTE dte = null;
            do
            {
                System.Threading.Thread.Sleep(2000);
                dte = FindRunningVSProWithOurProcess(devenv.Id);
            }
            while (dte == null);

            do
            {
                System.Threading.Thread.Sleep(1000);
            } while (dte.ItemOperations == null);
            return dte;
        }

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        static void HaveRunningVSProOpenFile(DTE dte, string filePath, int fileLine)
        {
            if (dte == null)
            {
                return;
            }

            dte.MainWindow.Activate();
            SetForegroundWindow(new IntPtr(dte.MainWindow.HWnd));

            var window = dte.ItemOperations.OpenFile(filePath);
            var textSelection = (TextSelection)window.Selection;
            textSelection.GotoLine(fileLine, true);
            Marshal.ReleaseComObject(dte);
        }
    }
}
