using System;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Threading;
using System.Runtime.InteropServices;

namespace pdf2wordtest
{
    internal class Program
    {

        #region 操作word等app需要使用的引用的声明
        //由于使用的是COM库，因此有许多变量需要用Missing.Value代替
        //表示对 NULL 的引用
        public static Object Nothing = System.Reflection.Missing.Value;
        //表示传递 true 的引用
        public static Object refTrue = true;
        //表示传递 false 的引用
        public static Object refFalse = false;
        #endregion

        //是否退出监测弹出窗口的线程的标志
        public static bool bIsExit = false;
        //监测弹出窗口的线程
        public static Thread closeDialogThread = null;

        #region 用于关闭word等弹出窗体的方法所引用的dll

        //static System.Threading.TimerCallback timerCallback = new System.Threading.TimerCallback(closeWindow);
        //public static System.Threading.Timer timer1 = new System.Threading.Timer(timerCallback);
        // For Windows Mobile, replace user32.dll with coredll.dll
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        // Find window by Caption only. Note you must pass IntPtr.Zero as the first parameter.

        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        static extern IntPtr FindWindowByCaption(IntPtr ZeroOnly, string lpWindowName);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32.dll", SetLastError = true)]
        static extern bool PostMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);

        static uint WM_CLOSE = 0x10;

        #endregion

        static void Main(string[] args)
        {

            //Console.WriteLine("args[0]="+args[0]+", args[1]="+args[1]+", args[2]="+args[2]);
            if ((args != null) && (args.Length < 3))
            {
                Console.WriteLine("Error: please pass at least 3 paramas.");
                return;
            }

            closeDialogThread = new Thread(new ThreadStart(closeWindowThreadStart));
            closeDialogThread.Start();


            switch (args[0])
            {
                //word2pdf
                case "WORD":
                    Console.WriteLine(convertWord2Pdf(args[1], args[2]));
                    break;
                case "POWERPOINT":
                    Console.WriteLine(convertPPT2Pdf(args[1], args[2]));
                    break;
                case "PDF":
                    Console.WriteLine(convertPdf2Word(args[1], args[2]));
                    break;

                default:
                    Console.WriteLine("Error: not implemented yet.");
                    break;
            }
            //Close Window thread is should exit
            bIsExit = true;
            // Console.WriteLine(convertPPT2Pdf("D:/print/testFiles/test2pdf.pptx", "D:/print/testFiles/test2pdf-ppt.pdf"));
            // Console.WriteLine(convertPPT2Pdf("D:/print/testFiles/test2pdf.pptx", "D:/print/testFiles/test2pdf-ppt2.pdf"));
            // Console.ReadKey();
        }

        #region 关闭窗体的代码（主要用来关闭word等app的弹出窗口，解决无法关闭app的问题）
        /// <summary>
        /// 关闭弹出窗口的线程的启动函数
        /// </summary>
        static void closeWindowThreadStart()
        {
            while (!bIsExit)
            {
                try
                {
                    //关闭word的已经损坏的文件的弹出窗口
                    closeWindow("显示修复", "bosa_sdm_msword");
                }
                catch (Exception e)
                {
                    //Console.WriteLine(e.Message);
                }
                finally
                {
                    Thread.Sleep(1000);
                }
            }
        }
        /// <summary>
        /// 关闭窗体
        /// </summary>
        /// <param name="caption">窗体标题</param>
        /// <param name="className">窗体类名</param>
        public static void closeWindow(String caption, String className)
        {

            // the caption and the className is for the Word -> File -> Options window
            // the caption and the className are got by using spy++ application and focussing on the window we are researching.
            //caption = "显示修复";
            //className = "bosa_sdm_msword";
            IntPtr hWnd = (IntPtr)(0);

            // Win 32 API being called through PInvoke
            hWnd = FindWindow(className, caption);

            /*bool retVal = false;
            if ((int)hWnd != 0)
            {
               // Win 32 API being called through PInvoke
              retVal = SetForegroundWindow(hWnd);
            }*/



            if ((int)hWnd != 0)
            {
                //Console.WriteLine("got ya!");
                CloseWindow2(hWnd);
                //Console.WriteLine("close over! exiting...");

                //CloseWindow(hWnd); // either sendMessage or PostMessage can be used.
            }
        }

        public static bool CloseWindow(IntPtr hWnd)
        {
            // Win 32 API being called through PInvoke
            SendMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            return true;
        }

        public static bool CloseWindow2(IntPtr hWnd)
        {
            // Win 32 API being called through PInvoke
            PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            return true;

        }
        #endregion


        /// <summary>
        /// word转换成pdf
        /// </summary>
        /// <param name="srcPath">源文件地址</param>
        /// <param name="destPath">目标件地址</param>
        /// <returns>True表示转换成功，False表示转换失败</returns>
        public static bool convertWord2Pdf(string srcPath, String destPath)
        {
            bool bIsOK = false;
            Microsoft.Office.Interop.Word.Application app = null;
            Document doc = null;
            try
            {
                destPath = destPath.Replace('/', '\\');
                app = new Microsoft.Office.Interop.Word.Application();
                app.Visible = false;
                object path = srcPath;

                //打开并尝试修复文档
                app.Documents.Open(ref path,
                ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                ref Nothing, ref refTrue, ref Nothing, ref refTrue, ref Nothing);
                doc = app.ActiveDocument;

                doc.SaveAs2(FileName: destPath, FileFormat: WdSaveFormat.wdFormatPDF);
                Console.WriteLine("True");
                Console.WriteLine(doc.ComputeStatistics(WdStatistic.wdStatisticPages));

                bIsOK = true;
            }
            catch (Exception e) { Console.WriteLine(e.Message); }
            finally
            {
                try
                {
                    if (doc != null)
                    {
                        doc.Close(ref refFalse, ref Nothing, ref Nothing);
                    }
                }
                catch (Exception e) { }
                try
                {
                    if (app != null)
                    {
                        app.Application.Quit(ref refFalse, ref Nothing, ref Nothing);
                        app.Quit(ref refFalse, ref Nothing, ref Nothing);
                    }
                }
                catch (Exception e) { }
            }
            return bIsOK;
        }



        /// <summary>
        /// pdf转换成word
        /// </summary>
        /// <param name="srcPath">源文件地址</param>
        /// <param name="destPath">目标件地址</param>
        /// <returns>True表示转换成功，False表示转换失败</returns>
        public static bool convertPdf2Word(string srcPath, String destPath)
        {
            bool bIsOK = false;
            Microsoft.Office.Interop.Word.Application app = null;
            Document doc = null;
            try
            {
                destPath = destPath.Replace('/', '\\');
                app = new Microsoft.Office.Interop.Word.Application();
                //app.Visible = false;
                object path = srcPath;
                foreach (Microsoft.Office.Interop.Word.FileConverter con in app.FileConverters) { 
                    Console.WriteLine(con.FormatName+"\t"+con.ClassName.ToString());
                }

                Microsoft.Office.Interop.Word.FileConverter converter = app.FileConverters[4];

                //For checking format name
                Console.WriteLine("File format : "+ converter.FormatName);

                //app.Open("PDF file path", Format: converter.OpenFormat);
                app.Documents.Open(ref path, Format: converter.OpenFormat);

                //打开并尝试修复文档
                //app.Documents.Open(ref path,
                //ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                //ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                //ref Nothing, ref refTrue, ref Nothing, ref refTrue, ref Nothing);

                doc = app.ActiveDocument;
                Console.WriteLine("content: "+ doc.Content.Text);
                doc.SaveAs2(FileName: destPath, FileFormat: WdSaveFormat.wdFormatDocumentDefault);
                //doc.SaveAs2(FileName: destPath, FileFormat: WdSaveFormat.wdFormatPDF);
                Console.WriteLine("True");
                Console.WriteLine(doc.ComputeStatistics(WdStatistic.wdStatisticPages));

                bIsOK = true;
            }
            catch (Exception e) { Console.WriteLine(e.Message); }
            finally
            {
                try
                {
                    if (doc != null)
                    {
                        doc.Close(ref refFalse, ref Nothing, ref Nothing);
                    }
                }
                catch (Exception e) { }
                try
                {
                    if (app != null)
                    {
                        app.Application.Quit(ref refFalse, ref Nothing, ref Nothing);
                        app.Quit(ref refFalse, ref Nothing, ref Nothing);
                    }
                }
                catch (Exception e) { }
            }
            return bIsOK;
        }






        #region -　打开并修复文档（已弃用） -


        public static bool OpenAndRepair(string filePath, bool isVisible, String destPath)
        {
            Microsoft.Office.Interop.Word.Application oWord = null;
            Microsoft.Office.Interop.Word.Document oDoc = null;
            try
            {
                oWord = new Microsoft.Office.Interop.Word.Application();
                oWord.Visible = isVisible;
                object path = filePath;

                oDoc = oWord.Documents.Open(ref path,
                 ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                 ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                 ref Nothing, ref refTrue, ref Nothing, ref refTrue, ref Nothing);
                Console.WriteLine("dialotgs:" + oWord.Dialogs.Count);
                //oWord.Dialogs[WdWordDialog.wdDialogShowRepairs]
                oDoc.SaveAs2(FileName: destPath, FileFormat: WdSaveFormat.wdFormatPDF);
                Console.WriteLine("True");
                Console.WriteLine(oDoc.ComputeStatistics(WdStatistic.wdStatisticPages));

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
            finally
            {
                try
                {
                    if (oDoc != null)
                    {
                        oDoc.Close(ref refFalse, ref Nothing, ref Nothing);
                    }
                }
                catch (Exception e) { }
                try
                {
                    if (oWord != null)
                    {
                        oWord.Application.Quit(ref refFalse, ref Nothing, ref Nothing);
                        oWord.Quit(ref refFalse, ref Nothing, ref Nothing);
                    }
                }
                catch (Exception e) { }

            }
            return true;
        }



        #endregion

        /// <summary>
        /// ppt转换成pdf
        /// </summary>
        /// <param name="srcPath">源文件地址</param>
        /// <param name="destPath">目标件地址</param>
        /// <returns>True表示转换成功，False表示转换失败</returns>
        public static bool convertPPT2Pdf(String srcPath, String destPath)
        {
            bool bIsOK = false;
            try
            {
                destPath = destPath.Replace('/', '\\');

                Microsoft.Office.Interop.PowerPoint.Application app = new Microsoft.Office.Interop.PowerPoint.Application();
                app.Presentations.Open(srcPath);
                Presentation presentation = app.ActivePresentation;
                presentation.SaveAs(FileName: destPath, FileFormat: PpSaveAsFileType.ppSaveAsPDF, MsoTriState.msoTriStateMixed);
                app.Quit();
                bIsOK = true;
            }
            catch (Exception e) { Console.WriteLine(e.Message); }
            return bIsOK;
        }
    }
}
