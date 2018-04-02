using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using MSWord = Microsoft.Office.Interop.Word;
using MSPowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace ESBasic.Office
{
    /// <summary>
    /// OfficeScanner 用于将word文档和ppt文档中的页转换为图片。
    /// </summary>
    public class OfficeScanner
    {
        #region Convert
        public int Convert(string filePath, string strOutputPath, string strPrefix, int nBeginPage, int nEndPage)
        {
            if (!File.Exists(filePath))
            {
                throw new Exception("File is not exist !");
            }

            string[] ary = filePath.Split('.');
            string extendName = ary[ary.Length - 1];

            if (extendName == "ppt" || extendName == "pptx")
            {
                return Scan4PowerPoint(filePath, true, strOutputPath, strPrefix, nBeginPage, nEndPage);
            }
            return 0;
        }

        public int GetInfo(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new Exception("File is not exist !");
            }

            string[] ary = filePath.Split('.');
            string extendName = ary[ary.Length - 1];
            extendName.ToLower();

            if (extendName == "ppt" || extendName == "pptx")
            {
                return Scan4PowerPoint(filePath, false);
            }
            return 0;
        }

        #region Scan4PowerPoint
        private int Scan4PowerPoint(string filePath, bool bOutputPng, string strOutputPath = "", string strPrefix = "", int nBeginPage = 0, int nEndPage = -1)
        {
            MSPowerPoint.ApplicationClass pptApplicationClass = new MSPowerPoint.ApplicationClass();

            try
            {
                MSPowerPoint.Presentation presentation = pptApplicationClass.Presentations.Open(filePath, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);
                int totalCount = presentation.Slides.Count;
                if (!bOutputPng)
                {
                    return totalCount;
                }
                int index = 0;
                foreach (MSPowerPoint.Slide slide in presentation.Slides)
                {
                    if (index >= nBeginPage && (nEndPage < 0 || index < nEndPage))
                    {
                        slide.Copy();

                        System.Windows.Forms.IDataObject data = Clipboard.GetDataObject();
                        if (data.GetDataPresent(DataFormats.MetafilePict))
                        {
                            object obj = data.GetData(DataFormats.MetafilePict);
                            Metafile metafile = MetafileHelper.GetEnhMetafileOnClipboard(IntPtr.Zero);
                            Bitmap bm = new Bitmap(metafile);
                            string strOutPutIndex = index.ToString();
                            bm.Save(strOutputPath + "\\" + strPrefix + strOutPutIndex + ".png", ImageFormat.Png);
                            Clipboard.Clear();
                        }
                    }
                    ++index;
                }

                presentation.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(presentation);

                return totalCount;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                pptApplicationClass.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(pptApplicationClass);
            }
        }
        #endregion

        #region Scan4Word
        private Bitmap[] Scan4Word(string filePath)
        {
            //复制目标文件，后续将操作副本
            string tmpFilePath = AppDomain.CurrentDomain.BaseDirectory + "\\" + Path.GetFileName(filePath) + ".tmp";
            File.Copy(filePath, tmpFilePath);

            List<Bitmap> bmList = new List<Bitmap>();
            MSWord.ApplicationClass wordApplicationClass = new MSWord.ApplicationClass();
            wordApplicationClass.Visible = false;
            object missing = System.Reflection.Missing.Value;

            try
            {
                object readOnly = false;
                object filePathObject = tmpFilePath;

                MSWord.Document document = wordApplicationClass.Documents.Open(ref filePathObject, ref missing,
                    ref readOnly, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing);

                bool finished = false;
                while (!finished)
                {
                    document.Content.CopyAsPicture(); //拷贝到粘贴板
                    System.Windows.Forms.IDataObject data = Clipboard.GetDataObject();
                    if (data.GetDataPresent(DataFormats.MetafilePict))
                    {
                        object obj = data.GetData(DataFormats.MetafilePict);
                        Metafile metafile = MetafileHelper.GetEnhMetafileOnClipboard(IntPtr.Zero); //从粘贴板获取数据
                        Bitmap bm = new Bitmap(metafile.Width, metafile.Height);
                        using (Graphics g = Graphics.FromImage(bm))
                        {
                            g.Clear(Color.White);
                            g.DrawImage(metafile, 0, 0, bm.Width, bm.Height);
                        }
                        bmList.Add(bm);
                        Clipboard.Clear();
                    }

                    object What = MSWord.WdGoToItem.wdGoToPage;
                    object Which = MSWord.WdGoToDirection.wdGoToFirst;
                    object startIndex = "1";
                    document.ActiveWindow.Selection.GoTo(ref What, ref Which, ref missing, ref startIndex); // 转到下一页
                    MSWord.Range start = document.ActiveWindow.Selection.Paragraphs[1].Range;
                    MSWord.Range end = start.GoToNext(MSWord.WdGoToItem.wdGoToPage);
                    finished = (start.Start == end.Start);
                    if (finished) //最后一页
                    {
                        end.Start = document.Content.End;
                    }

                    object oStart = start.Start;
                    object oEnd = end.Start;
                    document.Range(ref oStart, ref oEnd).Delete(ref missing, ref missing); //处理完一页，就删除一页。                    
                }

                ((MSWord._Document)document).Close(ref missing, ref missing, ref missing);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(document);

                return bmList.ToArray();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                wordApplicationClass.Quit(ref missing, ref missing, ref missing);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApplicationClass);
                File.Delete(tmpFilePath); //删除临时文件
            }
        }
        #endregion       
        #endregion
    }

    #region MetafileHelper
    /// <summary>
    /// MetafileHelper 用于从剪切板获取Metafile数据。
    /// </summary>
    internal class MetafileHelper
    {
        [DllImport("user32.dll")]
        static extern bool OpenClipboard(IntPtr hWndNewOwner);
        [DllImport("user32.dll")]
        static extern bool EmptyClipboard();
        [DllImport("user32.dll")]
        static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);
        [DllImport("user32.dll")]
        static extern IntPtr GetClipboardData(uint uFormat);
        [DllImport("user32.dll")]
        static extern bool IsClipboardFormatAvailable(uint uFormat);
        [DllImport("user32.dll")]
        static extern bool CloseClipboard();
        [DllImport("gdi32.dll")]
        static extern IntPtr CopyEnhMetaFile(IntPtr hemfSrc, IntPtr hNULL);
        [DllImport("gdi32.dll")]
        static extern IntPtr CopyEnhMetaFileA(IntPtr hemfSrc, string filename);
        [DllImport("gdi32.dll")]
        static extern bool DeleteEnhMetaFile(IntPtr hemf);

        // Metafile mf is set to a state that is not valid inside this function.
        public static bool PutEnhMetafileOnClipboard(IntPtr hWnd, Metafile mf)
        {
            bool bResult = false;
            IntPtr hEMF, hEMF2;
            hEMF = mf.GetHenhmetafile(); // invalidates mf
            if (!hEMF.Equals(new IntPtr(0)))
            {
                hEMF2 = CopyEnhMetaFile(hEMF, new IntPtr(0));
                if (!hEMF2.Equals(new IntPtr(0)))
                {
                    if (OpenClipboard(hWnd))
                    {
                        if (EmptyClipboard())
                        {
                            IntPtr hRes = SetClipboardData(14 /*CF_ENHMETAFILE*/, hEMF2);
                            bResult = hRes.Equals(hEMF2);
                            CloseClipboard();
                        }
                    }
                }
                DeleteEnhMetaFile(hEMF);
            }
            return bResult;
        }

        public static Metafile GetEnhMetafileOnClipboard(IntPtr hWnd)
        {
            if (OpenClipboard(hWnd))
            {
                uint format = 14/*CF_ENHMETAFILE*/;
                if (IsClipboardFormatAvailable(format))
                {
                    IntPtr hRes = GetClipboardData(format);

                    CloseClipboard();

                    if (!hRes.Equals(new IntPtr(0)))
                    {
                        IntPtr hEMF = CopyEnhMetaFile(hRes, new IntPtr(0));
                        Metafile mf = new Metafile(hEMF, true);
                        return mf;
                    }
                }
            }
            return null;
        }

        public static bool IsClipboardFormatAvailableENHMETAFILE(IntPtr hWnd)
        {
            if (OpenClipboard(hWnd))
            {
                uint format = 14/*CF_ENHMETAFILE*/;
                if (IsClipboardFormatAvailable(format))
                {
                    CloseClipboard();
                    return true;
                }
            }
            return false;
        }

        public static bool SaveMetafile(Metafile mf, string filename)
        {
            bool bResult = false;
            IntPtr hEMF, hEMF2;
            hEMF = mf.GetHenhmetafile(); // invalidates mf
            if (!hEMF.Equals(new IntPtr(0)))
            {
                //hEMF = CopyEnhMetaFile(hEMF, new IntPtr(0));
                hEMF2 = CopyEnhMetaFileA(hEMF, filename);
                DeleteEnhMetaFile(hEMF2);
                if (hEMF2 != IntPtr.Zero)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            return false;
        }
    }
    #endregion
}

class TaskTimerRecorder
{
    public System.Diagnostics.Stopwatch watch;
    public TaskTimerRecorder()
    {
        //计时器
        watch = new System.Diagnostics.Stopwatch();
        watch.Start();
    }
    ~TaskTimerRecorder()
    {
        watch.Stop();

        TimeSpan span = watch.Elapsed;
        Console.WriteLine("任务完成用时：" + span.TotalMilliseconds.ToString());
    }
}

namespace PPT2Img
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            TaskTimerRecorder taskTimerRecorder = new TaskTimerRecorder();

            //确认需要执行的操作
            //getinfo:获取总页数
            //convert:将指定起始页码转换到指定文件夹
            string strActionMode = "getinfo";
            //拿到输入文件名
            string strPPTFileName = "PPT路径";
            //拿到输出图片路径
            string strOutputFilePath = "图片输出路径";
            //拿到输出图片的名词前缀（后缀就是从0计数的页码）
            string strPrefix = "输出图片前缀";
            //需要输出的起始页码
            int nBeginPage = 0;
            //需要输出的结束页码(-1为全部输出)
            int nEndPage = -1;

            if (args.Length > 0)
            {
                strActionMode = args[0];
                if (args.Length > 1)
                {
                    strPPTFileName = args[1];
                }

                if (args.Length > 2)
                {
                    strOutputFilePath = args[2];
                    if (strActionMode == "getinfo")
                    {
                        Console.WriteLine("工作模式是:" + strActionMode);
                        Console.WriteLine("PPT路径是:" + strPPTFileName);
                        Console.WriteLine("结果输出路径是:" + strOutputFilePath);
                        Console.WriteLine("开始执行");

                        ESBasic.Office.OfficeScanner scaner = new ESBasic.Office.OfficeScanner();
                        int nPage = scaner.GetInfo(strPPTFileName);
                        FileStream fs = File.Create(strOutputFilePath + "pagecount.txt");
                        StreamWriter sw = new StreamWriter(fs);
                        sw.WriteLine(nPage.ToString());
                        sw.Close();
                        fs.Close();

                        Console.WriteLine("执行结束");
                        return;
                    }
                    else if (strActionMode == "convert")
                    {
                        if (args.Length > 3)
                        {
                            strPrefix = args[3];
                            if (args.Length > 4)
                            {
                                string strBeginPage = args[4];
                                nBeginPage = int.Parse(strBeginPage);
                                if (args.Length > 5)
                                {
                                    string strEndPage = args[5];
                                    nEndPage = int.Parse(strEndPage);
                                    Console.WriteLine("工作模式是:" + strActionMode);
                                    Console.WriteLine("PPT路径是:" + strPPTFileName);
                                    Console.WriteLine("图片输出路径是:" + strOutputFilePath);
                                    Console.WriteLine("图片输出前缀是:" + strPrefix);
                                    Console.WriteLine("输出起始页码是:" + nBeginPage.ToString());
                                    Console.WriteLine("输出结束页码是:" + nEndPage.ToString());
                                    Console.WriteLine("开始执行");

                                    ESBasic.Office.OfficeScanner scaner = new ESBasic.Office.OfficeScanner();
                                    int nPage = scaner.Convert(strPPTFileName, strOutputFilePath, strPrefix, nBeginPage, nEndPage);
                                    FileStream fs = File.Create(strOutputFilePath + "pagecount.txt");
                                    StreamWriter sw = new StreamWriter(fs);
                                    sw.WriteLine(nPage.ToString());
                                    sw.Close();
                                    fs.Close();

                                    Console.WriteLine("执行结束");
                                    return;
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
