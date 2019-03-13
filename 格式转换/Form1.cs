using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.IO;
using System.Diagnostics;


namespace 格式转换
{
    public partial class Form1 : Form
    {
        private BackgroundWorker bkWorker = new BackgroundWorker();
        private int percentValue = 0;  
        public Form1()
        {
            InitializeComponent();
            bkWorker.WorkerReportsProgress = true;
            bkWorker.WorkerSupportsCancellation = true;
            bkWorker.DoWork += new DoWorkEventHandler(DoWork);
            bkWorker.ProgressChanged += new ProgressChangedEventHandler(ProgessChanged);
            bkWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(CompleteWork);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                //打开选择框选择文件
                DialogResult result;
                // 打开文件选择窗口
                result = folderBrowserDialog1.ShowDialog();
                string filePath = "";
                if (result == DialogResult.OK)
                {
                    filePath = folderBrowserDialog1.SelectedPath;
                    textBox1.Text = filePath;
                    this.label2.Text = "转换进度:";
                    progressBar1.Value = 0;
                    
                }               

            }
            catch (Exception ex)
            {
                //LogHelper.WriteLog(typeof(Form1), ex.Message);
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            string filePath = textBox1.Text;
            if (String.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("请选择文件夹");
                return;
            }
            percentValue = 1;
            this.progressBar1.Maximum = 100;
            bkWorker.RunWorkerAsync();
           
        }
        //private void button3_Click(object sender, EventArgs e)
        //{
        //    string filePath = textBox1.Text;
        //    if (String.IsNullOrEmpty(filePath))
        //    {
        //        MessageBox.Show("请选择文件夹");
        //        return;
        //    }
        //    String[] files = Directory.GetFiles(filePath, "*.xls*", SearchOption.AllDirectories);
        //    if (files != null)
        //    {
        //        for (int i = 0; i < files.Length; i++)
        //        {
        //            int doc_index = files[i].LastIndexOf(".");
        //            string wpath = files[i];
        //            string ppath = files[i].Substring(0, doc_index) + ".pdf";
        //            Excel.XlFixedFormatType excelType = Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF;
        //            ExcelConvertPDF(wpath, ppath, excelType);

        //        }
        //    }
        //}
        private bool WordConvertPDF(string sourcePath, string targetPath, Word.WdExportFormat exportFormat)
        {
            bool result;
            object paramMissing = Type.Missing;
            Word.ApplicationClass wordApplication = new Word.ApplicationClass();
            Word.Document wordDocument = null;
            try
            {
                object paramSourceDocPath = sourcePath;
                string paramExportFilePath = targetPath;

                Word.WdExportFormat paramExportFormat = exportFormat;
                bool paramOpenAfterExport = false;
                Word.WdExportOptimizeFor paramExportOptimizeFor =
                        Word.WdExportOptimizeFor.wdExportOptimizeForPrint;
                Word.WdExportRange paramExportRange = Word.WdExportRange.wdExportAllDocument;
                int paramStartPage = 0;
                int paramEndPage = 0;
                Word.WdExportItem paramExportItem = Word.WdExportItem.wdExportDocumentContent;
                bool paramIncludeDocProps = true;
                bool paramKeepIRM = true;
                Word.WdExportCreateBookmarks paramCreateBookmarks =
                        Word.WdExportCreateBookmarks.wdExportCreateWordBookmarks;
                bool paramDocStructureTags = true;
                bool paramBitmapMissingFonts = true;
                bool paramUseISO19005_1 = false;

                wordDocument = wordApplication.Documents.Open(
                        ref paramSourceDocPath, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing);

                if (wordDocument != null)
                    wordDocument.ExportAsFixedFormat(paramExportFilePath,
                            paramExportFormat, paramOpenAfterExport,
                            paramExportOptimizeFor, paramExportRange, paramStartPage,
                            paramEndPage, paramExportItem, paramIncludeDocProps,
                            paramKeepIRM, paramCreateBookmarks, paramDocStructureTags,
                            paramBitmapMissingFonts, paramUseISO19005_1,
                            ref paramMissing);
                result = true;
                
            }
            finally
            {
                if (wordDocument != null)
                {
                    wordDocument.Close(ref paramMissing, ref paramMissing, ref paramMissing);
                    wordDocument = null;
                }
                if (wordApplication != null)
                {
                    wordApplication.Quit(ref paramMissing, ref paramMissing, ref paramMissing);
                    wordApplication = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }
        private bool ExcelConvertPDF(string sourcePath, string targetPath, Excel.XlFixedFormatType targetType)
        {
            bool result;
            object missing = Type.Missing;
            Excel.ApplicationClass application = null;
            Excel.Workbook workBook = null;
            try
            {
                application = new Excel.ApplicationClass();
                object target = targetPath;
                object type = targetType;
                workBook = application.Workbooks.Open(sourcePath, missing, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing, missing, missing, missing, missing);

                workBook.ExportAsFixedFormat(targetType, target, Excel.XlFixedFormatQuality.xlQualityStandard, true, true, missing, missing, missing, missing);
                result = true;
            }
            catch
            {
                result = false;
            }
            finally
            {
                if (workBook != null)
                {
                    workBook.Close(true, missing, missing);
                    workBook = null;
                }
                if (application != null)
                {
                    application.Quit();
                    application = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }
        public void DoWork(object sender, DoWorkEventArgs e)
        {
            e.Result = ProcessProgress(bkWorker, e);
        }

        public void ProgessChanged(object sender, ProgressChangedEventArgs e)
        {
            // bkWorker.ReportProgress 会调用到这里，此处可以进行自定义报告方式  
            this.progressBar1.Value = e.ProgressPercentage;
            int percent = (int)(e.ProgressPercentage / percentValue);
            this.label2.Text = "转换进度:" + Convert.ToString(percent) + "%";
        }
        public void CompleteWork(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("转换完毕!");
        }
        private int ProcessProgress(object sender, DoWorkEventArgs e)
        {
            string filePath = textBox1.Text;
            String[] files = Directory.GetFiles(filePath,"*.*", SearchOption.AllDirectories);
            int length = files.Length;
            if (files != null)
            {
                for (int i = 0; i < files.Length; i++)
                {
                    try
                    {
                        if (files[i].Contains(".doc"))
                        {
                            int doc_index = files[i].LastIndexOf(".");
                            string wpath = files[i];
                            if (wpath.Contains("$"))
                            {
                                continue;
                            }
                            string ppath = files[i].Substring(0, doc_index) + ".pdf";
                            Word.WdExportFormat wd = Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF;
                            WordConvertPDF(wpath, ppath, wd);
                            File.Delete(wpath);
                        }
                        else if (files[i].Contains(".xls"))
                        {
                            int doc_index = files[i].LastIndexOf(".");
                            string wpath = files[i];
                            string ppath = files[i].Substring(0, doc_index) + ".pdf";
                            Excel.XlFixedFormatType excelType = Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF;
                            ExcelConvertPDF(wpath, ppath, excelType);
                            File.Delete(wpath);
                            Process[] process = Process.GetProcessesByName("EXCEL.EXE");
                            foreach (Process p in process)
                            {
                                if (!p.HasExited)  // 如果程序没有关闭，结束程序
                                {
                                    p.Kill();
                                    p.WaitForExit();
                                }
                            }

                        }
                        if (bkWorker.CancellationPending)
                        {
                            e.Cancel = true;
                            return -1;
                        }
                        else
                        {
                            int pcount = (i + 1) * 100 / length;
                            // 状态报告  
                            bkWorker.ReportProgress(pcount);
                            // 等待，用于UI刷新界面，很重要  
                            System.Threading.Thread.Sleep(1);
                        }
                    }
                    catch (Exception ex) {
                        MessageBox.Show(files[i]);
                    }

                }
                
                //MessageBox.Show("完成");
            }
            
            return -1;
        }

    }
}
