using noxLogger.src;
using Microsoft.Office.Interop.Excel;
using System;
using ExcelLib.src.core.exceptions;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using ExcelLib.src.core.converter;
using ExcelLib.src.core.enumeration;
using System.Diagnostics;
using System.Linq;

namespace ExcelLib.src
{
    public partial class Heyxcel
    {
        #region Fields

        private string excelPath { get; set; }
        private Logger logger { get; set; }
        private Application application { get; set; }
        private Workbooks workbooks { get; set; }
        private Workbook workbook { get; set; }
        private Worksheet worksheet { get; set; }
        private Range range { get; set; }
        private bool saveAction { get; set; } = true;
        private FileState fileState { get; set; }
        #endregion

        #region Constructor

        public Heyxcel(string excelPath)
        {
            this.excelPath = excelPath;
            this.logger = new Logger();
            this.fileState = FileState.close;
            this.storedValues = new Dictionary<string, Dictionary<int, string>>();
        }

        #endregion

        #region Methods

        /// <summary>
        /// Open the current excel file path. And load it into the current object. 
        /// 
        /// Author : Lilian DAMIENS
        /// Date : January 2021
        /// </summary>
        public void Open()
        {
            this.KillExcelFileProcess();
            this.logger.Debug($"Trying to open \"{this.excelPath}\".");
            if (this.CheckPath())
            {
                try
                {
                    this.application = new Application();
                    this.workbooks = this.application.Workbooks;
                    this.workbook = this.application.Workbooks.Open(this.excelPath, 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", true, false, 0, true, 1, 0);
                    this.worksheet = (Worksheet)this.workbook.Sheets[1];
                    this.range = this.worksheet.UsedRange;
                    this.logger.Success($"Excel file \"{this.excelPath}\" opened !");
                    this.fileState = FileState.opened;
                }
                catch (Exception exception)
                {
                    this.QuickClose();
                    throw new HeyxcelBaseException($"Unable to open the file. \"{exception.Message}\"");
                }
            }
            else
            {
                throw new HeyxcelBaseException($"File doesn't exists. \"{this.excelPath}\"");
            }
        }

        /// <summary>
        /// Close the current opened file.
        /// 
        /// Author : Lilian DAMIENS
        /// Date : January 2021
        /// </summary>
        public void Close()
        {
            if (this.fileState.Equals(FileState.opened))
            {
                this.logger.Debug($"Trying to close \"{this.excelPath}\".");
                try
                {
                    this.SecureClose();
                    this.logger.Success($"Excel file \"{this.excelPath}\" closed !");
                    this.fileState = FileState.close;
                }
                catch (Exception exception)
                {
                    throw new HeyxcelBaseException($"Fail to close the current excel file \"{this.excelPath}\" : {exception}");
                }
            }
        }

        /// <summary>
        /// Change the current sheet to the one passed throught parameters.
        /// 
        /// Author : Lilian DAMIENS
        /// Date : January 2021
        /// </summary>
        /// <param name="sheetIndex"></param>
        public void ChangeSheet(int sheetIndex)
        {
            this.logger.Debug($"Trying to change the current worksheet to '{sheetIndex}'...");
            try
            {
                this.worksheet = (Worksheet)this.workbook.Sheets[sheetIndex];
                this.logger.Success($"Current worksheet setted to '{sheetIndex}' !");
            }
            catch(Exception exception)
            {
                this.QuickClose();
                throw new SheetException($"Fail while changing the current worksheet to {sheetIndex}, {exception}.");
            }
        }

        private bool CheckPath()
        {
            bool exists = false;
            if (File.Exists(this.excelPath))
            {
                exists = true;
            }
            return (exists);
        }

        /// <summary>
        /// Close the current file without saving changes.
        /// 
        /// Author : Lilian DAMIENS
        /// Date : January 2021
        /// </summary>
        private void QuickClose()
        {
            this.fileState = FileState.close;
            if(this.workbook != null)
            {
                this.workbook.Save();
                this.workbook.Close(true, null, null);
                Marshal.FinalReleaseComObject(this.workbook);
            }
            if(this.workbooks != null)
            {
                this.workbooks.Close();
                Marshal.FinalReleaseComObject(this.workbooks);
            }
            if(this.application != null)
            {
                this.application.Quit();
                Marshal.FinalReleaseComObject(this.application);
            }
            this.KillExcelFileProcess();
        }

        /// <summary>
        /// Close the current file and saves changes.
        /// 
        /// Date : Lilian DAMIENS
        /// Author : January 2021
        /// </summary>
        private void SecureClose()
        {
            this.fileState = FileState.close;
            this.workbook.Save();
            this.workbook.Close(true, null, null);
            this.workbooks.Close();
            this.application.Quit();

            Marshal.ReleaseComObject(this.workbook);
            Marshal.ReleaseComObject(this.workbooks);
            Marshal.ReleaseComObject(this.application);
        }

        private void KillExcelFileProcess()
        {
            string[] splitedPath = this.excelPath.Split('/');
            string excelName = splitedPath[splitedPath.Length - 1];
            var processes = from p in Process.GetProcessesByName("EXCEL") select p;
            foreach (var process in processes)
            {
                if(process.MainWindowTitle == $"Microsoft Excel - {excelName}")
                {
                    process.Kill();
                }
            }

        }

        #endregion
    }
}
