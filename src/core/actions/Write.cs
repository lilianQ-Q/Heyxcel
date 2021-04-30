using ExcelLib.src.core.converter;
using ExcelLib.src.core.entitie;
using ExcelLib.src.core.enumeration;
using ExcelLib.src.core.exceptions;
using ExcelLib.src.core.profiler;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelLib.src
{
    public partial class Heyxcel
    {

        /// <summary>
        /// Writes a new string value into the targeted cell.
        /// 
        /// Author : Lilian DAMIENS
        /// Date : January 2021
        /// </summary>
        /// <param name="column"></param>
        /// <param name="row"></param>
        /// <param name="value"></param>
        public void Write(string column, int row, string value)
        {
            if (this.fileState.Equals(FileState.opened))
            {
                if (this.saveAction)
                {
                    ActionProfiler.GetInstance.AddAction(new Actions(0, column, row, value));
                }
                this.worksheet.Cells[row, column] = value;
                this.logger.Debug($"Wrote \"{value}\" into \"[{row};{column}]\" cell.");
            }
            else
            {
                throw new HeyxcelBaseException("File is currently close, your can't write into it.");
            }
        }

        /// <summary>
        /// Writes a new integer value into the targeted cell.
        /// 
        /// Author : Lilian DAMIENS
        /// Date : January 2021
        /// </summary>
        /// <param name="column"></param>
        /// <param name="row"></param>
        /// <param name="value"></param>
        public void Write(string column, int row, int value)
        {
            //Method to check if file is opened ?
            if (this.fileState.Equals(FileState.opened))
            {
                if (this.saveAction)
                {
                    ActionProfiler.GetInstance.AddAction(new Actions(0, column, row, value.ToString()));
                }
                this.worksheet.Cells[row, column] = value;
                this.logger.Debug($"Wrote \"{value}\" into \"[{row};{column}]\" cell.");
            }
            else
            {
                throw new HeyxcelBaseException("File is currently close, your can't write into it.");
            }
        }

    }
}
