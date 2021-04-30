using ExcelLib.src.core.converter;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelLib.src
{
    public partial class Heyxcel
    {
        #region Fields
        private Dictionary<string, Dictionary<int, string>> storedValues { get; set; }
        #endregion

        #region Methods

        /// <summary>
        /// Clear the storedValues.
        /// </summary>
        public void ClearSavedValues()
        {
            this.storedValues.Clear();
        }

        public void Read(string column, int start, int end)
        {
            string cellValue = "";
            if (this.storedValues.ContainsKey(column))
            {
                this.storedValues.Remove(column);
            }
            if(this.worksheet != null)
            {
                Dictionary<int, string> tempDictionary = new Dictionary<int, string>();
                for(int i = start; i <= end; i++)
                {
                    Range range = this.worksheet.Cells[i, column] as Range;
                    cellValue = range.Value != null ? range.Value.ToString() : "null";
                    tempDictionary.Add(i, cellValue);
                    Console.Title = $"Column {column} - {i} / {end}";
                }
                this.storedValues.Add(column, tempDictionary);
            }
            else
            {
                this.logger.Error($"The current worksheet seems to be null. Unable to read \"{this.excelPath}\".");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="column"></param>
        /// <param name="start"></param>
        /// <param name="end"></param>
        public void Read(int column, int start, int end)
        {
            string columnName = ColumnConverter.Get(column);
            if (this.storedValues.ContainsKey(columnName))
            {
                this.storedValues.Remove(columnName);
            }
            if (this.worksheet != null)
            {
                Dictionary<int, string> tempDictionary = new Dictionary<int, string>();
                for (int i = start; i <= end; i++)
                {
                    Range range = this.worksheet.Cells[i, column] as Range;
                    string cellValue = range.Value != null ? range.Value.ToString() : "null";
                    tempDictionary.Add(i, cellValue);
                }
                this.storedValues.Add(columnName, tempDictionary);
            }
            else
            {
                this.logger.Error($"The current worksheet seems to be null. Unable to read \"{this.excelPath}\".");
            }
        }

        /// <summary>
        /// Get all row of a column in the stored values.
        /// 
        /// Author : Lilian DAMIENS
        /// Date : January 2021
        /// </summary>
        /// <param name="column"></param>
        /// <returns></returns>
        public Dictionary<int, string> GetColumnData(string column)
        {
            Dictionary<int, string> tempDictionary = new Dictionary<int, string>();
            if (this.storedValues.ContainsKey(column))
            {
                tempDictionary = this.storedValues[column];
            }
            else
            {
                this.logger.Error($"There is no stored values corresponding to {column} column.");
            }
            return (tempDictionary);
        }

        public Dictionary<string, string> ReadRowFromStoredValues(int rowIndex)
        {
            Dictionary<string, string> tmp = new Dictionary<string, string>();
            foreach(KeyValuePair<string, Dictionary<int, string>> pair in this.storedValues)
            {
                if (pair.Value.ContainsKey(rowIndex))
                {
                    tmp.Add(pair.Key, pair.Value[rowIndex]);
                }
                else
                {
                    tmp.Add(pair.Key, "");
                }
            }
            return (tmp);
        }

        #endregion
    }
}
