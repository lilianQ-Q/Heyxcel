using ExcelLib.src.core.exceptions;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelLib.src.core.converter
{
    public class ColumnConverter
    {
        #region Fields

        private static ColumnConverter _instance { get; set; }
        private Dictionary<string, int> columnToIndex { get; set; }
        private Dictionary<int, string> indexToColumn { get; set; }
        #endregion

        #region Constructor

        /// <summary>
        /// Create a new instance of ColumnConverter.
        /// </summary>
        private ColumnConverter()
        {
            this.columnToIndex = new Dictionary<string, int>();
            this.indexToColumn = new Dictionary<int, string>();
        }

        #endregion

        #region Singleton 

        /// <summary>
        /// Return the unique instance of the ColumnConverter class.
        /// 
        /// Author : Lilian DAMIENS
        /// Date : January 2021
        /// </summary>
        /// <returns></returns>
        private static ColumnConverter GetInstance()
        {
            if(_instance == null)
            {
                _instance = new ColumnConverter();
            }
            return (_instance);
        }

        #endregion

        /// <summary>
        /// Return the column's index throught the letters of this one.
        /// 
        /// Author : Lilian
        /// Date : January 2021
        /// </summary>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public static int Get(string columnName)
        {
            if (String.IsNullOrEmpty(columnName))
            {
                throw new ConvertColumnException($"Column's name is null or empty.");
            }
            ColumnConverter instance = GetInstance();
            if (instance.columnToIndex.ContainsKey(columnName))
            {
                return (instance.columnToIndex[columnName]);
            }
            columnName = columnName.ToUpperInvariant();
            int index = 0;
            foreach(char letter in columnName)
            {
                index *= 26;
                index += (letter - 'A' + 1);
            }
            instance.columnToIndex.Add(columnName, index);
            return (index);
        }

        /// <summary>
        /// Return the column's letters throught the index of this one.
        /// 
        /// Author : Lilian
        /// Date : January 2021
        /// </summary>
        /// <param name="columnIndex"></param>
        public static string Get(int columnIndex)
        {
            ColumnConverter instance = GetInstance();
            if (instance.indexToColumn.ContainsKey(columnIndex))
            {
                return (instance.indexToColumn[columnIndex]);
            }
            int dividend = columnIndex;
            string columnName = String.Empty;
            int modulo;
            while(dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }
            instance.indexToColumn.Add(columnIndex, columnName);
            return (columnName);
        }
    }
}
