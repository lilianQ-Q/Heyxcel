using ExcelLib.src.core.converter;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelLib.src.core.entitie.excel
{
    class Cell
    {
        #region Fields

        private Column cellColumn { get; set; }
        private Row cellRow { get; set; }
        private int intValue { get; set; }
        private string stringValue { get; set; }

        #endregion

        #region Constructor
        
        public Cell(string cellColumn, int rowIndex)
        {
            this.cellColumn = new Column(cellColumn);
            this.cellRow = new Row(rowIndex);
        }

        public Cell(int cellColumn, int rowIndex)
        {
            this.cellColumn = new Column(ColumnConverter.Get(cellColumn));
            this.cellRow = new Row(rowIndex);
        }

        #endregion

        #region Methods

        public void Bind(string value)
        {
            this.stringValue = value;
        }

        public void Bind(int value)
        {
            this.intValue = value;
        }

        public string GetString()
        {
            return (this.stringValue);
        }

        public int GetInt()
        {
            return (this.intValue);
        }

        #endregion
    }
}
