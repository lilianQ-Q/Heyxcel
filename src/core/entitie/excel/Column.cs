using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelLib.src.core.entitie.excel
{
    class Column
    {
        #region Fields

        public string columnName { get; set; }
        public List<Cell> cellList { get; set; }

        #endregion

        #region Constructeur

        public Column(string columnName)
        {
            this.columnName = columnName;
            this.cellList = new List<Cell>();
        }

        #endregion

        #region Methods

        public void Bind(List<Cell> cells)
        {
            this.cellList = cells;
        }

        #endregion
    }
}
