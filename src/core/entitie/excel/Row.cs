using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelLib.src.core.entitie.excel
{
    class Row
    {
        #region Fields

        private int rowIndex { get; set; }
        private List<Cell> cellList { get; set; }

        #endregion

        #region Constructor

        /// <summary>
        /// Instanciate a new row.
        /// </summary>
        /// <param name="rowIndex"></param>
        public Row(int rowIndex)
        {
            this.rowIndex = rowIndex;
            this.cellList = new List<Cell>();
        }

        #endregion

        #region Methods

        /// <summary>
        /// Binds cells into current row.
        /// </summary>
        /// <param name="cells"></param>
        public void Bind(List<Cell> cells)
        {
            this.cellList = cells;
        }

        #endregion
    }
}
