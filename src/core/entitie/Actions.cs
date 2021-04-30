using ExcelLib.src.core.enumeration;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelLib.src.core.entitie
{
    class Actions
    {
        #region Fields

        private string concernedColumn { get; set; }
        private int concernedRow { get; set; }
        private string assignedValue { get; set; }
        private string readValue { get; set; }

        private ActionType actionType { get; set; }

        #endregion

        #region Constructor

        public Actions(int actionType, string column, int row, string value)
        {
            this.actionType = (ActionType)actionType;
            this.concernedColumn = column;
            this.concernedRow = row;
            this.assignedValue = value;
        }

        #endregion

        #region Methods

        public override string ToString()
        {
            if (this.actionType.Equals(ActionType.read))
            {
                return ($"Cell [{this.concernedRow};{this.concernedColumn}] contained \"{this.assignedValue}\".");
            }
            else
            {
                return ($"Cell [{this.concernedRow};{this.concernedColumn}] filled with \"{this.assignedValue}\".");
            }
        }

        #endregion
    }
}
