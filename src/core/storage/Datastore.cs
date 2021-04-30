using ExcelLib.src.core.entitie.excel;
using ExcelLib.src.core.enumeration;
using ExcelLib.src.core.exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelLib.src.core.storage
{
    class Datastore
    {
        #region Fields

        private static Datastore _instance { get; set; }
        private DatastoreMode storeMode { get; set; }

        private List<Column> columnList { get; set; }
        private List<Row> rowList { get; set; }
        #endregion

        #region Constructor

        private Datastore(DatastoreMode mode)
        {
            this.storeMode = mode;
        }

        #endregion

        #region Singleton

        public static Datastore GetInstance
        {
            get
            {
                if(_instance == null)
                {
                    _instance = new Datastore(0);
                }
                return (_instance);
            }
        }

        #endregion

        #region Methods

        public void SwitchMode(int newMode)
        {
            this.storeMode = (DatastoreMode)newMode;
        }

        public void Add(string columnName, List<Cell> cellList)
        {
            if (this.storeMode.Equals(DatastoreMode.column))
            {
                if(this.columnList.Where(p => p.columnName.Equals(columnName)).Count() > 0)
                {
                    int columnIndex = this.columnList.FindIndex(p => p.columnName.Equals(columnName));
                    Column column = new Column(columnName);
                    column.Bind(cellList);
                    this.columnList[columnIndex].cellList = this.columnList[columnIndex].cellList.Union(cellList).ToList();
                }
                else
                {
                    Column column = new Column(columnName);
                    column.Bind(cellList);
                    this.columnList.Add(column);
                }
            }
            else
            {
                throw new DatastoreModeException("The datastore isn't in the proper mode, selected mode : row (1). You should turn it on column (0) mode.");
            }
        }

        public void Add(int rowIndex, List<Cell> cellList)
        {
            if (this.storeMode.Equals(DatastoreMode.row))
            {/*
                if(this.rowList.Where(p => p.))
                Row row = new Row(rowIndex);*/
            }
            else
            {
                throw new DatastoreModeException("The datastore isn't in the proper mode, selected mode : column (0). You should turn it on row (1) mode.");
            }
        }

        #endregion
    }
}
