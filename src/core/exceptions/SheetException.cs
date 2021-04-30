using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelLib.src.core.exceptions
{
    class SheetException : HeyxcelBaseException
    {
        #region Constructor

        public SheetException(string exceptionMessage) : base(exceptionMessage)
        {
        }

        #endregion
    }
}
