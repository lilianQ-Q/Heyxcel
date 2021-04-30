using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelLib.src.core.exceptions
{
    class ConvertColumnException : HeyxcelBaseException
    {
        #region Constructor
        public ConvertColumnException(string convertColumnExceptionMessage) : base(convertColumnExceptionMessage)
        {
        }
        #endregion
    }
}
