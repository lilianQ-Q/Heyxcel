using noxLogger.src;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelLib.src.core.exceptions
{
    class HeyxcelBaseException : Exception
    {
        #region Fields

        private static Logger logger = new Logger();

        #endregion

        #region Constructor
        public HeyxcelBaseException(string baseExceptionMessage) : base(baseExceptionMessage)
        {
            logger.Fatal($"Une exception a été levée : '{baseExceptionMessage}'.");
        }
        #endregion
    }
}
