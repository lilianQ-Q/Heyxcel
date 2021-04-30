using ExcelLib.src.core.entitie;
using noxLogger.src;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelLib.src.core.profiler
{
    class ActionProfiler
    {
        #region Fields

        private static ActionProfiler _instance { get; set; }
        private Logger logger { get; set; }
        private List<Actions> actionsList { get; set; }

        #endregion

        #region Constructor

        /// <summary>
        /// Return the instance of the profiler.
        /// </summary>
        private ActionProfiler()
        {
            this.logger = new Logger();
            this.actionsList = new List<Actions>();
        }

        #endregion

        #region Singleton

        /// <summary>
        /// Returns profiler's instance.
        /// </summary>
        public static ActionProfiler GetInstance
        {
            get
            {
                if(_instance == null)
                {
                    new Logger().Debug("Creating action profiler..");
                    _instance = new ActionProfiler();
                }
                return (_instance);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Add a new action into the profiler's action list.
        /// </summary>
        /// <param name="action"></param>
        public void AddAction(Actions action)
        {
            this.actionsList.Add(action);
        }

        /// <summary>
        /// Returns profiler's stored actions.
        /// </summary>
        /// <returns></returns>
        public List<Actions> GetActions()
        {
            return (this.actionsList);
        }

        public void FlushAction()
        {
            this.actionsList.Clear();
        }

        public void LogActions()
        {
            this.logger.Debug("=== Action list ===");
            foreach(Actions action in this.actionsList)
            {
                this.logger.Debug(action.ToString());
            }
        }

        #endregion
    }
}
