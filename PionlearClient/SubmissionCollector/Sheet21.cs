using System;

// ReSharper disable ArrangeThisQualifier
// ReSharper disable InconsistentNaming

namespace SubmissionCollector
{
    public partial class Sheet21
    {
        

        public void Sheet21_Startup(object sender, EventArgs e)
        {
            
           
            

            
            

            
        }

        private void Sheet21_Shutdown(object sender, EventArgs e)
        {
        }

        
        
        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += Sheet21_Startup;
            this.Shutdown += Sheet21_Shutdown;
        }

        #endregion

    }


    public class Profile
    {
        internal string SelectorRangeName { get; set; }
        internal string SelectionsRangeName { get; set; }
    }
}
