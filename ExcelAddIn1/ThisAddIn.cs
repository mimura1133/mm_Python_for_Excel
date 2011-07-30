using Microsoft.Scripting.Hosting;
using IronPython.Hosting;

namespace ExcelAddIn1
{
    public partial class ThisAddIn
    {
        public static string Code = "";

        static ScriptEngine _python;
        static ScriptScope _python_scope;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _python = Python.CreateEngine();
            _python_scope = _python.CreateScope();
            _python_scope.SetVariable("Application", Application);
            _python_scope.SetVariable("Function", Application.WorksheetFunction);
            _python_scope.SetVariable("Cells", Application.Cells);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public static void Run()
        {
            var cc =_python.CreateScriptSourceFromString(Code);
            cc.Execute(_python_scope);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
