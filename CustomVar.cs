namespace Auto
{
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.Runtime;

    /// <summary>
    /// Класс предназначен для установки внутренних системных переменных AutoCAD и дополнительных настроек в оптималный режим работы
    /// </summary>
    public sealed class CustomVar
    {
        [CommandMethod("CustomVar")]
        public void Variables()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;

            try
            {
                object zoomfactor = 100;
                object xrefnotify = 0;
                object whipthread = 3;
                object whiparc = 0;
                object vtfps = 1;
                object savetime = 10;
                object openpartial = 0;
                object maxactvp = 16;
                object lockui = 1;
                object highlight = 1;
                object hideprecision = 0;
                object gripobjlimit = 2;
                object dragp1 = 10000;
                object dragp2 = 10;
                object cmdinputhistorymax = 10;

                Application.SetSystemVariable("zoomfactor", zoomfactor);
                Application.SetSystemVariable("xrefnotify", xrefnotify);
                Application.SetSystemVariable("whipthread", whipthread);
                Application.SetSystemVariable("whiparc", whiparc);
                Application.SetSystemVariable("vtfps", vtfps);
                Application.SetSystemVariable("savetime", 10);
                Application.SetSystemVariable("openpartial", openpartial);
                Application.SetSystemVariable("maxactvp", maxactvp);
                Application.SetSystemVariable("lockui", lockui);
                Application.SetSystemVariable("highlight", highlight);
                Application.SetSystemVariable("hideprecision", hideprecision);
                Application.SetSystemVariable("gripobjlimit", gripobjlimit);
                Application.SetSystemVariable("dragp1", dragp1);
                Application.SetSystemVariable("dragp2", dragp2);
                Application.SetSystemVariable("cmdinputhistorymax", cmdinputhistorymax);
            }
            catch (Exception ex)
            {
                editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace);
            }
        }
    }
}