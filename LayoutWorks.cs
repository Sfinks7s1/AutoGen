namespace Auto
{
    //using ExcelExample;
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.EditorInput;
    using Autodesk.AutoCAD.Runtime;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text.RegularExpressions;
    using System.Threading;
    using TDMS.Interop;

    /// <summary>
    /// Класс содержит методы для работы с Layput в AutoCAD
    /// </summary>
    public sealed class LayoutWorks : IExtensionApplication
    {
        private const string ModuleName = "C_PSD_TO_LAYOUTS";
        private const string FunctionName = "LoadLayout";

        public void Initialize()
        {
        }

        public void Terminate()
        {
        }

        ///<summary>
        /// Метод производит внедрение всех внешних ссылок в открытом чертеже.
        ///</summary>
        [CommandMethod("Binding")]
        public void Binding()
        {
            Autodesk.AutoCAD.ApplicationServices.Core.Application.SetSystemVariable("CTAB", "Model");
            var editor = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                var cvPath = new CurrentVersion();
                var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;

                var myKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(cvPath.pathLanguage());
                var path = (String)(myKey.GetValue("LocalRootFolder"));
                editor.WriteMessage("\n " + path);

                var local = path.ToLower();
                string parseGuid = null;
                var reg = new Regex("rus", RegexOptions.IgnoreCase);
                var mc = reg.Matches(local);
                foreach (Match mat in mc)
                {
                    parseGuid = mat.Value;
                }
                if (parseGuid == "rus")
                {
                    doc.Database.CloseInput(true);
                    //Для русской версии AutoCAD
                    doc.SendStringToExecute("XREFBIND" + "\n" + "ВСЕ" + "\n", true, false, false);
                    doc.SendStringToExecute("SaveActiveDrawing" + "\n", true, false, false);

                    editor.WriteMessage("\n " + parseGuid);
                }
                else
                {
                    //Для английской версии AutoCAD

                    doc.Database.CloseInput(true);

                    doc.SendStringToExecute("XREFBIND" + "\n" + "ALL" + "\n" + "\n", true, false, false);
                    doc.SendStringToExecute("SaveActiveDrawing" + "\n", true, false, false);

                    editor.WriteMessage("\n " + parseGuid);
                }
                editor.WriteMessage("\n Внешние ссылки обновлены.");
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Метод для дублирования основного чертежа с удалением листов.
        /// При выполнении все чертежи-дубликаты по очереди открываются в AutoCAD.
        /// </summary>
        /// <param name="docPathname"></param>
        /// <param name="layoutName"></param>
        public void CopyLayoutWithOpen(string docPathname, string layoutName)
        {
            var expDocPathname = Path.Combine(@"D:\_TDMSLAYOUT\", layoutName + ".dwg");

            File.Copy(docPathname, expDocPathname, true);

            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.Open(expDocPathname, false);
            Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument = doc;
            var db = doc.Database;
            var editor = doc.Editor;
            editor.WriteMessage("\n tmpPathname " + docPathname);
            editor.WriteMessage("\n expDocPathname " + expDocPathname);

            try
            {
                using (doc.LockDocument())
                {
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        LayoutManager.Current.CurrentLayout = layoutName;
                        var lytDct = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                        foreach (var dictEntry in lytDct)
                        {
                            if (dictEntry.Key != "Model" && dictEntry.Key != layoutName)
                            {
                                var lyt = (Layout)tr.GetObject(dictEntry.Value, OpenMode.ForWrite);
                                lyt.Erase();
                            }
                        }
                        tr.Commit();
                        editor.Regen();
                    }
                    db.SaveAs(expDocPathname, true, DwgVersion.Current, db.SecurityParameters);
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught: " + ex);
            }
            doc.CloseAndDiscard();
        }

        /// <summary>
        /// Метод для дублирования основного чертежа с удалением листов.
        /// При выполнении чертежи-дубликаты не открываются в AutoCAD.
        /// Метод взаимодействует с TDMS!
        /// </summary>
        /// <param name="docPathname"></param>
        /// <param name="layoutName"></param>
        public void CopyLayoutWithoutOpenFromTdms(string docPathname, string layoutName)
        {
            try
            {
                var tmpPathname = @"D:\_TDMSLAYOUT\" + layoutName + ".dwg";

                File.Copy(docPathname, tmpPathname, true);

                var dbX = new Database(false, true);
                dbX.ReadDwgFile(tmpPathname, FileShare.ReadWrite, true, "");

                HostApplicationServices.WorkingDatabase = dbX;

                using (var trX = dbX.TransactionManager.StartTransaction())
                {
                    LayoutManager.Current.CurrentLayout = layoutName;
                    var lytDct = (DBDictionary)trX.GetObject(dbX.LayoutDictionaryId, OpenMode.ForWrite);
                    foreach (var dictEntry in lytDct)
                    {
                        var lyt = (Layout)trX.GetObject(dictEntry.Value, OpenMode.ForWrite);

                        if (dictEntry.Key != "Model" && dictEntry.Key != layoutName)
                        {
                            lyt.Erase();
                        }
                    }
                    dbX.CloseInput(true);
                    dbX.SaveAs(tmpPathname, DwgVersion.AC1027);
                    trX.Commit();
                }
                var attr = new Attribute();
                var tdmsApp = new TDMSApplication();
                var check = new Condition();
                attr.FindAttribute(dbX);
                tdmsApp.ExecuteScript(ModuleName, FunctionName, check.ParseGuid(docPathname), attr.GetOboznach(), layoutName, attr.GetPageNub(), tmpPathname);
            }
            catch (Autodesk.AutoCAD.Runtime.Exception)
            {
            }
        }

        /// <summary>
        /// Метод для дублирования основного чертежа с удалением листов.
        /// При выполнении чертежи-дубликаты не открываются в AutoCAD.
        /// </summary>
        /// <param name="docPathname"></param>
        /// <param name="layoutName"></param>
        public void CopyLayoutWithoutOpen(string docPathname, string layoutName)
        {
            try
            {
                var tmpPathname = @"D:\_TDMSLAYOUT\" + layoutName + ".dwg";

                File.Copy(docPathname, tmpPathname, true);
                var dbX = new Database(false, true);

                dbX.ReadDwgFile(tmpPathname, FileShare.ReadWrite, true, "");

                HostApplicationServices.WorkingDatabase = dbX;

                using (var trX = dbX.TransactionManager.StartTransaction())
                {
                    LayoutManager.Current.CurrentLayout = layoutName;
                    var lytDct = (DBDictionary)trX.GetObject(dbX.LayoutDictionaryId, OpenMode.ForWrite);
                    foreach (var dictEntry in lytDct)
                    {
                        var lyt = (Layout)trX.GetObject(dictEntry.Value, OpenMode.ForWrite);

                        if (dictEntry.Key != "Model" && dictEntry.Key != layoutName)
                        {
                            lyt.Erase();
                        }
                    }
                    dbX.CloseInput(true);
                    dbX.SaveAs(Path.GetDirectoryName(tmpPathname) + @"\" + Path.GetFileName(tmpPathname) + "_test" + Path.GetExtension(tmpPathname), DwgVersion.AC1027);
                    trX.Commit();
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception)
            {
            }
        }

        ///<summary>
        /// Метод для экспорта листов. Метод запускается по команде из AutoCAD.
        ///</summary>
        [CommandMethod("ExportLayoutWithoutTDMS")]
        public void ExportLayoutWithoutTdms()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            try
            {
                Directory.CreateDirectory(@"D:\_TDMSLayout\");
                foreach (var lytName in GetLayoutNames(db))
                {
                    CopyLayoutWithoutOpen(doc.Name, lytName);
                }
                HostApplicationServices.WorkingDatabase = db;
                Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument = doc;

                doc.SendStringToExecute("CloseAndDiscard" + "\n", true, false, false);
            }
            catch (System.Exception)
            {
                // ignored
            }
        }

        ///<summary>
        /// Метод для экспорта листов. Метод запускается по команде из AutoCAD. Каждый чертёж открывается в AutoCAD.
        ///</summary>
        [CommandMethod("DivideLayout", CommandFlags.Session)]
        public void DivideLayout()
        {
            if (System.Windows.Forms.MessageBox.Show(@"Выполнение команды может занять продолжительное время. Выполнить?", @"Экспорт листов", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question) != System.Windows.Forms.DialogResult.Yes)
            {
                return;
            }

            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var editor = doc.Editor;

            try
            {
                Directory.CreateDirectory(@"D:\_TDMSLayout\");

                foreach (var lytName in GetLayoutNames(doc.Database))
                {
                    var expDocPathname = @"D:\_TDMSLAYOUT\" + lytName + "_temp" + ".dwg";
                    File.Copy(doc.Name, expDocPathname, true);
                    this.ExportLayFindAttr(expDocPathname, lytName, doc.Name);
                }

                //var docx = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;

                //HostApplicationServices.WorkingDatabase = docx.Database;
                //Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument = docx;
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught: " + ex);
            }
        }

        /// <summary>
        /// Метод для дублирования основного чертежа с удалением листов
        /// </summary>
        /// <param name="expDocPathname"></param>
        /// <param name="layoutName"></param>
        /// <param name="originalPathName"></param>
        /// <param name="strDwgName"></param>
        public void ExportLayFindAttr(string expDocPathname, string layoutName, string strDwgName)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.Open(expDocPathname);
            var editor = doc.Editor;

            Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument = doc;
            Autodesk.AutoCAD.ApplicationServices.Core.Application.SetSystemVariable("CTAB", "Model");
            
            try
            {
                using (doc.LockDocument())
                {
                    using (var tr = doc.Database.TransactionManager.StartTransaction())
                    {
                        LayoutManager.Current.CurrentLayout = layoutName;
                        var lytDct = (DBDictionary)tr.GetObject(doc.Database.LayoutDictionaryId, OpenMode.ForRead);
                        foreach (var dictEntry in lytDct)
                        {
                            if (dictEntry.Key == "Model" || dictEntry.Key == layoutName)
                            {
                                continue;
                            }
                            tr.GetObject(dictEntry.Value, OpenMode.ForWrite).Erase();
                        }
                        tr.Commit();
                    }

                    //doc.Database.SaveAs(originalPathName, true, DwgVersion.Current, doc.Database.SecurityParameters);

                    //var attr = new Attribute();
                    //var tdmsApp = new TDMSApplication();
                    //var check = new Condition();

                    //attr.FindAttribute(db);

                    //if (tdmsApp.ExecuteScript(ModuleName, FunctionName, check.ParseGuid(strDwgName), attr.GetOboznach(), layoutName, attr.GetPageNub(), originalPathName))
                    //{
                    //    Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog("True return!!!");
                    //}
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught: " + ex);
            }
        }

        /// <summary>
        /// Метод возвращает список листов в чертеже
        /// </summary>
        /// <param name="db"></param>
        /// <returns></returns>
        public List<string> GetLayoutNames(Database db)
        {
            var layoutNames = new List<string>();
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var lytDct = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                    foreach (var lytDe in lytDct)
                    {
                        if (lytDe.Key == "Model") continue;
                        var lyt = (Layout)tr.GetObject(lytDe.Value, OpenMode.ForRead);
                        var ltr = (BlockTableRecord)tr.GetObject(lyt.BlockTableRecordId, OpenMode.ForRead);
                        if (ltr.Cast<ObjectId>().Count() > 1)
                        {
                            layoutNames.Add(lyt.LayoutName);
                        }
                    }
                    tr.Commit();
                }
            }
            catch (System.Exception)
            {
                // ignored
            }
            return layoutNames;
        }
    }
}