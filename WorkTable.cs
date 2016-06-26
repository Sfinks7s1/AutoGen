using System;
using Microsoft.Win32;

namespace Auto
{
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.EditorInput;
    using Autodesk.AutoCAD.Runtime;
    using System.IO;
    using System.Text.RegularExpressions;
    using TDMS.Interop;
    using ShdTbl = Autodesk.Aec.PropertyData.DatabaseServices.ScheduleTable;

    /// <summary>
    /// Класс хранит в себе методы для реализации добавления внешних ссылок в AEC компонент Shedule Table в чертеже.
    /// Ссылки даются на чертёж из TDMS. Также реализовано обновление внешних ссылок в таблицах.
    /// </summary>
    public sealed class WorkTable
    {
        /// <summary>
        /// Метод реализует добавление ссылки на чертёж в TDMS в конкретно выбранную Shedule Table.
        /// </summary>
        [CommandMethod("AddExternalSheduleTable")]
        public void AddExternalSheduleTable()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor editor = doc.Editor;
            string docName = doc.Name;

            try
            {
                using (Transaction tr = doc.TransactionManager.StartTransaction())
                {
                    PromptSelectionResult acSSPrompt = doc.Editor.GetSelection();
                    if (acSSPrompt.Status == PromptStatus.OK)
                    {
                        SelectionSet acSSet = acSSPrompt.Value;

                        foreach (SelectedObject acSSObj in acSSet)
                        {
                            if (acSSObj != null)
                            {
                                ShdTbl acEnt = tr.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as ShdTbl;
                                if (acEnt != null)
                                {
                                    if (System.Convert.ToString(acEnt.GetType()) == "Autodesk.Aec.PropertyData.DatabaseServices.ScheduleTable")
                                    {
                                        TDMSApplication tdmsApp = new TDMSApplication();

                                        Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Minimized;
                                        tdmsApp.Visible = false;
                                        tdmsApp.Visible = true;

                                        string moduleName = "CMD_SYSLIB";
                                        string functionName = "CheckOutSelObj";

                                        string pathName = tdmsApp.ExecuteScript(moduleName, functionName);

                                        string extension = Path.GetExtension(pathName);

                                        if (extension.ToLower() == ".dwg")
                                        {
                                            if (acEnt.ScanExternalReferences)
                                            {
                                                TDMSObject tdmsObj;

                                                //Берём GUID из пути файла выгруженного на диск
                                                string guidFromFile = docName;
                                                string parseGuidFromFile = null;
                                                Regex regFF = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                                                MatchCollection mcFF = regFF.Matches(guidFromFile);
                                                foreach (Match mat in mcFF)
                                                {
                                                    parseGuidFromFile += mat.Value.ToString();
                                                }
                                                parseGuidFromFile = parseGuidFromFile.Remove(0, 38);

                                                //Берём GUID из выбранного в TDMS файла - внешней ссылки
                                                string guidFromTDMS = pathName;
                                                string parseGuidFromTDMS = null;
                                                Regex regFT = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                                                MatchCollection mcFT = regFT.Matches(guidFromTDMS);
                                                foreach (Match mat in mcFT)
                                                {
                                                    parseGuidFromTDMS += mat.Value.ToString();
                                                }
                                                parseGuidFromTDMS = parseGuidFromTDMS.Remove(0, 38);

                                                tdmsObj = tdmsApp.GetObjectByGUID(parseGuidFromTDMS);

                                                if (parseGuidFromFile != parseGuidFromTDMS)
                                                {
                                                    string guid = pathName;
                                                    string parseGuid = null;
                                                    Regex reg = new Regex(".dwg", RegexOptions.IgnoreCase);
                                                    MatchCollection mc = reg.Matches(guid);
                                                    foreach (Match mat in mc)
                                                    {
                                                        parseGuid += mat.Value.ToString();
                                                    }
                                                    if (parseGuid.ToLower() == ".dwg")
                                                    {
                                                        editor.WriteMessage("\n PathNameFromTDMS: " + parseGuidFromTDMS);
                                                        editor.WriteMessage("\n docName: " + parseGuidFromFile);

                                                        var mainFile = tdmsObj.Files.Main;
                                                        mainFile.CheckOut(pathName);

                                                        Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Maximized;

                                                        acEnt.SetScheduleDrawingName(pathName); //GetScheduleDrawingName());
                                                        acEnt.RegenerateTable(true);
                                                    }
                                                    else
                                                    {
                                                        acEnt.ScanExternalReferences = true;
                                                        acEnt.RegenerateTable(true);

                                                        editor.WriteMessage("\n PathNameFromTDMS: " + parseGuidFromTDMS);
                                                        editor.WriteMessage("\n docName: " + parseGuidFromFile);

                                                        //tdmsObj.CheckOut();
                                                        var mainFile = tdmsObj.Files.Main;
                                                        mainFile.CheckOut(pathName);

                                                        Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Maximized;

                                                        acEnt.SetScheduleDrawingName(pathName);
                                                        acEnt.RegenerateTable(forceUpdate: true);
                                                    }
                                                }
                                                else
                                                {
                                                    editor.WriteMessage("\nВыбранный файл не является файлом dwg.");
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            // ДЛЯ ПЕРЕБОРА ВСЕХ ТАБЛИЦ И ЗАМЕНЫ ЗНАЧЕНИ ПУТЕЙ В НИХ
                            //BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                            //BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                            //foreach (ObjectId ObjId in btr)
                            //{
                            //    ShdTbl tab = (ShdTbl)tr.GetObject(ObjId, OpenMode.ForWrite);

                            //    if (System.Convert.ToString(tab.GetType()) == "Autodesk.Aec.PropertyData.DatabaseServices.ScheduleTable")
                            //    {
                            //        Autodesk.AutoCAD.Windows.OpenFileDialog sfd = new Autodesk.AutoCAD.Windows.OpenFileDialog("Открыть файл на локальном диске", doc.Name, "dwg", "Open file", Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.DefaultIsFolder);
                            //        if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                            //        {
                            //            //editor.WriteMessage("\n Hyperlinks activate: {0}", tab.ScanExternalReferences.ToString());
                            //            tab.SetScheduleDrawingName(sfd.Filename); //GetScheduleDrawingName());
                            //            //editor.WriteMessage("\n Hyperlinks Relative Path: {0}", tab.GetScheduleDrawingName());
                            //            tab.RegenerateTable(forceUpdate: true);
                            //        }
                            //    }
                            //    else
                            //    {
                            //        return;
                            //    }
                            //}
                            tr.Commit();
                        }
                    }
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: \n{0}", ex); }
        }

        /// <summary>
        /// Метод для обновления ссылок в Shedule Table, которые не относятся к чертежам в TDMS
        /// </summary>
        [CommandMethod("FindSheduleTable")]
        public void FindSheduleTable()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor editor = doc.Editor;

            try
            {
                using (Transaction tr = doc.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    foreach (ObjectId ObjId in btr)
                    {
                        ShdTbl tab = (ShdTbl)tr.GetObject(ObjId, OpenMode.ForWrite);

                        if (System.Convert.ToString(tab.GetType()) == "Autodesk.Aec.PropertyData.DatabaseServices.ScheduleTable")
                        {
                        }
                        else
                        {
                            return;
                        }
                    }
                    tr.Commit();
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: \n{0}", ex); }
        }

        /// <summary>
        /// Метод реализует поиск в чертеже Shedule Table и в них заполненных полей с путями, которые ведут к внешним чертежам, данные из которых берутся для заполнения
        /// Shedule Table. После поиска происходит запрос в TDMS на выгрузку чертежей, GUID которых хранится в возвращённых путях.
        /// Метод актуален как для полных и относительных путей в ссылках.
        /// </summary>
        [CommandMethod("UpdateTable")]
        public void UpdateTable()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor editor = doc.Editor;

            try
            {
                Condition Check = new Condition();
                if (Check.CheckPath())
                {
                    if (Check.CheckTdmsProcess())
                    {
                        TDMSApplication tdmsApp = new TDMSApplication();
                        TDMSObject tdmsObj = null;

                        using (Transaction tr = doc.TransactionManager.StartTransaction())
                        {
                            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                            foreach (ObjectId ObjId in btr)
                            {
                                ShdTbl tab = tr.GetObject(ObjId, OpenMode.ForWrite) as ShdTbl;
                                if (tab != null)
                                {
                                    if (System.Convert.ToString(tab.GetType()) == "Autodesk.Aec.PropertyData.DatabaseServices.ScheduleTable")
                                    {
                                        if (tab.ScanExternalReferences)
                                        {
                                            string fullPath = tab.GetScheduleDrawingName();
                                            if (fullPath != "")
                                            {
                                                editor.WriteMessage("\n Full Path for external table: " + fullPath);

                                                string pathGuid = fullPath;

                                                Condition parse = new Condition();

                                                //парсинг пути, получение актуального GUID объекта
                                                pathGuid = parse.ParseGuid(pathGuid);
                                                editor.WriteMessage("\n GUID object - external table: " + pathGuid);

                                                //Получаем объект ТДМС
                                                tdmsObj = tdmsApp.GetObjectByGUID(pathGuid);

                                                //Выгрузка чертежа из базы ТДМС
                                                var mainFile = tdmsObj.Files.Main;

                                                string PathMainFile = Path.GetDirectoryName(doc.Name).ToString();
                                                PathMainFile = PathMainFile.Remove(PathMainFile.Length - 39);
                                                string fileName = Path.GetFileName(fullPath);
                                                string FullUserPath = PathMainFile + @"\" + pathGuid + @"\" + fileName;

                                                mainFile.CheckOut(FullUserPath);
                                                tab.SetScheduleDrawingName(FullUserPath);
                                                tab.RegenerateTable(forceUpdate: true);

                                                pathGuid = null;
                                            }
                                        }
                                    }
                                }
                            }
                            tr.Commit();
                        }
                        CurrentVersion cvPath = new CurrentVersion();
                        RegistryKey myKey = Registry.CurrentUser.OpenSubKey(cvPath.pathLanguage());
                        String path = (String)(myKey.GetValue("LocalRootFolder"));
                        editor.WriteMessage("\n " + path);

                        string local = path.ToLower();
                        string parseGuid = null;
                        Regex reg = new Regex("rus", RegexOptions.IgnoreCase);
                        MatchCollection mc = reg.Matches(local);
                        foreach (Match mat in mc)
                        {
                            parseGuid = mat.Value;
                        }
                        if (parseGuid == "rus")
                        {
                            //Для русской версии AutoCAD
                            doc.SendStringToExecute("_SCHEDULEUPDATENOW" + "\n" + "Все " + "\n", true, false, false);
                        }
                        else
                        {
                            //Для английской версии AutoCAD
                            doc.SendStringToExecute("_SCHEDULEUPDATENOW" + "\n" + "ALL " + "\n", true, false, false);
                        }
                    }
                    else
                    {
                        editor.WriteMessage("\nНевозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                    }
                }
                else
                {
                    editor.WriteMessage("\nДокумент не принадлежит TDMS!");
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\nException caught: {0}\n" + ex);
            }
        }
    }
}