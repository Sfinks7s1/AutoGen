using System.Threading;
using System.Threading.Tasks;

namespace Auto
{
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.EditorInput;
    using Autodesk.AutoCAD.Runtime;
    using System;
    using System.IO;
    using TDMS.Interop;

    /// <summary>
    /// Класс включает в себя методы для сохранения \ импорта чертежей в базе
    /// </summary>
    public class SaveOptions
    {
        private readonly Document _doc;
        private readonly Editor _editor;

        public SaveOptions()
        {
            _doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            _editor = _doc.Editor;
        }

        #region Опции сохранения

        public void Initialize()
        {
        }

        public void Terminate()
        {
        }

        /// <summary>
        /// Сохранение чертежа на локальном диске
        /// </summary>
        /// <returns></returns>
        public static string SaveDialog(Document doc)
        {
            var sfd = new Autodesk.AutoCAD.Windows.SaveFileDialog("Сохранить на локальном диске", doc.Name, "dwg", "saving file", Autodesk.AutoCAD.Windows.SaveFileDialog.SaveFileDialogFlags.DefaultIsFolder);
            return (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK) ? sfd.Filename : String.Empty;
        }

        /// <summary>
        /// Открытие чертежа с локального диска
        /// </summary>
        /// <returns></returns>
        public string OpenDialog()
        {
            var sfd = new Autodesk.AutoCAD.Windows.OpenFileDialog("Открыть файл на локальном диске", "Drawing1", "dwg", "Open file", Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.DefaultIsFolder);
            return (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK) ? sfd.Filename : String.Empty;
        }

        /// <summary>
        /// Метод для локального сохранения чертежа по заданному пути
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="strDwgName"></param>
        /// <param name="pathName"></param>
        public static void LocalSave(Document doc, string strDwgName, string pathName)
        {
            try
            {
                var fullPath = pathName + "\\" + strDwgName;
                doc.Database.CloseInput(true);
                doc.Database.SaveAs(fullPath, true, DwgVersion.Current, doc.Database.SecurityParameters);
            }
            catch (System.Exception ex)
            {
                TraceSource.TraceMessage(TraceType.Error, ex.ToString());
                doc.Editor.WriteMessage("\n Exception caught: \n" + ex);
            }
        }

        /// <summary>
        /// Метод для локального сохранения чертежа по заданному пути (1 параметр, путь + имя чертежа)
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="fullPathAndDwgName"></param>
        public static void LocalSave(Document doc, string fullPathAndDwgName)
        {
            try
            {
                doc.Database.CloseInput(true);
                doc.Database.SaveAs(fullPathAndDwgName, true, DwgVersion.Current, doc.Database.SecurityParameters);
            }
            catch (System.Exception ex)
            {
                doc.Editor.WriteMessage("\n Exception caught: \n" + ex);
            }
        }

        #endregion Опции сохранения

        #region Методы, реализующие кнопки для сохранения чертежей

        /// <summary>
        /// Метод для сохранения чертежа локально, если чертёж открыт не из TDMS и сохранения чертежа в базе TDMS если чертёж открыт из базы (Флаг блокировки для объекта в TDMS не снимается).
        /// </summary>
        [CommandMethod("SaveActiveDrawing", CommandFlags.Session | CommandFlags.NoBlockEditor)]
        public void SaveActiveDrawing()
        {
            TraceSource.TraceMessage(TraceType.Information, "[---- Сохранить ----]");
           
            var check = new Condition();
            try
            {
                var strDwgName = _doc.Name;

                TraceSource.TraceMessage(TraceType.Information, "Имя документа: " + strDwgName);

                if (!string.IsNullOrEmpty(Path.GetDirectoryName(strDwgName)))
                {
                    TraceSource.TraceMessage(TraceType.Information, "Путь к чертежу не пустой, чертёж уже был сохранён ранее.");

                    if (check.CheckPath())
                    {
                        TraceSource.TraceMessage(TraceType.Information, "Чертёж принадлежит TDMS. Путь сохранения файла совпадает с путём " +  @"C:\TEMP\");

                        if (check.CheckTdmsProcess())
                        {
                            TraceSource.TraceMessage(TraceType.Information, "TDMS запущен.");
                            var tdmsApp = new TDMSApplication();

                            //Получение объекта по GUID
                            TDMSObject tdmsObj = tdmsApp.GetObjectByGUID(check.ParseGuid(strDwgName));

                            //Заблокирован ли чертёж текущим пользователем?
                            var lockOwner = tdmsObj.Permissions.LockOwner;
                            if (lockOwner)
                            {
                                TraceSource.TraceMessage(TraceType.Information, "Чертёж открыт для редактирования.");

                                //Сохранение изменений в текущем чертеже
                                _doc.Database.CloseInput(true);
                                _doc.Database.SaveAs(strDwgName, true, DwgVersion.Current, _doc.Database.SecurityParameters);

                                TraceSource.TraceMessage(TraceType.Information, "Чертёж сохранён успешно.");
                                
                                //Загрузка в базу ТДМС
                                tdmsObj.CheckIn();
                                tdmsObj.Update();

                                TraceSource.TraceMessage(TraceType.Information, "Чертёж загружен и обновлён в TDMS успешно.");
                            }
                            else
                            {
                                TraceSource.TraceMessage(TraceType.Information, "Чертёж открыт на просмотр.");

                                _editor.WriteMessage("\n Документ открыт на просмотр, изменения не будут сохранены в TDMS. Сохраните изменения локально.");
                            }
                        }
                        else
                        {
                            TraceSource.TraceMessage(TraceType.Information, "TDMS не запущен или количество запущенных приложений TDMS более одного.");

                            Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog("\n Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                        }
                    }
                    else
                    {
                        TraceSource.TraceMessage(TraceType.Information, "Чертёж не принадлежит TDMS. И должен быть сохранён локально.");

                        _doc.Database.SaveAs(strDwgName, true, DwgVersion.Current, _doc.Database.SecurityParameters);
                        _editor.WriteMessage("\n Документ сохранён локально, не принадлежит TDMS!");

                        TraceSource.TraceMessage(TraceType.Information, "Документ сохранён локально.");
                    }
                }
                else
                {
                    try
                    {
                        TraceSource.TraceMessage(TraceType.Information, "Путь к чертежу пустой, чертёж не был сохранён ранее.");

                        _doc.Database.SaveAs(SaveDialog(_doc), true, DwgVersion.Current, _doc.Database.SecurityParameters);

                        TraceSource.TraceMessage(TraceType.Information, "Чертёж сохранён.");
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex)
                    {
                        if (ex.ErrorStatus == ErrorStatus.FileAccessErr || ex.ErrorStatus == ErrorStatus.FileLockedByAutoCAD)
                        {
                            TraceSource.TraceMessage(TraceType.Warning, "Невозможно сохранить, доступ к файлу заблокирован приложением или уже используется. Возможно недостаточно прав на перезапись.");

                            Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog("\n Невозможно сохранить, доступ к файлу заблокирован приложением или уже используется. Возможно недостаточно прав на перезапись в данный каталог.");
                        }
                        else if (ex.ErrorStatus == ErrorStatus.InvalidInput)
                        {
                            TraceSource.TraceMessage(TraceType.Warning, "Команда отменена.");

                            _editor.WriteMessage("\n Команда отменена \n");
                        }
                        else
                        {
                            TraceSource.TraceMessage(TraceType.Error, ex.ToString());

                            _editor.WriteMessage(ex.ToString());
                        }
                    }
                }
                TraceSource.TraceMessage(TraceType.Information, "[-------------------]");
            }
            catch (System.Exception ex)
            {
                TraceSource.TraceMessage(TraceType.Error, ex.ToString());

                _editor.WriteMessage("\n Exception caught: \n" + ex);
            }
        }

        /// <summary>
        /// Метод для закрытия чертежа без сохранения изменений
        /// </summary>
        [CommandMethod("CloseAndDiscard", CommandFlags.Session | CommandFlags.NoBlockEditor)]
        public void CloseAndDiscard()
        {
            TraceSource.TraceMessage(TraceType.Information, "[---- Закрыть без сохранения ----]");
            if (System.Windows.Forms.MessageBox.Show(@"Действительно закрыть чертёж?", @"Закрыть чертёж", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question) != System.Windows.Forms.DialogResult.Yes)
            {
                return;
            }

            var check = new Condition();

            try
            {
                if (check.CheckPath())
                {
                    TraceSource.TraceMessage(TraceType.Information, "Чертёж принадлежит TDMS. Путь сохранения файла совпадает с путём " + @"C:\TEMP\");

                    if (check.CheckTdmsProcess())
                    {
                        TraceSource.TraceMessage(TraceType.Information, "TDMS запущен.");

                        var tdmsApp = new TDMSApplication();

                        var strDwgName = _doc.Name;

                        //Получение объекта по GUID
                        var tdmsObj = tdmsApp.GetObjectByGUID(check.ParseGuid(strDwgName));

                        //Заблокировани ли чертёж текущим пользователем
                        var lockOwner = tdmsObj.Permissions.LockOwner;
                        if (lockOwner)
                        {
                            TraceSource.TraceMessage(TraceType.Information, "Чертёж открыт для редактирования.");

                            tdmsObj.UnlockCheckIn(0);
                            TraceSource.TraceMessage(TraceType.Information, "Чертёж разблокирован в TDMS.");

                            _doc.CloseAndDiscard();
                            TraceSource.TraceMessage(TraceType.Information, "Чертёж закрыт без изменений.");

                            tdmsObj.Update();
                            TraceSource.TraceMessage(TraceType.Information, "Чертёж обновлён в TDMS успешно.");
                        }
                        else
                        {
                            _doc.CloseAndDiscard();
                            TraceSource.TraceMessage(TraceType.Information, "Чертёж закрыт без изменений.");
                        }
                    }
                    else
                    {
                        TraceSource.TraceMessage(TraceType.Information, "Чертёж не принадлежит TDMS. И должен быть сохранён локально.");

                        Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog("Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                    }
                }
                else
                {
                    _doc.CloseAndDiscard();// Закрыть без сохранения
                    TraceSource.TraceMessage(TraceType.Information, "Чертёж закрыт без изменений.");
                }
                TraceSource.TraceMessage(TraceType.Information, "[--------------------------------]");
            }
            catch (System.Exception ex)
            {
                TraceSource.TraceMessage(TraceType.Error, ex.ToString());
                _editor.WriteMessage("\n Exception caught: \n" + ex);
            }
        }

        /// <summary>
        /// Метод для сохранения чертежа локально, если чертёж открыт не из TDMS и сохранения чертежа в базе TDMS если чертёж открыт из базы (Флаг блокировки для объекта в TDMS снимается). После окончания процедуры сохранения чертёж в AutoCAD закрывается.
        /// </summary>
        [CommandMethod("SaveAndClose", CommandFlags.Session | CommandFlags.NoBlockEditor)]
        public void SaveAndClose()
        {
            TraceSource.TraceMessage(TraceType.Information, "[---- Сохранить и закрыть ----]");
            var check = new Condition();
            try
            {
                var strDwgName = _doc.Name;
                if (!string.IsNullOrEmpty(Path.GetDirectoryName(strDwgName)))
                {
                    TraceSource.TraceMessage(TraceType.Information, "Путь к чертежу не пустой, чертёж уже был сохранён ранее.");

                    if (check.CheckPath())
                    {
                        TraceSource.TraceMessage(TraceType.Information, "Чертёж принадлежит TDMS. Путь сохранения файла совпадает с путём " + @"C:\TEMP\");

                        if (check.CheckTdmsProcess())
                        {
                            TraceSource.TraceMessage(TraceType.Information, "TDMS запущен.");
                            var tdmsApp = new TDMSApplication();

                            //Получение объекта по GUID
                            var tdmsObj = tdmsApp.GetObjectByGUID(check.ParseGuid(strDwgName));

                            //Заблокировани ли чертёж текущим пользователем
                            var lockOwner = tdmsObj.Permissions.LockOwner;
                            if (lockOwner)
                            {
                                TraceSource.TraceMessage(TraceType.Information, "Чертёж открыт для редактирования.");
                                _doc.Database.CloseInput(true);
                                _doc.CloseAndSave(_doc.Name);
                                TraceSource.TraceMessage(TraceType.Information, "Чертёж сохранён закрыт успешно.");
                                //Сохранение в базе и разблокировка
                                tdmsObj.UnlockCheckIn(1);
                                TraceSource.TraceMessage(TraceType.Information, "Чертёж сохранён в TDMS и разблокирован.");
                                tdmsObj.Update();
                                TraceSource.TraceMessage(TraceType.Information, "Чертёж загружен и обновлён в TDMS успешно.");
                            }
                            else
                            {
                                Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog("Невозможно сохранить в TDMS, т.к. документ открыт на просмотр! Сохраните файл на локальном диске.");
                                TraceSource.TraceMessage(TraceType.Information, "Невозможно сохранить в TDMS, т.к. документ открыт на просмотр! Сохраните файл на локальном диске.");
                                Technical.TitleDoc();
                            }
                        }
                        else
                        {
                            Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog("\n Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                            TraceSource.TraceMessage(TraceType.Information, "Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                        }
                    }
                    else
                    {
                        TraceSource.TraceMessage(TraceType.Information, "Чертёж не из TDMS, сохранён локально: " + _doc.Name + " и закрыт в AutoCAD.");
                        _doc.CloseAndSave(_doc.Name);
                    }
                }
                else
                {
                    try
                    {
                        var fullPath = SaveDialog(_doc);
                        TraceSource.TraceMessage(TraceType.Information, "Чертёж не из TDMS и не был сохранён ранее, сохранён локально: " + fullPath + " и закрыт в AutoCAD.");
                        _doc.CloseAndSave(fullPath);
                    }
                    catch(System.Exception ex)
                    {
                        TraceSource.TraceMessage(TraceType.Information, "Команда отменена " + ex.ToString());
                        _editor.WriteMessage("\n Команда отменена \n");
                    }
                }
                TraceSource.TraceMessage(TraceType.Information, "[-----------------------------]");
            }
            catch (System.Exception ex)
            {
                TraceSource.TraceMessage(TraceType.Information, ex.ToString());
                _editor.WriteMessage("\n Exception caught: \n" + ex);
            }
        }

        [CommandMethod("CloseAndDiscardAll", CommandFlags.Session | CommandFlags.NoBlockEditor)]
        public void CloseAndDiscardAll()
        {
            TraceSource.TraceMessage(TraceType.Information, "[---- Закрыть всё без сохранения ----]");
            if (System.Windows.Forms.MessageBox.Show(@"Действительно хотите закрыть все чертежи без сохранения изменений? ", @"Закрыть все чертежи", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question) != System.Windows.Forms.DialogResult.Yes)
            {
                return;
            }

            var check = new Condition();

            try
            {
                DocumentCollection docs = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager;
                foreach (Document doc in docs)
                {
                    if (check.CheckPath(doc.Name))
                    {
                        TraceSource.TraceMessage(TraceType.Information, "Чертёж принадлежит TDMS. Путь сохранения файла совпадает с путём " + @"C:\TEMP\");

                        if (check.CheckTdmsProcess())
                        {
                            TraceSource.TraceMessage(TraceType.Information, "TDMS запущен.");

                            var tdmsApp = new TDMSApplication();

                            var strDwgName = doc.Name;

                            //Получение объекта по GUID
                            var tdmsObj = tdmsApp.GetObjectByGUID(check.ParseGuid(strDwgName));

                            //Заблокировани ли чертёж текущим пользователем
                            var lockOwner = tdmsObj.Permissions.LockOwner;
                            if (lockOwner)
                            {
                                TraceSource.TraceMessage(TraceType.Information, "Чертёж открыт для редактирования.");

                                TraceSource.TraceMessage(TraceType.Information, "Чертёж " + doc.Name + " разблокирован в TDMS.");
                                tdmsObj.UnlockCheckIn(0);

                                TraceSource.TraceMessage(TraceType.Information, "Чертёж " + doc.Name + " закрыт без изменений.");
                                doc.CloseAndDiscard();

                                tdmsObj.Update();
                                TraceSource.TraceMessage(TraceType.Information, "Чертёж обновлён в TDMS успешно.");
                            }
                            else
                            {
                                TraceSource.TraceMessage(TraceType.Information, "Чертёж " + doc.Name + " закрыт без изменений.");
                                doc.CloseAndDiscard();
                            }
                        }
                        else
                        {
                            TraceSource.TraceMessage(TraceType.Information, "Чертёж не принадлежит TDMS. И должен быть сохранён локально.");
                            Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog("Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                        }
                    }
                    else
                    {
                        TraceSource.TraceMessage(TraceType.Information, "Чертёж " + doc.Name + " закрыт без изменений.");
                        doc.CloseAndDiscard();
                        Thread.Sleep(TimeSpan.FromSeconds(1));
                    }
                }
                TraceSource.TraceMessage(TraceType.Information, "[------------------------------------]");
            }
            catch (System.Exception ex)
            {
                TraceSource.TraceMessage(TraceType.Error, ex.ToString());
                _editor.WriteMessage("\n Exception caught: \n" + ex);
            }
        }

        [CommandMethod("SaveAndCloseAll", CommandFlags.Session | CommandFlags.NoBlockEditor)]
        public void SaveAndCloseAll()
        {
            TraceSource.TraceMessage(TraceType.Information, "[---- Сохранить и закрыть все чертежи ----]");

            if (System.Windows.Forms.MessageBox.Show(@"Действительно хотите сохранить и закрыть все чертежи?", @"Сохранить и закрыть все чертежи", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question) != System.Windows.Forms.DialogResult.Yes)
            {
                return;
            }

            var check = new Condition();
            try
            {
                DocumentCollection docs = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager;
                foreach (Document doc in docs)
                {
                    var strDwgName = doc.Name;
                    if (!string.IsNullOrEmpty(Path.GetDirectoryName(strDwgName)))
                    {
                        TraceSource.TraceMessage(TraceType.Information, "Путь к чертежу не пустой, чертёж уже был сохранён ранее.");

                        if (check.CheckPath(strDwgName))
                        {
                            TraceSource.TraceMessage(TraceType.Information, "Чертёж принадлежит TDMS. Путь сохранения файла совпадает с путём " + @"C:\TEMP\");

                            if (check.CheckTdmsProcess())
                            {
                                TraceSource.TraceMessage(TraceType.Information, "TDMS запущен.");
                                var tdmsApp = new TDMSApplication();

                                //Получение объекта по GUID
                                var tdmsObj = tdmsApp.GetObjectByGUID(check.ParseGuid(strDwgName));

                                //Заблокировани ли чертёж текущим пользователем
                                var lockOwner = tdmsObj.Permissions.LockOwner;
                                if (lockOwner)
                                {
                                    TraceSource.TraceMessage(TraceType.Information, "Чертёж открыт для редактирования.");
                                    doc.Database.CloseInput(true);
                                    doc.CloseAndSave(doc.Name);
                                    TraceSource.TraceMessage(TraceType.Information, "Чертёж сохранён закрыт успешно.");
                                    //Сохранение в базе и разблокировка
                                    tdmsObj.UnlockCheckIn(1);
                                    TraceSource.TraceMessage(TraceType.Information, "Чертёж сохранён в TDMS и разблокирован.");
                                    tdmsObj.Update();
                                    TraceSource.TraceMessage(TraceType.Information, "Чертёж загружен и обновлён в TDMS успешно.");
                                }
                                else
                                {
                                    Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog("Невозможно сохранить в TDMS, т.к. документ открыт на просмотр! Сохраните файл на локальном диске.");
                                    TraceSource.TraceMessage(TraceType.Information, "Невозможно сохранить в TDMS, т.к. документ открыт на просмотр! Сохраните файл на локальном диске.");
                                    Technical.TitleDoc();
                                }
                            }
                            else
                            {
                                Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog("\n Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                                TraceSource.TraceMessage(TraceType.Information, "Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                            }
                        }
                        else
                        {
                            TraceSource.TraceMessage(TraceType.Information, "Чертёж не из TDMS, сохранён локально: " + doc.Name + " и закрыт в AutoCAD.");
                            doc.CloseAndSave(doc.Name);
                        }
                    }
                    else
                    {
                        try
                        {
                            var fullPath = SaveDialog(doc);
                            TraceSource.TraceMessage(TraceType.Information, "Чертёж не из TDMS и не был сохранён ранее, сохранён локально: " + fullPath + " и закрыт в AutoCAD.");
                            doc.CloseAndSave(fullPath);
                        }
                        catch (System.Exception ex)
                        {
                            TraceSource.TraceMessage(TraceType.Information, "Команда сохранения отменена " + ex);
                            _editor.WriteMessage("\n Команда сохранения отменена \n");
                        }
                    }
                }
                TraceSource.TraceMessage(TraceType.Information, "[-----------------------------]");
            }
            catch (System.Exception ex)
            {
                TraceSource.TraceMessage(TraceType.Information, ex.ToString());
                _editor.WriteMessage("\n Exception caught: \n" + ex);
            }
        }

        #endregion Методы, реализующие кнопки для сохранения чертежей

        #region Методы, реализующие сохранение и закрытие редактируемой внешней ссылки в пространстве чертежа

        ///<summary>
        /// Сохранение и закрытие редактируемой внешней ссылки в пространстве чертежа
        ///</summary>
        [CommandMethod("SaveCloseXREFInPlace", CommandFlags.Session | CommandFlags.NoBlockEditor)]
        public void SaveCloseXrefInPlace()
        {
            TraceSource.TraceMessage(TraceType.Information, "[---- Сохранить и закрыть редактируемую внешнюю ссылку ----]");
            var tdmsObjXref = Xref.GetTdmsObjXref();
            var check = new Condition();
            try
            {
                if (check.CheckPath())
                {
                    if (tdmsObjXref != null)
                    {
                        if (check.CheckTdmsProcess())
                        {
                            //Сохранение изменений в чертеже - внешней ссылке
                            _doc.SendStringToExecute("_refclose" + "\n", true, false, false);
                            _doc.SendStringToExecute("_sav" + "\n", true, false, false);

                            //Загрузка в базу ТДМС

                            tdmsObjXref.UnlockCheckIn(1);
                            tdmsObjXref.Update();
                        }
                        else
                        {
                            Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog("\n Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                        }
                    }
                    else
                    {
                        _editor.WriteMessage("\n Ссылка не выбрана.");
                    }
                }
                else
                {
                    _editor.WriteMessage("\n Внешняя ссылка не принадлежит TDMS, но изменения будут сохранены.");
                    //Сохранение изменений в чертеже - внешней ссылке
                    _doc.SendStringToExecute("_refclose" + "\n", true, false, false);
                    _doc.SendStringToExecute("_sav" + "\n", true, false, false);
                }
            }
            catch (System.Exception ex)
            {
                _editor.WriteMessage("\n Exception caught: \n" + ex);
            }
        }

        ///<summary>
        /// Закрытие редактируемой внешней ссылки в пространстве чертежа без сохранения
        ///</summary>
        [CommandMethod("CloseXREFInPlace", CommandFlags.Session | CommandFlags.NoBlockEditor)]
        public void CloseXrefInPlace()
        {
            TraceSource.TraceMessage(TraceType.Information, "[---- Закрыть внешнюю ссылку без сохранения ----]");
            var tdmsObjXref = Xref.GetTdmsObjXref();
            var check = new Condition();
            try
            {
                if (check.CheckPath())
                {
                    if (tdmsObjXref != null)
                    {
                        if (check.CheckTdmsProcess())
                        {
                            //Закрытие внешней ссылки в чертеже без сохранения изменений
                            _doc.SendStringToExecute("_refclose" + "\n", true, false, false);
                            _doc.SendStringToExecute("_disc" + "\n", true, false, false);

                            //Загрузка в базу ТДМС
                            tdmsObjXref.UnlockCheckIn(0);
                            tdmsObjXref.Update();
                        }
                        else
                        {
                            Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog("\n Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                        }
                    }
                    else
                    {
                        _editor.WriteMessage("\n Ссылка не выбрана.");
                    }
                }
                else
                {
                    _doc.SendStringToExecute("_refclose" + "\n", true, false, false);
                    _doc.SendStringToExecute("_disc" + "\n", true, false, false);
                }
            }
            catch (System.Exception ex)
            {
                _editor.WriteMessage("\n Exception caught: \n" + ex);
            }
        }

        #endregion Методы, реализующие сохранение и закрытие редактируемой внешней ссылки в пространстве чертежа
    }
}