namespace Auto
{
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.EditorInput;
    using Autodesk.AutoCAD.Geometry;
    using Autodesk.AutoCAD.Runtime;
    using Microsoft.Win32;
    using System.Reflection;
    using System.Text.RegularExpressions;
    using TDMS.Interop;

    public sealed class Starter : IExtensionApplication
    {
        public void Initialize()
        {
            try
            {
                TraceSource.CreateTraceFile();

                ContextMenu cMenu = new ContextMenu();
                cMenu.Initialize();

                TraceSource.TraceMessage(TraceType.Information, "Context menu tdms is created");

                CreateRibbon myApp = new CreateRibbon();
                myApp.Initialize();

                TraceSource.TraceMessage(TraceType.Information, "Ribbon panel tdms is created");
            }
            catch (System.Exception ex)
            {
            }
        }

        public void Terminate()
        {
        }
    }

    public sealed class Technical : IExtensionApplication
    {
        private const string textStyle = "STAMP";
        private static Document doc = Application.DocumentManager.MdiActiveDocument;
        private Editor editor = doc.Editor;

        public void Initialize()
        {
            try
            {
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        public void Terminate()
        {
        }

        public static void TitleDoc()
        {
            Editor editor = doc.Editor;
            Condition Check = new Condition();
            try
            {
                if (Check.CheckPath())
                {
                    if (Check.CheckTdmsProcess())
                    {
                        TDMSApplication tdmsApp = new TDMSApplication();
                        TDMSObject tdmsObj = null;
                        Document acDoc = Application.DocumentManager.MdiActiveDocument;
                        object obj = Application.GetSystemVariable("DWGTITLED");
                        string strDwgName = acDoc.Name;

                        //Возвращение пути к файлу чертежа
                        string guid = strDwgName;

                        //парсинг пути, получение актуального GUID объекта
                        string parseGuid = null;
                        Regex reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                        MatchCollection mc = reg.Matches(guid);
                        foreach (Match mat in mc)
                        {
                            parseGuid += mat.Value.ToString();
                        }
                        parseGuid = parseGuid.Remove(0, 38);
                        //Получение объекта по GUID
                        tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);
                        bool lockOwner = tdmsObj.Permissions.LockOwner;
                        //Заблокирован ли чертёж текущим пользователем
                        if (lockOwner)
                        {
                            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Window.Text += " TDMS - редактирование!";
                        }
                        else
                        {
                            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Window.Text += " TDMS - просмотр!";
                        }
                    }
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        //Атрибут Автокада, ссылающийся на функцию, должен иметь формат:
        //CMD_SYSLIB - имя модуля
        //GetTitle - имя функции
        //далее идут имена параметров
        //Вызов происходит так
        //attrValue = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetDataFromObjForAcad", tdmsObj, "A_STAGE_CLSF");
        //запускаем создание атрибута из функции, с разными параметрами (метод перегружен!!!)
        //проверяем, запущен ли процесс "TDMS.exe" возвращаем GUID чертежа, запускаем создание атрибута и присваиваем ему значение из объекта ТДМС.
        //без параметра
        public static void AddMultilineAttributeFunction(double x, double y, double height, double width, double widthFactor, double rotate, double oblique, string moduleName, string functionName, int blockName)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            Condition Check = new Condition();
            string attrName = moduleName + "_" + functionName + "#FUNCTION";
            editor.WriteMessage("\n ADDRESS: " + attrName);
            try
            {
                if (Check.CheckPath())
                {
                    if (Check.CheckTdmsProcess())
                    {
                        string attrValue = null;

                        TDMSApplication tdmsApp = new TDMSApplication();

                        TDMSObject tdmsObj = null;
                        Document acDoc = Application.DocumentManager.MdiActiveDocument;
                        object obj = Application.GetSystemVariable("DWGTITLED");
                        string strDwgName = acDoc.Name;
                        string guid = strDwgName;
                        string parseGuid = null;

                        Regex reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                        MatchCollection mc = reg.Matches(guid);
                        foreach (Match mat in mc)
                        {
                            parseGuid += mat.Value.ToString();
                        }
                        parseGuid = parseGuid.Remove(0, 38);
                        editor.WriteMessage("\n ADDRESS: " + parseGuid);
                        tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);
                        editor.WriteMessage("\n ADDRESS: TDMSOBJECT;");
                        attrValue = tdmsApp.ExecuteScript(moduleName, functionName, tdmsObj);
                        editor.WriteMessage("\n ADDRESS: " + attrValue);
                        if (attrValue != "")
                        {
                            Creator.CreateMultilineStampAtribut(x, y, attrValue, height, width, widthFactor, rotate, oblique, attrName, blockName, textStyle);
                        }
                        else { Creator.CreateMultilineStampAtribut(x, y, "Text", height, width, widthFactor, rotate, oblique, attrName, blockName, textStyle); }
                    }
                    else { Creator.CreateMultilineStampAtribut(x, y, "Text", height, width, widthFactor, rotate, oblique, attrName, blockName, textStyle); }
                }
                else { Creator.CreateMultilineStampAtribut(x, y, "Text", height, width, widthFactor, rotate, oblique, attrName, blockName, textStyle); }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        //1 параметр
        public static void AddAttributeFunction(double x, double y, double height, double widthFactor, double rotate, double oblique, string moduleName, string functionName, int blockName, string parameter)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            Condition Check = new Condition();
            string attrName = moduleName + "_" + functionName + "|" + parameter + "#FUNCTION";

            try
            {
                if (Check.CheckPath())
                {
                    if (Check.CheckTdmsProcess())
                    {
                        AttributeDefinition adAttr = new AttributeDefinition();
                        string attrValue = null;

                        TDMSApplication tdmsApp = new TDMSApplication();

                        TDMSObject tdmsObj = null;
                        Document acDoc = Application.DocumentManager.MdiActiveDocument;
                        string strDwgName = acDoc.Name;
                        string guid = strDwgName;
                        string parseGuid = null;

                        Regex reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                        MatchCollection mc = reg.Matches(guid);
                        foreach (Match mat in mc)
                        {
                            parseGuid += mat.Value.ToString();
                        }
                        parseGuid = parseGuid.Remove(0, 38);

                        tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);
                        if (parameter != "GKAB_")
                        {
                            attrValue = tdmsApp.ExecuteScript(moduleName, functionName, tdmsObj, parameter);
                        }
                        else { attrValue = tdmsApp.ExecuteScript(moduleName, functionName, parameter); }

                        if (attrValue != "")
                        {
                            Creator.CreateStampAtribut(x, y, attrValue, height, widthFactor, rotate, oblique, attrName, blockName, textStyle);
                        }
                        else { Creator.CreateStampAtribut(x, y, "Text", height, widthFactor, rotate, oblique, attrName, blockName, textStyle); }
                    }
                    else { Creator.CreateStampAtribut(x, y, "Text", height, widthFactor, rotate, oblique, attrName, blockName, textStyle); }
                }
                else { Creator.CreateStampAtribut(x, y, "Text", height, widthFactor, rotate, oblique, attrName, blockName, textStyle); }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        // 2 дополнительных параметра
        public static void AddAttributeFunction(double x, double y, double height, double widthFactor, double rotate, double oblique, string moduleName, string functionName, int blockName, string parameter, string parameter2)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            Condition Check = new Condition();
            string attrName = moduleName + "_" + functionName + "|" + parameter + "|" + parameter2 + "#FUNCTION";
            editor.WriteMessage("\n" + attrName);
            try
            {
                if (Check.CheckPath())
                {
                    if (Check.CheckTdmsProcess())
                    {
                        string attrValue = null;
                        TDMSApplication tdmsApp = new TDMSApplication();
                        TDMSObject tdmsObj = null;
                        Document acDoc = Application.DocumentManager.MdiActiveDocument;
                        object obj = Application.GetSystemVariable("DWGTITLED");
                        string strDwgName = acDoc.Name;
                        string guid = strDwgName;
                        string parseGuid = null;

                        Regex reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                        MatchCollection mc = reg.Matches(guid);
                        foreach (Match mat in mc)
                        {
                            parseGuid += mat.Value.ToString();
                        }
                        parseGuid = parseGuid.Remove(0, 38);

                        tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);

                        attrValue = tdmsApp.ExecuteScript(moduleName, functionName, tdmsObj, parameter, parameter2);

                        editor.WriteMessage("\n Значение" + attrValue);
                        if (attrValue != "")
                        {
                            Creator.CreateStampAtribut(x, y, attrValue, height, widthFactor, rotate, oblique, attrName, blockName, textStyle);
                        }
                        else { Creator.CreateStampAtribut(x, y, "Text", height, widthFactor, rotate, oblique, attrName, blockName, textStyle); }
                    }
                    else { Creator.CreateStampAtribut(x, y, "Text", height, widthFactor, rotate, oblique, attrName, blockName, textStyle); }
                }
                else { Creator.CreateStampAtribut(x, y, "Text", height, widthFactor, rotate, oblique, attrName, blockName, textStyle); }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        //Добавляем обычный атрибут НЕ ФУНКЦИЯ!
        //проверяем, запущен ли процесс "TDMS.exe" возвращаем GUID чертежа, запускаем создание атрибута и присваиваем ему значение из объекта ТДМС.
        public static void AddAttribute(double x, double y, double height, double widthFactor, double rotate, double oblique, string attrName, int blockName)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            Condition Check = new Condition();
            try
            {
                if (Check.CheckPath())
                {
                    if (Check.CheckTdmsProcess())
                    {
                        AttributeDefinition adAttr = new AttributeDefinition();
                        string attrValue = null;
                        TDMSApplication tdmsApp = new TDMSApplication();
                        TDMSObject tdmsObj = null;
                        Document acDoc = Application.DocumentManager.MdiActiveDocument;
                      
                        string strDwgName = acDoc.Name;
                        string guid = strDwgName;
                        string parseGuid = null;
                        Regex reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                        MatchCollection mc = reg.Matches(guid);
                        foreach (Match mat in mc)
                        {
                            parseGuid += mat.Value.ToString();
                        }
                        parseGuid = parseGuid.Remove(0, 38);
                        tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);
                        attrValue = tdmsObj.Attributes[attrName].Value;
                        if (attrValue != "")
                        {
                            Creator.CreateStampAtribut(x, y, attrValue, height, widthFactor, rotate, oblique, attrName, blockName, textStyle);
                        }
                        else { Creator.CreateStampAtribut(x, y, "Text", height, widthFactor, rotate, oblique, attrName, blockName, textStyle); }
                    }
                    else { Creator.CreateStampAtribut(x, y, "Text", height, widthFactor, rotate, oblique, attrName, blockName, textStyle); }
                }
                else { Creator.CreateStampAtribut(x, y, "Text", height, widthFactor, rotate, oblique, attrName, blockName, textStyle); }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        // Многострочный атрибут без параметра
        public static void AddAttributeMultiline(double x, double y, double height, double width, double widthFactor, double rotate, double oblique, string attrName, int blockName)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            Condition Check = new Condition();
            try
            {
                if (Check.CheckPath())
                {
                    if (Check.CheckTdmsProcess())
                    {
                        AttributeDefinition adAttr = new AttributeDefinition();
                        string attrValue = null;
                        TDMSApplication tdmsApp = new TDMSApplication();
                        TDMSObject tdmsObj = null;
                        Document acDoc = Application.DocumentManager.MdiActiveDocument;
                     
                        string strDwgName = acDoc.Name;
                        string guid = strDwgName;
                        string parseGuid = null;
                        Regex reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                        MatchCollection mc = reg.Matches(guid);
                        foreach (Match mat in mc)
                        {
                            parseGuid += mat.Value.ToString();
                        }
                        parseGuid = parseGuid.Remove(0, 38);
                        tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);
                        attrValue = tdmsObj.Attributes[attrName].Value;
                        if (attrValue != "")
                        {
                            Creator.CreateMultilineStampAtribut(x, y, attrValue, height, width, widthFactor, rotate, oblique, attrName, blockName, textStyle);
                        }
                        else { Creator.CreateMultilineStampAtribut(x, y, "Text", height, width, widthFactor, rotate, oblique, attrName, blockName, textStyle); }
                    }
                    else { Creator.CreateMultilineStampAtribut(x, y, "Text", height, width, widthFactor, rotate, oblique, attrName, blockName, textStyle); }
                }
                else { Creator.CreateMultilineStampAtribut(x, y, "Text", height, width, widthFactor, rotate, oblique, attrName, blockName, textStyle); }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        //Обновление атрибутов в уже вставленных блоках. Атрибуты берутся из ТДМС.
        public void RefreshAttribute(string attrName)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor editor = doc.Editor;
            string blockName = "ATTRBLK";

            TDMSApplication tdmsApp = new TDMSApplication();
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            try
            {
                string strDwgName = acDoc.Name;
                string guid = strDwgName;
                string parseGuid = null;
                Regex reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                MatchCollection mc = reg.Matches(guid);
                foreach (Match mat in mc)
                {
                    parseGuid += mat.Value.ToString();
                }
                parseGuid = parseGuid.Remove(0, 38);
                var tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);
                string attrValue = tdmsObj.Attributes[attrName].Value;
                if (attrValue != "")
                {
                    Commands refresh = new Commands();
                    refresh.UpdateAttributesInDatabase(db, blockName, attrName, attrValue);
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        //Обновление атрибута - функции без параметров
        public void RefreshAttribute(string moduleName, string functionName)
        {
            string attrName = moduleName + "_" + functionName + "#FUNCTION";
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor editor = doc.Editor;
            string blockName = "ATTRBLK";
          
            string attrValue = null;
            TDMSApplication tdmsApp = new TDMSApplication();
            TDMSObject tdmsObj = null;

            try
            {
                object obj = Application.GetSystemVariable("DWGTITLED");
                string strDwgName = doc.Name;
                string guid = strDwgName;
                string parseGuid = null;
                Regex reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                MatchCollection mc = reg.Matches(guid);
                foreach (Match mat in mc)
                {
                    parseGuid += mat.Value.ToString();
                }
                parseGuid = parseGuid.Remove(0, 38);
                tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);

                attrValue = tdmsApp.ExecuteScript(moduleName, functionName, tdmsObj);

                if (attrValue != "")
                {
                    Commands refresh = new Commands();
                    refresh.UpdateAttributesInDatabase(db, blockName, attrName, attrValue);
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        //Обновление атрибута - функции с 1 параметром
        public void RefreshAttribute(string moduleName, string functionName, string parameter)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor editor = doc.Editor;

            string blockName = "ATTRBLK";

            TDMSApplication tdmsApp = new TDMSApplication();
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            try
            {
              
                string attrName = moduleName + "_" + functionName + "|" + parameter + "#FUNCTION";
                string strDwgName = acDoc.Name;
                string guid = strDwgName;
                string parseGuid = null;
                Regex reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                MatchCollection mc = reg.Matches(guid);
                foreach (Match mat in mc)
                {
                    parseGuid += mat.Value.ToString();
                }
                parseGuid = parseGuid.Remove(0, 38);
                var tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);

                string attrValue = null;
                if (parameter != "GKAB_")
                {
                    attrValue = tdmsApp.ExecuteScript(moduleName, functionName, tdmsObj, parameter);
                }
                else { attrValue = tdmsApp.ExecuteScript(moduleName, functionName, parameter); }

                if (attrValue != "")
                {
                    Commands refresh = new Commands();
                    refresh.UpdateAttributesInDatabase(db, blockName, attrName, attrValue);
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        //Обновление атрибута - функции с 2 параметрами
        public void RefreshAttribute(string moduleName, string functionName, string parameter, string parameter2)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                string attrName = moduleName + "_" + functionName + "|" + parameter + "|" + parameter2 + "#FUNCTION";

                Document doc = Application.DocumentManager.MdiActiveDocument;

                Database db = doc.Database;

                string blockName = "ATTRBLK";

                string attrValue = null;
                TDMSApplication tdmsApp = new TDMSApplication();
                TDMSObject tdmsObj = null;
                Document acDoc = Application.DocumentManager.MdiActiveDocument;
              
                string strDwgName = acDoc.Name;
                string guid = strDwgName;
                string parseGuid = null;

                Regex reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                MatchCollection mc = reg.Matches(guid);
                foreach (Match mat in mc)
                {
                    parseGuid += mat.Value.ToString();
                }
                parseGuid = parseGuid.Remove(0, 38);
                tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);

                attrValue = tdmsApp.ExecuteScript(moduleName, functionName, tdmsObj, parameter, parameter2);
                if (attrValue != "")
                {
                    Commands refresh = new Commands();
                    refresh.UpdateAttributesInDatabase(db, blockName, attrName, attrValue);
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        //Добавляем логотип
        public void AddLogo(double x, double y)
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Editor editor = acDoc.Editor;

            try
            {
                Database acCurDb = acDoc.Database;

                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.PaperSpace], OpenMode.ForWrite) as BlockTableRecord;

                    //////////////////////////////////////////////////////////////////////////////////////////
                    int roomNumPoint = 0;
                    Polyline Poly = new Polyline();
                    Poly.SetDatabaseDefaults();
                    Poly.AddVertexAt(roomNumPoint, new Point2d(x - 48.258, y + 6.215), 0, 0, 0);
                    Poly.AddVertexAt(++roomNumPoint, new Point2d(x - 44.240, y + 11.958), 0, 0, 0);
                    Poly.AddVertexAt(++roomNumPoint, new Point2d(x - 48.258, y + 11.958), 0, 0, 0);
                    Poly.Closed = true;
                    Poly.LineWeight = 0;
                    Poly.ConstantWidth = 0;
                    Poly.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly);
                    acTrans.AddNewlyCreatedDBObject(Poly, true);

                    ObjectIdCollection acObjIdColl = new ObjectIdCollection();
                    acObjIdColl.Add(Poly.ObjectId);

                    //////////////////////////////////////////////////////////////////////////////////////////
                    roomNumPoint = 0;
                    Polyline Poly1 = new Polyline();
                    Poly1.SetDatabaseDefaults();
                    Poly1.AddVertexAt(roomNumPoint, new Point2d(x - 45.346, y + 9.432), 0, 0, 0);
                    Poly1.AddVertexAt(++roomNumPoint, new Point2d(x - 47.597, y + 6.215), 0, 0, 0);
                    Poly1.AddVertexAt(++roomNumPoint, new Point2d(x - 45.346, y + 6.215), 0, 0, 0);
                    Poly1.Closed = true;
                    Poly1.LineWeight = 0;
                    Poly1.ConstantWidth = 0;
                    Poly1.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly1);
                    acTrans.AddNewlyCreatedDBObject(Poly1, true);

                    ObjectIdCollection acObjIdColl1 = new ObjectIdCollection();
                    acObjIdColl1.Add(Poly1.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    roomNumPoint = 0;
                    Polyline Poly2 = new Polyline();
                    Poly2.SetDatabaseDefaults();
                    Poly2.AddVertexAt(roomNumPoint, new Point2d(x - 45.346, y + 2.989), 0, 0, 0);
                    Poly2.AddVertexAt(++roomNumPoint, new Point2d(x - 45.346, y + 4.673), 0, 0, 0);
                    Poly2.AddVertexAt(++roomNumPoint, new Point2d(x - 48.258, y + 4.673), 0, 0, 0);
                    Poly2.AddVertexAt(++roomNumPoint, new Point2d(x - 48.258, y + 2.989), 0, 0, 0);
                    Poly2.Closed = true;
                    Poly2.LineWeight = 0;
                    Poly2.ConstantWidth = 0;
                    Poly2.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly2);
                    acTrans.AddNewlyCreatedDBObject(Poly2, true);

                    ObjectIdCollection acObjIdColl2 = new ObjectIdCollection();
                    acObjIdColl2.Add(Poly2.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    roomNumPoint = 0;
                    Polyline Poly3 = new Polyline();
                    Poly3.SetDatabaseDefaults();
                    Poly3.AddVertexAt(roomNumPoint, new Point2d(x - 39.360, y + 2.989), 0, 0, 0);
                    Poly3.AddVertexAt(++roomNumPoint, new Point2d(x - 39.360, y + 4.673), 0, 0, 0);
                    Poly3.AddVertexAt(++roomNumPoint, new Point2d(x - 42.272, y + 4.673), 0, 0, 0);
                    Poly3.AddVertexAt(++roomNumPoint, new Point2d(x - 42.272, y + 6.215), 0, 0, 0);
                    Poly3.AddVertexAt(++roomNumPoint, new Point2d(x - 38.253, y + 11.958), 0, 0, 0);
                    Poly3.AddVertexAt(++roomNumPoint, new Point2d(x - 43.347, y + 11.958), 0, 0, 0);
                    Poly3.AddVertexAt(++roomNumPoint, new Point2d(x - 43.347, y + 6.215), 0, 0, 0);
                    Poly3.AddVertexAt(++roomNumPoint, new Point2d(x - 42.608, y + 6.215), 0, 0, 0);
                    Poly3.AddVertexAt(++roomNumPoint, new Point2d(x - 42.608, y + 4.673), 0, 0, 0);
                    Poly3.AddVertexAt(++roomNumPoint, new Point2d(x - 43.347, y + 4.673), 0, 0, 0);
                    Poly3.AddVertexAt(++roomNumPoint, new Point2d(x - 43.347, y + 2.989), 0, 0, 0);
                    Poly3.Closed = true;
                    Poly3.LineWeight = 0;
                    Poly3.ConstantWidth = 0;
                    Poly3.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly3);
                    acTrans.AddNewlyCreatedDBObject(Poly3, true);

                    ObjectIdCollection acObjIdColl3 = new ObjectIdCollection();
                    acObjIdColl3.Add(Poly3.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    roomNumPoint = 0;
                    Polyline Poly4 = new Polyline();
                    Poly4.SetDatabaseDefaults();
                    Poly4.AddVertexAt(roomNumPoint, new Point2d(x - 39.360, y + 6.215), 0, 0, 0);
                    Poly4.AddVertexAt(++roomNumPoint, new Point2d(x - 39.360, y + 9.432), 0, 0, 0);
                    Poly4.AddVertexAt(++roomNumPoint, new Point2d(x - 41.610, y + 6.215), 0, 0, 0);
                    Poly4.Closed = true;
                    Poly4.LineWeight = 0;
                    Poly4.ConstantWidth = 0;
                    Poly4.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly4);
                    acTrans.AddNewlyCreatedDBObject(Poly4, true);

                    ObjectIdCollection acObjIdColl4 = new ObjectIdCollection();
                    acObjIdColl4.Add(Poly4.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    roomNumPoint = 0;
                    Polyline Poly5 = new Polyline();
                    Poly5.SetDatabaseDefaults();
                    Poly5.AddVertexAt(roomNumPoint, new Point2d(x - 36.674, y + 11.958), 0, 0, 0);
                    Poly5.AddVertexAt(++roomNumPoint, new Point2d(x - 37.361, y + 11.958), 0, 0, 0);
                    Poly5.AddVertexAt(++roomNumPoint, new Point2d(x - 37.361, y + 6.215), 0, 0, 0);
                    Poly5.AddVertexAt(++roomNumPoint, new Point2d(x - 36.674, y + 6.215), 0, 0, 0);
                    Poly5.Closed = true;
                    Poly5.LineWeight = 0;
                    Poly5.ConstantWidth = 0;
                    Poly5.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly5);
                    acTrans.AddNewlyCreatedDBObject(Poly5, true);

                    ObjectIdCollection acObjIdColl5 = new ObjectIdCollection();
                    acObjIdColl5.Add(Poly5.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    roomNumPoint = 0;
                    Polyline Poly6 = new Polyline();
                    Poly6.SetDatabaseDefaults();
                    Poly6.AddVertexAt(roomNumPoint, new Point2d(x - 36.674, y + 4.673), 0, 0, 0);
                    Poly6.AddVertexAt(++roomNumPoint, new Point2d(x - 37.361, y + 4.673), 0, 0, 0);
                    Poly6.AddVertexAt(++roomNumPoint, new Point2d(x - 37.361, y + 2.989), 0, 0, 0);
                    Poly6.AddVertexAt(++roomNumPoint, new Point2d(x - 36.674, y + 2.989), 0, 0, 0);
                    Poly6.Closed = true;
                    Poly6.LineWeight = 0;
                    Poly6.ConstantWidth = 0;
                    Poly6.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly6);
                    acTrans.AddNewlyCreatedDBObject(Poly6, true);

                    ObjectIdCollection acObjIdColl6 = new ObjectIdCollection();
                    acObjIdColl6.Add(Poly6.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    // C
                    roomNumPoint = 0;
                    Polyline Poly7 = new Polyline();
                    Poly7.SetDatabaseDefaults();
                    Poly7.AddVertexAt(roomNumPoint, new Point2d(x - 47.080, y + 13.763), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.031, y + 13.963), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.146, y + 14.006), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.230, y + 14.029), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.324, y + 14.044), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.423, y + 14.054), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.552, y + 14.055), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.693, y + 14.036), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.812, y + 13.999), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.939, y + 13.938), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 48.065, y + 13.845), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 48.160, y + 13.728), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 48.217, y + 13.614), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 48.253, y + 13.486), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 48.262, y + 13.424), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 48.262, y + 13.244), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 48.238, y + 13.117), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 48.202, y + 13.037), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 48.161, y + 12.972), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 48.097, y + 12.900), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 48.016, y + 12.826), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.899, y + 12.757), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.765, y + 12.707), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.665, y + 12.684), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.589, y + 12.676), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.394, y + 12.675), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.267, y + 12.693), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.165, y + 12.720), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.102, y + 12.750), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.082, y + 12.771), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.056, y + 12.863), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.045, y + 12.898), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.047, y + 12.921), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.067, y + 12.933), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.123, y + 12.894), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.190, y + 12.857), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.283, y + 12.817), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.361, y + 12.799), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.517, y + 12.799), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.571, y + 12.808), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.693, y + 12.856), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.727, y + 12.879), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.814, y + 12.964), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.884, y + 13.095), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.903, y + 13.163), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.920, y + 13.260), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.920, y + 13.447), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.906, y + 13.555), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.869, y + 13.642), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.825, y + 13.732), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.796, y + 13.778), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.709, y + 13.857), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.606, y + 13.912), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.520, y + 13.930), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.383, y + 13.930), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.286, y + 13.902), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.231, y + 13.871), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.171, y + 13.827), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.126, y + 13.783), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.108, y + 13.761), 0, 0, 0);
                    Poly7.Closed = true;
                    Poly7.LineWeight = 0;
                    Poly7.ConstantWidth = 0;
                    Poly7.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly7);
                    acTrans.AddNewlyCreatedDBObject(Poly7, true);

                    ObjectIdCollection acObjIdColl7 = new ObjectIdCollection();
                    acObjIdColl7.Add(Poly7.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    // T
                    roomNumPoint = 0;
                    Polyline Poly8 = new Polyline();
                    Poly8.SetDatabaseDefaults();
                    Poly8.AddVertexAt(roomNumPoint, new Point2d(x - 45.370, y + 13.885), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.370, y + 13.874), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.249, y + 13.874), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.249, y + 13.864), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.171, y + 13.864), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.171, y + 13.880), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.182, y + 13.880), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.182, y + 13.990), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.170, y + 13.994), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.170, y + 14.017), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 46.227, y + 14.017), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 46.227, y + 13.990), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 46.221, y + 13.990), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 46.221, y + 13.882), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 46.231, y + 13.882), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 46.231, y + 13.855), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 46.192, y + 13.855), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 46.192, y + 13.865), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 46.113, y + 13.865), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 46.113, y + 13.875), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 46.012, y + 13.875), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 46.012, y + 13.885), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.846, y + 13.885), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.846, y + 12.694), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.766, y + 12.694), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.766, y + 12.706), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.628, y + 12.706), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.628, y + 12.694), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.537, y + 12.694), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.537, y + 13.885), 0, 0, 0);
                    Poly8.Closed = true;
                    Poly8.LineWeight = 0;
                    Poly8.ConstantWidth = 0;
                    Poly8.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly8);
                    acTrans.AddNewlyCreatedDBObject(Poly8, true);

                    ObjectIdCollection acObjIdColl8 = new ObjectIdCollection();
                    acObjIdColl8.Add(Poly8.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    // У
                    roomNumPoint = 0;
                    Polyline Poly9 = new Polyline();
                    Poly9.SetDatabaseDefaults();
                    Poly9.AddVertexAt(roomNumPoint, new Point2d(x - 44.346, y + 14.026), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 44.270, y + 14.026), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 44.265, y + 14.017), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 44.069, y + 14.017), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 44.064, y + 14.026), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.979, y + 14.026), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.588, y + 13.366), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.488, y + 13.547), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.367, y + 13.793), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.259, y + 14.027), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.190, y + 14.027), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.190, y + 14.017), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.125, y + 14.017), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.125, y + 14.027), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.061, y + 14.027), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.378, y + 13.454), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.513, y + 13.187), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.757, y + 12.695), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.780, y + 12.695), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.780, y + 12.707), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.893, y + 12.707), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.893, y + 12.695), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.920, y + 12.695), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.920, y + 12.707), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.740, y + 13.045), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.820, y + 13.190), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.909, y + 13.341), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 44.003, y + 13.485), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 44.148, y + 13.712), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 44.346, y + 14.026), 0, 0, 0);
                    Poly9.Closed = true;
                    Poly9.LineWeight = 0;
                    Poly9.ConstantWidth = 0;
                    Poly9.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly9);
                    acTrans.AddNewlyCreatedDBObject(Poly9, true);

                    ObjectIdCollection acObjIdColl9 = new ObjectIdCollection();
                    acObjIdColl9.Add(Poly9.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    // Д и то что внутри
                    roomNumPoint = 0;
                    Polyline Poly10 = new Polyline();
                    Poly10.SetDatabaseDefaults();
                    Poly10.AddVertexAt(roomNumPoint, new Point2d(x - 42.238, y + 12.640), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 42.238, y + 12.653), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 42.227, y + 12.653), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 42.227, y + 12.821), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 42.238, y + 12.821), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 42.238, y + 12.847), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 42.150, y + 12.847), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.850, y + 13.518), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.646, y + 14.028), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.600, y + 14.028), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.600, y + 14.017), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.518, y + 14.017), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.518, y + 14.028), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.472, y + 14.028), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.320, y + 13.648), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 40.983, y + 12.847), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 40.888, y + 12.847), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 40.888, y + 12.821), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 40.899, y + 12.821), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 40.900, y + 12.653), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 40.888, y + 12.653), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 40.888, y + 12.640), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 40.921, y + 12.640), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 40.976, y + 12.668), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.108, y + 12.693), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 42.032, y + 12.693), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 42.126, y + 12.677), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 42.192, y + 12.655), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 42.201, y + 12.647), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 42.201, y + 12.640), 0, 0, 0);
                    Poly10.Closed = true;
                    Poly10.LineWeight = 0;
                    Poly10.ConstantWidth = 0;
                    Poly10.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly10);
                    acTrans.AddNewlyCreatedDBObject(Poly10, true);

                    ObjectIdCollection acObjIdColl10 = new ObjectIdCollection();
                    acObjIdColl10.Add(Poly10.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly_10 = new Polyline();
                    Poly_10.SetDatabaseDefaults();
                    Poly_10.AddVertexAt(roomNumPoint, new Point2d(x - 41.645, y + 13.639), 0, 0, 0);
                    Poly_10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.980, y + 12.833), 0, 0, 0);
                    Poly_10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.883, y + 12.833), 0, 0, 0);
                    Poly_10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.880, y + 12.819), 0, 0, 0);
                    Poly_10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.529, y + 12.819), 0, 0, 0);
                    Poly_10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.335, y + 12.850), 0, 0, 0);
                    Poly_10.Closed = true;
                    Poly_10.LineWeight = 0;
                    Poly_10.ConstantWidth = 0;
                    Poly_10.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly_10);
                    acTrans.AddNewlyCreatedDBObject(Poly_10, true);

                    ObjectIdCollection acObjIdColl_10 = new ObjectIdCollection();
                    acObjIdColl_10.Add(Poly_10.ObjectId);

                    //////////////////////////////////////////////////////////////////////////////////////////
                    // И
                    roomNumPoint = 0;
                    Polyline Poly11 = new Polyline();
                    Poly11.SetDatabaseDefaults();
                    Poly11.AddVertexAt(roomNumPoint, new Point2d(x - 39.934, y + 14.027), 0, 0, 0);
                    Poly11.AddVertexAt(++roomNumPoint, new Point2d(x - 39.934, y + 12.695), 0, 0, 0);
                    Poly11.AddVertexAt(++roomNumPoint, new Point2d(x - 39.786, y + 12.695), 0, 0, 0);
                    Poly11.AddVertexAt(++roomNumPoint, new Point2d(x - 39.003, y + 13.614), 0, 0, 0);
                    Poly11.AddVertexAt(++roomNumPoint, new Point2d(x - 39.003, y + 12.695), 0, 0, 0);
                    Poly11.AddVertexAt(++roomNumPoint, new Point2d(x - 38.693, y + 12.695), 0, 0, 0);
                    Poly11.AddVertexAt(++roomNumPoint, new Point2d(x - 38.698, y + 14.028), 0, 0, 0);
                    Poly11.AddVertexAt(++roomNumPoint, new Point2d(x - 38.854, y + 14.028), 0, 0, 0);
                    Poly11.AddVertexAt(++roomNumPoint, new Point2d(x - 39.242, y + 13.550), 0, 0, 0);
                    Poly11.AddVertexAt(++roomNumPoint, new Point2d(x - 39.626, y + 13.115), 0, 0, 0);
                    Poly11.AddVertexAt(++roomNumPoint, new Point2d(x - 39.624, y + 14.028), 0, 0, 0);
                    Poly11.Closed = true;
                    Poly11.LineWeight = 0;
                    Poly11.ConstantWidth = 0;
                    Poly11.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly11);
                    acTrans.AddNewlyCreatedDBObject(Poly11, true);

                    ObjectIdCollection acObjIdColl11 = new ObjectIdCollection();
                    acObjIdColl11.Add(Poly11.ObjectId);

                    //////////////////////////////////////////////////////////////////////////////////////////
                    // Я и пустота внутри
                    roomNumPoint = 0;
                    Polyline Poly12 = new Polyline();
                    Poly12.SetDatabaseDefaults();

                    Poly12.AddVertexAt(roomNumPoint, new Point2d(x - 37.290, y + 14.016), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 36.674, y + 14.018), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 36.674, y + 12.692), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 36.987, y + 12.692), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 36.987, y + 13.319), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.374, y + 12.690), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.742, y + 12.690), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.302, y + 13.351), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.367, y + 13.358), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.438, y + 13.382), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.496, y + 13.417), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.539, y + 13.457), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.573, y + 13.499), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.601, y + 13.550), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.621, y + 13.615), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.628, y + 13.683), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.614, y + 13.780), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.569, y + 13.873), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.515, y + 13.933), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.447, y + 13.980), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.375, y + 14.006), 0, 0, 0);
                    Poly12.Closed = true;
                    Poly12.LineWeight = 0;
                    Poly12.ConstantWidth = 0;
                    Poly12.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly12);
                    acTrans.AddNewlyCreatedDBObject(Poly12, true);

                    ObjectIdCollection acObjIdColl12 = new ObjectIdCollection();
                    acObjIdColl12.Add(Poly12.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly13 = new Polyline();
                    Poly13.SetDatabaseDefaults();
                    Poly13.AddVertexAt(roomNumPoint, new Point2d(x - 36.986, y + 13.918), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 36.986, y + 13.385), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 37.105, y + 13.385), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 37.174, y + 13.406), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 37.229, y + 13.446), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 37.268, y + 13.493), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 37.301, y + 13.583), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 37.307, y + 13.655), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 37.308, y + 13.709), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 37.293, y + 13.796), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 37.275, y + 13.839), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 37.253, y + 13.872), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 37.221, y + 13.891), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 37.185, y + 13.909), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 37.130, y + 13.918), 0, 0, 0);
                    Poly13.Closed = true;
                    Poly13.LineWeight = 0;
                    Poly13.ConstantWidth = 0;
                    Poly13.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly13);
                    acTrans.AddNewlyCreatedDBObject(Poly13, true);

                    ObjectIdCollection acObjIdColl13 = new ObjectIdCollection();
                    acObjIdColl13.Add(Poly13.ObjectId);

                    //////////////////////////////////////////////////////////////////////////////////////////
                    // S
                    roomNumPoint = 0;
                    Polyline Poly14 = new Polyline();
                    Poly14.SetDatabaseDefaults();
                    Poly14.AddVertexAt(roomNumPoint, new Point2d(x - 47.446, y + 2.263), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.512, y + 2.060), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.549, y + 2.060), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.555, y + 2.093), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.582, y + 2.136), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.641, y + 2.194), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.700, y + 2.224), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.736, y + 2.234), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.846, y + 2.234), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.915, y + 2.213), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.968, y + 2.180), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.998, y + 2.131), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.012, y + 2.086), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.012, y + 2.012), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.985, y + 1.956), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.932, y + 1.905), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.890, y + 1.875), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.829, y + 1.845), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.699, y + 1.776), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.546, y + 1.695), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.508, y + 1.661), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.449, y + 1.603), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.425, y + 1.559), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.397, y + 1.484), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.388, y + 1.445), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.388, y + 1.325), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.405, y + 1.252), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.442, y + 1.180), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.459, y + 1.151), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.536, y + 1.076), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.596, y + 1.037), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.647, y + 1.009), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.717, y + 0.985), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.783, y + 0.969), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.846, y + 0.959), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.022, y + 0.959), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.083, y + 0.967), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.157, y + 0.986), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.211, y + 1.004), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.248, y + 1.023), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.270, y + 1.034), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.207, y + 1.286), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.171, y + 1.286), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.139, y + 1.206), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.122, y + 1.173), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.087, y + 1.138), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.047, y + 1.109), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.007, y + 1.090), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.942, y + 1.070), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.824, y + 1.070), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.779, y + 1.082), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.737, y + 1.100), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.704, y + 1.126), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.670, y + 1.163), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.655, y + 1.190), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.645, y + 1.222), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.645, y + 1.308), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.656, y + 1.343), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.681, y + 1.385), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.733, y + 1.431), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.810, y + 1.477), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.924, y + 1.535), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.108, y + 1.626), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.131, y + 1.642), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.178, y + 1.689), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.224, y + 1.763), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.253, y + 1.832), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.261, y + 1.875), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.263, y + 1.994), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.252, y + 2.047), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.225, y + 2.116), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.167, y + 2.204), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.055, y + 2.292), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.923, y + 2.334), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.853, y + 2.345), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.700, y + 2.345), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.597, y + 2.327), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.515, y + 2.300), 0, 0, 0);
                    Poly14.Closed = true;
                    Poly14.LineWeight = 0;
                    Poly14.ConstantWidth = 0;
                    Poly14.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly14);
                    acTrans.AddNewlyCreatedDBObject(Poly14, true);

                    ObjectIdCollection acObjIdColl14 = new ObjectIdCollection();
                    acObjIdColl14.Add(Poly14.ObjectId);

                    //////////////////////////////////////////////////////////////////////////////////////////
                    // T
                    roomNumPoint = 0;
                    Polyline Poly15 = new Polyline();
                    Poly15.SetDatabaseDefaults();
                    Poly15.AddVertexAt(roomNumPoint, new Point2d(x - 45.692, y + 2.173), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 45.692, y + 2.162), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 45.572, y + 2.162), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 45.572, y + 2.153), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 45.493, y + 2.153), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 45.493, y + 2.168), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 45.504, y + 2.168), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 45.504, y + 2.279), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 45.493, y + 2.283), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 45.493, y + 2.306), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.550, y + 2.306), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.550, y + 2.279), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.544, y + 2.279), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.544, y + 2.170), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.553, y + 2.170), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.553, y + 2.144), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.515, y + 2.144), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.515, y + 2.154), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.436, y + 2.154), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.436, y + 2.164), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.334, y + 2.164), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.334, y + 2.173), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.168, y + 2.173), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.168, y + 0.983), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.088, y + 0.983), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.088, y + 0.995), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 45.950, y + 0.995), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 45.950, y + 0.983), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 45.860, y + 0.983), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 45.860, y + 2.173), 0, 0, 0);
                    Poly15.Closed = true;
                    Poly15.LineWeight = 0;
                    Poly15.ConstantWidth = 0;
                    Poly15.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly15);
                    acTrans.AddNewlyCreatedDBObject(Poly15, true);

                    ObjectIdCollection acObjIdColl15 = new ObjectIdCollection();
                    acObjIdColl15.Add(Poly15.ObjectId);

                    //////////////////////////////////////////////////////////////////////////////////////////
                    // U
                    roomNumPoint = 0;
                    Polyline Poly16 = new Polyline();
                    Poly16.SetDatabaseDefaults();
                    Poly16.AddVertexAt(roomNumPoint, new Point2d(x - 44.466, y + 2.307), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.158, y + 2.307), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.158, y + 1.464), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.126, y + 1.310), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.094, y + 1.238), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.048, y + 1.180), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.971, y + 1.125), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.858, y + 1.087), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.713, y + 1.087), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.643, y + 1.099), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.555, y + 1.134), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.475, y + 1.216), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.436, y + 1.278), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.409, y + 1.345), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.390, y + 1.440), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.383, y + 1.519), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.383, y + 2.307), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.216, y + 2.307), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.216, y + 1.470), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.236, y + 1.336), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.273, y + 1.233), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.320, y + 1.162), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.384, y + 1.094), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.469, y + 1.037), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.522, y + 1.010), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.595, y + 0.986), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.680, y + 0.969), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.741, y + 0.959), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.936, y + 0.959), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.101, y + 0.987), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.200, y + 1.022), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.281, y + 1.067), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.348, y + 1.126), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.386, y + 1.175), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.431, y + 1.262), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.451, y + 1.330), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.458, y + 1.366), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.466, y + 1.435), 0, 0, 0);
                    Poly16.Closed = true;
                    Poly16.LineWeight = 0;
                    Poly16.ConstantWidth = 0;
                    Poly16.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly16);
                    acTrans.AddNewlyCreatedDBObject(Poly16, true);

                    ObjectIdCollection acObjIdColl16 = new ObjectIdCollection();
                    acObjIdColl16.Add(Poly16.ObjectId);

                    //////////////////////////////////////////////////////////////////////////////////////////
                    // D
                    roomNumPoint = 0;
                    Polyline Poly17 = new Polyline();
                    Poly17.SetDatabaseDefaults();
                    Poly17.AddVertexAt(roomNumPoint, new Point2d(x - 42.052, y + 2.298), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 42.052, y + 0.983), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 41.372, y + 0.983), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 41.203, y + 1.017), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 41.066, y + 1.073), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 40.957, y + 1.145), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 40.887, y + 1.207), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 40.806, y + 1.316), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 40.755, y + 1.421), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 40.723, y + 1.538), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 40.715, y + 1.603), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 40.715, y + 1.766), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 40.743, y + 1.915), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 40.797, y + 2.033), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 40.834, y + 2.088), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 40.944, y + 2.188), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 41.028, y + 2.246), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 41.120, y + 2.282), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 41.192, y + 2.299), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 41.256, y + 2.308), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 42.052, y + 2.308), 0, 0, 0);
                    Poly17.Closed = true;
                    Poly17.LineWeight = 0;
                    Poly17.ConstantWidth = 0;
                    Poly17.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly17);
                    acTrans.AddNewlyCreatedDBObject(Poly17, true);

                    ObjectIdCollection acObjIdColl17 = new ObjectIdCollection();
                    acObjIdColl17.Add(Poly17.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly18 = new Polyline();
                    Poly18.SetDatabaseDefaults();
                    Poly18.AddVertexAt(roomNumPoint, new Point2d(x - 41.749, y + 2.204), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.749, y + 1.089), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.479, y + 1.089), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.440, y + 1.096), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.365, y + 1.118), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.273, y + 1.162), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.225, y + 1.196), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.163, y + 1.267), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.127, y + 1.334), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.095, y + 1.410), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.075, y + 1.476), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.065, y + 1.576), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.060, y + 1.677), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.065, y + 1.780), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.085, y + 1.898), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.110, y + 1.963), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.149, y + 2.035), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.235, y + 2.119), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.290, y + 2.151), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.356, y + 2.178), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.440, y + 2.196), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.454, y + 2.199), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.495, y + 2.204), 0, 0, 0);
                    Poly18.Closed = true;
                    Poly18.LineWeight = 0;
                    Poly18.ConstantWidth = 0;
                    Poly18.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly18);
                    acTrans.AddNewlyCreatedDBObject(Poly18, true);

                    ObjectIdCollection acObjIdColl18 = new ObjectIdCollection();
                    acObjIdColl18.Add(Poly18.ObjectId);

                    /////////////////////////////////////////////////////////////////////////////////////////
                    // I
                    roomNumPoint = 0;
                    Polyline Poly19 = new Polyline();
                    Poly19.SetDatabaseDefaults();
                    Poly19.AddVertexAt(roomNumPoint, new Point2d(x - 39.625, y + 2.298), 0, 0, 0);
                    Poly19.AddVertexAt(++roomNumPoint, new Point2d(x - 39.315, y + 2.298), 0, 0, 0);
                    Poly19.AddVertexAt(++roomNumPoint, new Point2d(x - 39.315, y + 0.983), 0, 0, 0);
                    Poly19.AddVertexAt(++roomNumPoint, new Point2d(x - 39.625, y + 0.983), 0, 0, 0);
                    Poly19.Closed = true;
                    Poly19.LineWeight = 0;
                    Poly19.ConstantWidth = 0;
                    Poly19.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly19);
                    acTrans.AddNewlyCreatedDBObject(Poly19, true);

                    ObjectIdCollection acObjIdColl19 = new ObjectIdCollection();
                    acObjIdColl19.Add(Poly19.ObjectId);

                    /////////////////////////////////////////////////////////////////////////////////////////
                    // O
                    roomNumPoint = 0;
                    Polyline Poly20 = new Polyline();
                    Poly20.SetDatabaseDefaults();
                    Poly20.AddVertexAt(roomNumPoint, new Point2d(x - 37.550, y + 2.346), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.340, y + 2.346), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.174, y + 2.321), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.065, y + 2.283), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 36.973, y + 2.236), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 36.919, y + 2.199), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 36.808, y + 2.086), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 36.759, y + 2.021), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 36.709, y + 1.909), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 36.680, y + 1.802), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 36.674, y + 1.738), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 36.674, y + 1.572), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 36.680, y + 1.516), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 36.718, y + 1.395), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 36.767, y + 1.299), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 36.818, y + 1.235), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 36.902, y + 1.148), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.002, y + 1.075), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.095, y + 1.030), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.167, y + 1.003), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.267, y + 0.975), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.386, y + 0.957), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.598, y + 0.957), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.772, y + 0.986), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.877, y + 1.023), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 38.002, y + 1.088), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 38.065, y + 1.143), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 38.120, y + 1.198), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 38.195, y + 1.305), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 38.230, y + 1.389), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 38.256, y + 1.490), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 38.266, y + 1.554), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 38.266, y + 1.730), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 38.249, y + 1.835), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 38.202, y + 1.958), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 38.119, y + 2.077), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 38.045, y + 2.151), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.943, y + 2.225), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.835, y + 2.281), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.711, y + 2.319), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.621, y + 2.336), 0, 0, 0);
                    Poly20.Closed = true;
                    Poly20.LineWeight = 0;
                    Poly20.ConstantWidth = 0;
                    Poly20.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly20);
                    acTrans.AddNewlyCreatedDBObject(Poly20, true);

                    ObjectIdCollection acObjIdColl20 = new ObjectIdCollection();
                    acObjIdColl20.Add(Poly20.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly21 = new Polyline();
                    Poly21.SetDatabaseDefaults();
                    Poly21.AddVertexAt(roomNumPoint, new Point2d(x - 37.524, y + 2.230), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.391, y + 2.230), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.306, y + 2.208), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.221, y + 2.169), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.127, y + 2.072), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.061, y + 1.973), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.035, y + 1.864), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.020, y + 1.774), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.020, y + 1.591), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.041, y + 1.463), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.077, y + 1.357), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.120, y + 1.268), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.177, y + 1.187), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.261, y + 1.123), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.324, y + 1.089), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.417, y + 1.068), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.556, y + 1.068), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.652, y + 1.093), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.733, y + 1.140), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.808, y + 1.215), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.865, y + 1.307), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.898, y + 1.392), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.917, y + 1.505), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.917, y + 1.692), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.906, y + 1.799), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.889, y + 1.884), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.867, y + 1.958), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.831, y + 2.032), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.801, y + 2.074), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.716, y + 2.160), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.640, y + 2.201), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.567, y + 2.227), 0, 0, 0);
                    Poly21.Closed = true;
                    Poly21.LineWeight = 0;
                    Poly21.ConstantWidth = 0;
                    Poly21.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly21);
                    acTrans.AddNewlyCreatedDBObject(Poly21, true);

                    ObjectIdCollection acObjIdColl21 = new ObjectIdCollection();
                    acObjIdColl21.Add(Poly21.ObjectId);

                    /////////////////////////////////////////////////////////////////////////////////////////
                    // Надпись ООО
                    roomNumPoint = 0;
                    Polyline Poly22 = new Polyline();
                    Poly22.SetDatabaseDefaults();
                    Poly22.AddVertexAt(roomNumPoint, new Point2d(x - 34.820, y + 9.270), 0, 0, 0);
                    Poly22.AddVertexAt(++roomNumPoint, new Point2d(x - 34.750, y + 9.030), 0, 0, 0);
                    Poly22.AddVertexAt(++roomNumPoint, new Point2d(x - 34.640, y + 8.870), 0, 0, 0);
                    Poly22.AddVertexAt(++roomNumPoint, new Point2d(x - 34.510, y + 8.760), 0, 0, 0);
                    Poly22.AddVertexAt(++roomNumPoint, new Point2d(x - 34.360, y + 8.680), 0, 0, 0);
                    Poly22.AddVertexAt(++roomNumPoint, new Point2d(x - 34.190, y + 8.640), 0, 0, 0);
                    Poly22.AddVertexAt(++roomNumPoint, new Point2d(x - 34.010, y + 8.640), 0, 0, 0);
                    Poly22.AddVertexAt(++roomNumPoint, new Point2d(x - 33.850, y + 8.680), 0, 0, 0);
                    Poly22.AddVertexAt(++roomNumPoint, new Point2d(x - 33.700, y + 8.760), 0, 0, 0);
                    Poly22.AddVertexAt(++roomNumPoint, new Point2d(x - 33.580, y + 8.870), 0, 0, 0);
                    Poly22.AddVertexAt(++roomNumPoint, new Point2d(x - 33.470, y + 9.030), 0, 0, 0);
                    Poly22.AddVertexAt(++roomNumPoint, new Point2d(x - 33.390, y + 9.230), 0, 0, 0);
                    Poly22.AddVertexAt(++roomNumPoint, new Point2d(x - 33.400, y + 9.520), 0, 0, 0);
                    Poly22.AddVertexAt(++roomNumPoint, new Point2d(x - 33.470, y + 9.710), 0, 0, 0);
                    Poly22.AddVertexAt(++roomNumPoint, new Point2d(x - 33.600, y + 9.870), 0, 0, 0);
                    Poly22.AddVertexAt(++roomNumPoint, new Point2d(x - 33.760, y + 10.000), 0, 0, 0);
                    Poly22.AddVertexAt(++roomNumPoint, new Point2d(x - 33.910, y + 10.030), 0, 0, 0);
                    Poly22.AddVertexAt(++roomNumPoint, new Point2d(x - 34.300, y + 10.030), 0, 0, 0);
                    Poly22.AddVertexAt(++roomNumPoint, new Point2d(x - 34.460, y + 10.000), 0, 0, 0);
                    Poly22.AddVertexAt(++roomNumPoint, new Point2d(x - 34.630, y + 9.870), 0, 0, 0);
                    Poly22.AddVertexAt(++roomNumPoint, new Point2d(x - 34.750, y + 9.710), 0, 0, 0);
                    Poly22.AddVertexAt(++roomNumPoint, new Point2d(x - 34.820, y + 9.490), 0, 0, 0);
                    Poly22.Closed = true;
                    Poly22.LineWeight = 0;
                    Poly22.ConstantWidth = 0;
                    Poly22.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly22);
                    acTrans.AddNewlyCreatedDBObject(Poly22, true);

                    ObjectIdCollection acObjIdColl22 = new ObjectIdCollection();
                    acObjIdColl22.Add(Poly22.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly23 = new Polyline();
                    Poly23.SetDatabaseDefaults();
                    Poly23.AddVertexAt(roomNumPoint, new Point2d(x - 34.640, y + 9.440), 0, 0, 0);
                    Poly23.AddVertexAt(++roomNumPoint, new Point2d(x - 34.600, y + 9.660), 0, 0, 0);
                    Poly23.AddVertexAt(++roomNumPoint, new Point2d(x - 34.520, y + 9.820), 0, 0, 0);
                    Poly23.AddVertexAt(++roomNumPoint, new Point2d(x - 34.440, y + 9.920), 0, 0, 0);
                    Poly23.AddVertexAt(++roomNumPoint, new Point2d(x - 34.330, y + 9.970), 0, 0, 0);
                    Poly23.AddVertexAt(++roomNumPoint, new Point2d(x - 34.110, y + 9.980), 0, 0, 0);
                    Poly23.AddVertexAt(++roomNumPoint, new Point2d(x - 33.880, y + 9.970), 0, 0, 0);
                    Poly23.AddVertexAt(++roomNumPoint, new Point2d(x - 33.770, y + 9.920), 0, 0, 0);
                    Poly23.AddVertexAt(++roomNumPoint, new Point2d(x - 33.680, y + 9.820), 0, 0, 0);
                    Poly23.AddVertexAt(++roomNumPoint, new Point2d(x - 33.600, y + 9.660), 0, 0, 0);
                    Poly23.AddVertexAt(++roomNumPoint, new Point2d(x - 33.560, y + 9.440), 0, 0, 0);
                    Poly23.AddVertexAt(++roomNumPoint, new Point2d(x - 33.560, y + 9.160), 0, 0, 0);
                    Poly23.AddVertexAt(++roomNumPoint, new Point2d(x - 33.630, y + 8.970), 0, 0, 0);
                    Poly23.AddVertexAt(++roomNumPoint, new Point2d(x - 33.760, y + 8.820), 0, 0, 0);
                    Poly23.AddVertexAt(++roomNumPoint, new Point2d(x - 33.920, y + 8.750), 0, 0, 0);
                    Poly23.AddVertexAt(++roomNumPoint, new Point2d(x - 34.070, y + 8.720), 0, 0, 0);
                    Poly23.AddVertexAt(++roomNumPoint, new Point2d(x - 34.140, y + 8.720), 0, 0, 0);
                    Poly23.AddVertexAt(++roomNumPoint, new Point2d(x - 34.300, y + 8.750), 0, 0, 0);
                    Poly23.AddVertexAt(++roomNumPoint, new Point2d(x - 34.450, y + 8.820), 0, 0, 0);
                    Poly23.AddVertexAt(++roomNumPoint, new Point2d(x - 34.570, y + 8.970), 0, 0, 0);
                    Poly23.AddVertexAt(++roomNumPoint, new Point2d(x - 34.640, y + 9.150), 0, 0, 0);
                    Poly23.Closed = true;
                    Poly23.LineWeight = 0;
                    Poly23.ConstantWidth = 0;
                    Poly23.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly23);
                    acTrans.AddNewlyCreatedDBObject(Poly23, true);

                    ObjectIdCollection acObjIdColl23 = new ObjectIdCollection();
                    acObjIdColl23.Add(Poly23.ObjectId);
                    /////////////////////////////////////////////////////////////////////////////////////////
                    roomNumPoint = 0;
                    Polyline Poly24 = new Polyline();
                    Poly24.SetDatabaseDefaults();
                    Poly24.AddVertexAt(roomNumPoint, new Point2d(x - 34.820 + 1.900, y + 9.270), 0, 0, 0);
                    Poly24.AddVertexAt(++roomNumPoint, new Point2d(x - 34.750 + 1.900, y + 9.030), 0, 0, 0);
                    Poly24.AddVertexAt(++roomNumPoint, new Point2d(x - 34.640 + 1.900, y + 8.870), 0, 0, 0);
                    Poly24.AddVertexAt(++roomNumPoint, new Point2d(x - 34.510 + 1.900, y + 8.760), 0, 0, 0);
                    Poly24.AddVertexAt(++roomNumPoint, new Point2d(x - 34.360 + 1.900, y + 8.680), 0, 0, 0);
                    Poly24.AddVertexAt(++roomNumPoint, new Point2d(x - 34.190 + 1.900, y + 8.640), 0, 0, 0);
                    Poly24.AddVertexAt(++roomNumPoint, new Point2d(x - 34.010 + 1.900, y + 8.640), 0, 0, 0);
                    Poly24.AddVertexAt(++roomNumPoint, new Point2d(x - 33.850 + 1.900, y + 8.680), 0, 0, 0);
                    Poly24.AddVertexAt(++roomNumPoint, new Point2d(x - 33.700 + 1.900, y + 8.760), 0, 0, 0);
                    Poly24.AddVertexAt(++roomNumPoint, new Point2d(x - 33.580 + 1.900, y + 8.870), 0, 0, 0);
                    Poly24.AddVertexAt(++roomNumPoint, new Point2d(x - 33.470 + 1.900, y + 9.030), 0, 0, 0);
                    Poly24.AddVertexAt(++roomNumPoint, new Point2d(x - 33.390 + 1.900, y + 9.230), 0, 0, 0);
                    Poly24.AddVertexAt(++roomNumPoint, new Point2d(x - 33.400 + 1.900, y + 9.520), 0, 0, 0);
                    Poly24.AddVertexAt(++roomNumPoint, new Point2d(x - 33.470 + 1.900, y + 9.710), 0, 0, 0);
                    Poly24.AddVertexAt(++roomNumPoint, new Point2d(x - 33.600 + 1.900, y + 9.870), 0, 0, 0);
                    Poly24.AddVertexAt(++roomNumPoint, new Point2d(x - 33.760 + 1.900, y + 10.000), 0, 0, 0);
                    Poly24.AddVertexAt(++roomNumPoint, new Point2d(x - 33.910 + 1.900, y + 10.030), 0, 0, 0);
                    Poly24.AddVertexAt(++roomNumPoint, new Point2d(x - 34.300 + 1.900, y + 10.030), 0, 0, 0);
                    Poly24.AddVertexAt(++roomNumPoint, new Point2d(x - 34.460 + 1.900, y + 10.000), 0, 0, 0);
                    Poly24.AddVertexAt(++roomNumPoint, new Point2d(x - 34.630 + 1.900, y + 9.870), 0, 0, 0);
                    Poly24.AddVertexAt(++roomNumPoint, new Point2d(x - 34.750 + 1.900, y + 9.710), 0, 0, 0);
                    Poly24.AddVertexAt(++roomNumPoint, new Point2d(x - 34.820 + 1.900, y + 9.490), 0, 0, 0);
                    Poly24.Closed = true;
                    Poly24.LineWeight = 0;
                    Poly24.ConstantWidth = 0;
                    Poly24.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly24);
                    acTrans.AddNewlyCreatedDBObject(Poly24, true);

                    ObjectIdCollection acObjIdColl24 = new ObjectIdCollection();
                    acObjIdColl24.Add(Poly24.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly25 = new Polyline();
                    Poly25.SetDatabaseDefaults();
                    Poly25.AddVertexAt(roomNumPoint, new Point2d(x - 34.640 + 1.900, y + 9.440), 0, 0, 0);
                    Poly25.AddVertexAt(++roomNumPoint, new Point2d(x - 34.600 + 1.900, y + 9.660), 0, 0, 0);
                    Poly25.AddVertexAt(++roomNumPoint, new Point2d(x - 34.520 + 1.900, y + 9.820), 0, 0, 0);
                    Poly25.AddVertexAt(++roomNumPoint, new Point2d(x - 34.440 + 1.900, y + 9.920), 0, 0, 0);
                    Poly25.AddVertexAt(++roomNumPoint, new Point2d(x - 34.330 + 1.900, y + 9.970), 0, 0, 0);
                    Poly25.AddVertexAt(++roomNumPoint, new Point2d(x - 34.110 + 1.900, y + 9.980), 0, 0, 0);
                    Poly25.AddVertexAt(++roomNumPoint, new Point2d(x - 33.880 + 1.900, y + 9.970), 0, 0, 0);
                    Poly25.AddVertexAt(++roomNumPoint, new Point2d(x - 33.770 + 1.900, y + 9.920), 0, 0, 0);
                    Poly25.AddVertexAt(++roomNumPoint, new Point2d(x - 33.680 + 1.900, y + 9.820), 0, 0, 0);
                    Poly25.AddVertexAt(++roomNumPoint, new Point2d(x - 33.600 + 1.900, y + 9.660), 0, 0, 0);
                    Poly25.AddVertexAt(++roomNumPoint, new Point2d(x - 33.560 + 1.900, y + 9.440), 0, 0, 0);
                    Poly25.AddVertexAt(++roomNumPoint, new Point2d(x - 33.560 + 1.900, y + 9.160), 0, 0, 0);
                    Poly25.AddVertexAt(++roomNumPoint, new Point2d(x - 33.630 + 1.900, y + 8.970), 0, 0, 0);
                    Poly25.AddVertexAt(++roomNumPoint, new Point2d(x - 33.760 + 1.900, y + 8.820), 0, 0, 0);
                    Poly25.AddVertexAt(++roomNumPoint, new Point2d(x - 33.920 + 1.900, y + 8.750), 0, 0, 0);
                    Poly25.AddVertexAt(++roomNumPoint, new Point2d(x - 34.070 + 1.900, y + 8.720), 0, 0, 0);
                    Poly25.AddVertexAt(++roomNumPoint, new Point2d(x - 34.140 + 1.900, y + 8.720), 0, 0, 0);
                    Poly25.AddVertexAt(++roomNumPoint, new Point2d(x - 34.300 + 1.900, y + 8.750), 0, 0, 0);
                    Poly25.AddVertexAt(++roomNumPoint, new Point2d(x - 34.450 + 1.900, y + 8.820), 0, 0, 0);
                    Poly25.AddVertexAt(++roomNumPoint, new Point2d(x - 34.570 + 1.900, y + 8.970), 0, 0, 0);
                    Poly25.AddVertexAt(++roomNumPoint, new Point2d(x - 34.640 + 1.900, y + 9.150), 0, 0, 0);
                    Poly25.Closed = true;
                    Poly25.LineWeight = 0;
                    Poly25.ConstantWidth = 0;
                    Poly25.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly25);
                    acTrans.AddNewlyCreatedDBObject(Poly25, true);

                    ObjectIdCollection acObjIdColl25 = new ObjectIdCollection();
                    acObjIdColl25.Add(Poly25.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    roomNumPoint = 0;
                    Polyline Poly26 = new Polyline();
                    Poly26.SetDatabaseDefaults();
                    Poly26.AddVertexAt(roomNumPoint, new Point2d(x - 34.820 + 3.800, y + 9.270), 0, 0, 0);
                    Poly26.AddVertexAt(++roomNumPoint, new Point2d(x - 34.750 + 3.800, y + 9.030), 0, 0, 0);
                    Poly26.AddVertexAt(++roomNumPoint, new Point2d(x - 34.640 + 3.800, y + 8.870), 0, 0, 0);
                    Poly26.AddVertexAt(++roomNumPoint, new Point2d(x - 34.510 + 3.800, y + 8.760), 0, 0, 0);
                    Poly26.AddVertexAt(++roomNumPoint, new Point2d(x - 34.360 + 3.800, y + 8.680), 0, 0, 0);
                    Poly26.AddVertexAt(++roomNumPoint, new Point2d(x - 34.190 + 3.800, y + 8.640), 0, 0, 0);
                    Poly26.AddVertexAt(++roomNumPoint, new Point2d(x - 34.010 + 3.800, y + 8.640), 0, 0, 0);
                    Poly26.AddVertexAt(++roomNumPoint, new Point2d(x - 33.850 + 3.800, y + 8.680), 0, 0, 0);
                    Poly26.AddVertexAt(++roomNumPoint, new Point2d(x - 33.700 + 3.800, y + 8.760), 0, 0, 0);
                    Poly26.AddVertexAt(++roomNumPoint, new Point2d(x - 33.580 + 3.800, y + 8.870), 0, 0, 0);
                    Poly26.AddVertexAt(++roomNumPoint, new Point2d(x - 33.470 + 3.800, y + 9.030), 0, 0, 0);
                    Poly26.AddVertexAt(++roomNumPoint, new Point2d(x - 33.390 + 3.800, y + 9.230), 0, 0, 0);
                    Poly26.AddVertexAt(++roomNumPoint, new Point2d(x - 33.400 + 3.800, y + 9.520), 0, 0, 0);
                    Poly26.AddVertexAt(++roomNumPoint, new Point2d(x - 33.470 + 3.800, y + 9.710), 0, 0, 0);
                    Poly26.AddVertexAt(++roomNumPoint, new Point2d(x - 33.600 + 3.800, y + 9.870), 0, 0, 0);
                    Poly26.AddVertexAt(++roomNumPoint, new Point2d(x - 33.760 + 3.800, y + 10.000), 0, 0, 0);
                    Poly26.AddVertexAt(++roomNumPoint, new Point2d(x - 33.910 + 3.800, y + 10.030), 0, 0, 0);
                    Poly26.AddVertexAt(++roomNumPoint, new Point2d(x - 34.300 + 3.800, y + 10.030), 0, 0, 0);
                    Poly26.AddVertexAt(++roomNumPoint, new Point2d(x - 34.460 + 3.800, y + 10.000), 0, 0, 0);
                    Poly26.AddVertexAt(++roomNumPoint, new Point2d(x - 34.630 + 3.800, y + 9.870), 0, 0, 0);
                    Poly26.AddVertexAt(++roomNumPoint, new Point2d(x - 34.750 + 3.800, y + 9.710), 0, 0, 0);
                    Poly26.AddVertexAt(++roomNumPoint, new Point2d(x - 34.820 + 3.800, y + 9.490), 0, 0, 0);
                    Poly26.Closed = true;
                    Poly26.LineWeight = 0;
                    Poly26.ConstantWidth = 0;
                    Poly26.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly26);
                    acTrans.AddNewlyCreatedDBObject(Poly26, true);

                    ObjectIdCollection acObjIdColl26 = new ObjectIdCollection();
                    acObjIdColl26.Add(Poly26.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly27 = new Polyline();
                    Poly27.SetDatabaseDefaults();
                    Poly27.AddVertexAt(roomNumPoint, new Point2d(x - 34.640 + 3.800, y + 9.440), 0, 0, 0);
                    Poly27.AddVertexAt(++roomNumPoint, new Point2d(x - 34.600 + 3.800, y + 9.660), 0, 0, 0);
                    Poly27.AddVertexAt(++roomNumPoint, new Point2d(x - 34.520 + 3.800, y + 9.820), 0, 0, 0);
                    Poly27.AddVertexAt(++roomNumPoint, new Point2d(x - 34.440 + 3.800, y + 9.920), 0, 0, 0);
                    Poly27.AddVertexAt(++roomNumPoint, new Point2d(x - 34.330 + 3.800, y + 9.970), 0, 0, 0);
                    Poly27.AddVertexAt(++roomNumPoint, new Point2d(x - 34.110 + 3.800, y + 9.980), 0, 0, 0);
                    Poly27.AddVertexAt(++roomNumPoint, new Point2d(x - 33.880 + 3.800, y + 9.970), 0, 0, 0);
                    Poly27.AddVertexAt(++roomNumPoint, new Point2d(x - 33.770 + 3.800, y + 9.920), 0, 0, 0);
                    Poly27.AddVertexAt(++roomNumPoint, new Point2d(x - 33.680 + 3.800, y + 9.820), 0, 0, 0);
                    Poly27.AddVertexAt(++roomNumPoint, new Point2d(x - 33.600 + 3.800, y + 9.660), 0, 0, 0);
                    Poly27.AddVertexAt(++roomNumPoint, new Point2d(x - 33.560 + 3.800, y + 9.440), 0, 0, 0);
                    Poly27.AddVertexAt(++roomNumPoint, new Point2d(x - 33.560 + 3.800, y + 9.160), 0, 0, 0);
                    Poly27.AddVertexAt(++roomNumPoint, new Point2d(x - 33.630 + 3.800, y + 8.970), 0, 0, 0);
                    Poly27.AddVertexAt(++roomNumPoint, new Point2d(x - 33.760 + 3.800, y + 8.820), 0, 0, 0);
                    Poly27.AddVertexAt(++roomNumPoint, new Point2d(x - 33.920 + 3.800, y + 8.750), 0, 0, 0);
                    Poly27.AddVertexAt(++roomNumPoint, new Point2d(x - 34.070 + 3.800, y + 8.720), 0, 0, 0);
                    Poly27.AddVertexAt(++roomNumPoint, new Point2d(x - 34.140 + 3.800, y + 8.720), 0, 0, 0);
                    Poly27.AddVertexAt(++roomNumPoint, new Point2d(x - 34.300 + 3.800, y + 8.750), 0, 0, 0);
                    Poly27.AddVertexAt(++roomNumPoint, new Point2d(x - 34.450 + 3.800, y + 8.820), 0, 0, 0);
                    Poly27.AddVertexAt(++roomNumPoint, new Point2d(x - 34.570 + 3.800, y + 8.970), 0, 0, 0);
                    Poly27.AddVertexAt(++roomNumPoint, new Point2d(x - 34.640 + 3.800, y + 9.150), 0, 0, 0);
                    Poly27.Closed = true;
                    Poly27.LineWeight = 0;
                    Poly27.ConstantWidth = 0;
                    Poly27.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly27);
                    acTrans.AddNewlyCreatedDBObject(Poly27, true);

                    ObjectIdCollection acObjIdColl27 = new ObjectIdCollection();
                    acObjIdColl27.Add(Poly27.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    // "
                    roomNumPoint = 0;
                    Polyline Poly28 = new Polyline();
                    Poly28.SetDatabaseDefaults();
                    Poly28.AddVertexAt(roomNumPoint, new Point2d(x - 28.290, y + 9.500), 0, 0, 0);
                    Poly28.AddVertexAt(++roomNumPoint, new Point2d(x - 28.170, y + 9.530), 0, 0, 0);
                    Poly28.AddVertexAt(++roomNumPoint, new Point2d(x - 28.010, y + 10.030), 0, 0, 0);
                    Poly28.AddVertexAt(++roomNumPoint, new Point2d(x - 28.100, y + 9.970), 0, 0, 0);
                    Poly28.AddVertexAt(++roomNumPoint, new Point2d(x - 28.200, y + 9.810), 0, 0, 0);
                    Poly28.AddVertexAt(++roomNumPoint, new Point2d(x - 28.260, y + 9.660), 0, 0, 0);
                    Poly28.Closed = true;
                    Poly28.LineWeight = 0;
                    Poly28.ConstantWidth = 0;
                    Poly28.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly28);
                    acTrans.AddNewlyCreatedDBObject(Poly28, true);

                    ObjectIdCollection acObjIdColl28 = new ObjectIdCollection();
                    acObjIdColl28.Add(Poly28.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly29 = new Polyline();
                    Poly29.SetDatabaseDefaults();
                    Poly29.AddVertexAt(roomNumPoint, new Point2d(x - 28.290 + 0.350, y + 9.500), 0, 0, 0);
                    Poly29.AddVertexAt(++roomNumPoint, new Point2d(x - 28.170 + 0.350, y + 9.530), 0, 0, 0);
                    Poly29.AddVertexAt(++roomNumPoint, new Point2d(x - 28.010 + 0.350, y + 10.030), 0, 0, 0);
                    Poly29.AddVertexAt(++roomNumPoint, new Point2d(x - 28.100 + 0.350, y + 9.970), 0, 0, 0);
                    Poly29.AddVertexAt(++roomNumPoint, new Point2d(x - 28.200 + 0.350, y + 9.810), 0, 0, 0);
                    Poly29.AddVertexAt(++roomNumPoint, new Point2d(x - 28.260 + 0.350, y + 9.660), 0, 0, 0);
                    Poly29.Closed = true;
                    Poly29.LineWeight = 0;
                    Poly29.ConstantWidth = 0;
                    Poly29.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly29);
                    acTrans.AddNewlyCreatedDBObject(Poly29, true);

                    ObjectIdCollection acObjIdColl29 = new ObjectIdCollection();
                    acObjIdColl29.Add(Poly29.ObjectId);
                    /////////////////////////////////////////////////////////////////////////////////////////
                    // А
                    roomNumPoint = 0;
                    Polyline Poly30 = new Polyline();
                    Poly30.SetDatabaseDefaults();
                    Poly30.AddVertexAt(roomNumPoint, new Point2d(x - 27.510, y + 8.650), 0, 0, 0);
                    Poly30.AddVertexAt(++roomNumPoint, new Point2d(x - 27.410, y + 8.650), 0, 0, 0);
                    Poly30.AddVertexAt(++roomNumPoint, new Point2d(x - 27.350, y + 8.750), 0, 0, 0);
                    Poly30.AddVertexAt(++roomNumPoint, new Point2d(x - 27.290, y + 8.810), 0, 0, 0);
                    Poly30.AddVertexAt(++roomNumPoint, new Point2d(x - 27.250, y + 8.870), 0, 0, 0);
                    Poly30.AddVertexAt(++roomNumPoint, new Point2d(x - 27.190, y + 9.060), 0, 0, 0);
                    Poly30.AddVertexAt(++roomNumPoint, new Point2d(x - 27.160, y + 9.190), 0, 0, 0);
                    Poly30.AddVertexAt(++roomNumPoint, new Point2d(x - 26.630, y + 9.190), 0, 0, 0);
                    Poly30.AddVertexAt(++roomNumPoint, new Point2d(x - 26.440, y + 8.650), 0, 0, 0);
                    Poly30.AddVertexAt(++roomNumPoint, new Point2d(x - 26.220, y + 8.650), 0, 0, 0);
                    Poly30.AddVertexAt(++roomNumPoint, new Point2d(x - 26.310, y + 8.970), 0, 0, 0);
                    Poly30.AddVertexAt(++roomNumPoint, new Point2d(x - 26.530, y + 9.410), 0, 0, 0);
                    Poly30.AddVertexAt(++roomNumPoint, new Point2d(x - 26.720, y + 9.850), 0, 0, 0);
                    Poly30.AddVertexAt(++roomNumPoint, new Point2d(x - 26.880, y + 10.030), 0, 0, 0);
                    Poly30.Closed = true;
                    Poly30.LineWeight = 0;
                    Poly30.ConstantWidth = 0;
                    Poly30.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly30);
                    acTrans.AddNewlyCreatedDBObject(Poly30, true);

                    ObjectIdCollection acObjIdColl30 = new ObjectIdCollection();
                    acObjIdColl30.Add(Poly30.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly31 = new Polyline();
                    Poly31.SetDatabaseDefaults();
                    Poly31.AddVertexAt(roomNumPoint, new Point2d(x - 26.910, y + 9.780), 0, 0, 0);
                    Poly31.AddVertexAt(++roomNumPoint, new Point2d(x - 26.690, y + 9.310), 0, 0, 0);
                    Poly31.AddVertexAt(++roomNumPoint, new Point2d(x - 27.100, y + 9.310), 0, 0, 0);
                    Poly31.Closed = true;
                    Poly31.LineWeight = 0;
                    Poly31.ConstantWidth = 0;
                    Poly31.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly31);
                    acTrans.AddNewlyCreatedDBObject(Poly31, true);

                    ObjectIdCollection acObjIdColl31 = new ObjectIdCollection();
                    acObjIdColl31.Add(Poly31.ObjectId);
                    /////////////////////////////////////////////////////////////////////////////////////////
                    // аРхитектуРное бюРо
                    roomNumPoint = 0;
                    Polyline Poly32 = new Polyline();
                    Poly32.SetDatabaseDefaults();
                    Poly32.AddVertexAt(roomNumPoint, new Point2d(x - 25.780, y + 8.650), 0, 0, 0);
                    Poly32.AddVertexAt(++roomNumPoint, new Point2d(x - 25.620, y + 8.650), 0, 0, 0);
                    Poly32.AddVertexAt(++roomNumPoint, new Point2d(x - 25.620, y + 9.280), 0, 0, 0);
                    Poly32.AddVertexAt(++roomNumPoint, new Point2d(x - 25.430, y + 9.280), 0, 0, 0);
                    Poly32.AddVertexAt(++roomNumPoint, new Point2d(x - 25.240, y + 9.310), 0, 0, 0);
                    Poly32.AddVertexAt(++roomNumPoint, new Point2d(x - 25.130, y + 9.370), 0, 0, 0);
                    Poly32.AddVertexAt(++roomNumPoint, new Point2d(x - 25.030, y + 9.440), 0, 0, 0);
                    Poly32.AddVertexAt(++roomNumPoint, new Point2d(x - 24.960, y + 9.530), 0, 0, 0);
                    Poly32.AddVertexAt(++roomNumPoint, new Point2d(x - 24.930, y + 9.590), 0, 0, 0);
                    Poly32.AddVertexAt(++roomNumPoint, new Point2d(x - 24.930, y + 9.850), 0, 0, 0);
                    Poly32.AddVertexAt(++roomNumPoint, new Point2d(x - 24.990, y + 9.940), 0, 0, 0);
                    Poly32.AddVertexAt(++roomNumPoint, new Point2d(x - 25.060, y + 9.970), 0, 0, 0);
                    Poly32.AddVertexAt(++roomNumPoint, new Point2d(x - 25.120, y + 10.000), 0, 0, 0);
                    Poly32.AddVertexAt(++roomNumPoint, new Point2d(x - 25.310, y + 10.030), 0, 0, 0);
                    Poly32.AddVertexAt(++roomNumPoint, new Point2d(x - 25.530, y + 10.030), 0, 0, 0);
                    Poly32.AddVertexAt(++roomNumPoint, new Point2d(x - 25.780, y + 10.030), 0, 0, 0);
                    Poly32.Closed = true;
                    Poly32.LineWeight = 0;
                    Poly32.ConstantWidth = 0;
                    Poly32.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly32);
                    acTrans.AddNewlyCreatedDBObject(Poly32, true);

                    ObjectIdCollection acObjIdColl32 = new ObjectIdCollection();
                    acObjIdColl32.Add(Poly32.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly33 = new Polyline();
                    Poly33.SetDatabaseDefaults();
                    Poly33.AddVertexAt(roomNumPoint, new Point2d(x - 25.620, y + 9.940), 0, 0, 0);
                    Poly33.AddVertexAt(++roomNumPoint, new Point2d(x - 25.210, y + 9.940), 0, 0, 0);
                    Poly33.AddVertexAt(++roomNumPoint, new Point2d(x - 25.120, y + 9.850), 0, 0, 0);
                    Poly33.AddVertexAt(++roomNumPoint, new Point2d(x - 25.120, y + 9.590), 0, 0, 0);
                    Poly33.AddVertexAt(++roomNumPoint, new Point2d(x - 25.150, y + 9.530), 0, 0, 0);
                    Poly33.AddVertexAt(++roomNumPoint, new Point2d(x - 25.280, y + 9.410), 0, 0, 0);
                    Poly33.AddVertexAt(++roomNumPoint, new Point2d(x - 25.340, y + 9.370), 0, 0, 0);
                    Poly33.AddVertexAt(++roomNumPoint, new Point2d(x - 25.620, y + 9.370), 0, 0, 0);
                    Poly33.Closed = true;
                    Poly33.LineWeight = 0;
                    Poly33.ConstantWidth = 0;
                    Poly33.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly33);
                    acTrans.AddNewlyCreatedDBObject(Poly33, true);

                    ObjectIdCollection acObjIdColl33 = new ObjectIdCollection();
                    acObjIdColl33.Add(Poly33.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    roomNumPoint = 0;
                    Polyline Poly34 = new Polyline();
                    Poly34.SetDatabaseDefaults();
                    Poly34.AddVertexAt(roomNumPoint, new Point2d(x - 25.780 + 11.370, y + 8.650), 0, 0, 0);
                    Poly34.AddVertexAt(++roomNumPoint, new Point2d(x - 25.620 + 11.370, y + 8.650), 0, 0, 0);
                    Poly34.AddVertexAt(++roomNumPoint, new Point2d(x - 25.620 + 11.370, y + 9.280), 0, 0, 0);
                    Poly34.AddVertexAt(++roomNumPoint, new Point2d(x - 25.430 + 11.370, y + 9.280), 0, 0, 0);
                    Poly34.AddVertexAt(++roomNumPoint, new Point2d(x - 25.240 + 11.370, y + 9.310), 0, 0, 0);
                    Poly34.AddVertexAt(++roomNumPoint, new Point2d(x - 25.130 + 11.370, y + 9.370), 0, 0, 0);
                    Poly34.AddVertexAt(++roomNumPoint, new Point2d(x - 25.030 + 11.370, y + 9.440), 0, 0, 0);
                    Poly34.AddVertexAt(++roomNumPoint, new Point2d(x - 24.960 + 11.370, y + 9.530), 0, 0, 0);
                    Poly34.AddVertexAt(++roomNumPoint, new Point2d(x - 24.930 + 11.370, y + 9.590), 0, 0, 0);
                    Poly34.AddVertexAt(++roomNumPoint, new Point2d(x - 24.930 + 11.370, y + 9.850), 0, 0, 0);
                    Poly34.AddVertexAt(++roomNumPoint, new Point2d(x - 24.990 + 11.370, y + 9.940), 0, 0, 0);
                    Poly34.AddVertexAt(++roomNumPoint, new Point2d(x - 25.060 + 11.370, y + 9.970), 0, 0, 0);
                    Poly34.AddVertexAt(++roomNumPoint, new Point2d(x - 25.120 + 11.370, y + 10.000), 0, 0, 0);
                    Poly34.AddVertexAt(++roomNumPoint, new Point2d(x - 25.310 + 11.370, y + 10.030), 0, 0, 0);
                    Poly34.AddVertexAt(++roomNumPoint, new Point2d(x - 25.530 + 11.370, y + 10.030), 0, 0, 0);
                    Poly34.AddVertexAt(++roomNumPoint, new Point2d(x - 25.780 + 11.370, y + 10.030), 0, 0, 0);
                    Poly34.Closed = true;
                    Poly34.LineWeight = 0;
                    Poly34.ConstantWidth = 0;
                    Poly34.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly34);
                    acTrans.AddNewlyCreatedDBObject(Poly34, true);

                    ObjectIdCollection acObjIdColl34 = new ObjectIdCollection();
                    acObjIdColl34.Add(Poly34.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly35 = new Polyline();
                    Poly35.SetDatabaseDefaults();
                    Poly35.AddVertexAt(roomNumPoint, new Point2d(x - 25.620 + 11.370, y + 9.940), 0, 0, 0);
                    Poly35.AddVertexAt(++roomNumPoint, new Point2d(x - 25.210 + 11.370, y + 9.940), 0, 0, 0);
                    Poly35.AddVertexAt(++roomNumPoint, new Point2d(x - 25.120 + 11.370, y + 9.850), 0, 0, 0);
                    Poly35.AddVertexAt(++roomNumPoint, new Point2d(x - 25.120 + 11.370, y + 9.590), 0, 0, 0);
                    Poly35.AddVertexAt(++roomNumPoint, new Point2d(x - 25.150 + 11.370, y + 9.530), 0, 0, 0);
                    Poly35.AddVertexAt(++roomNumPoint, new Point2d(x - 25.280 + 11.370, y + 9.410), 0, 0, 0);
                    Poly35.AddVertexAt(++roomNumPoint, new Point2d(x - 25.340 + 11.370, y + 9.370), 0, 0, 0);
                    Poly35.AddVertexAt(++roomNumPoint, new Point2d(x - 25.620 + 11.370, y + 9.370), 0, 0, 0);
                    Poly35.Closed = true;
                    Poly35.LineWeight = 0;
                    Poly35.ConstantWidth = 0;
                    Poly35.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly35);
                    acTrans.AddNewlyCreatedDBObject(Poly35, true);

                    ObjectIdCollection acObjIdColl35 = new ObjectIdCollection();
                    acObjIdColl35.Add(Poly35.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    roomNumPoint = 0;
                    Polyline Poly36 = new Polyline();
                    Poly36.SetDatabaseDefaults();
                    Poly36.AddVertexAt(roomNumPoint, new Point2d(x - 25.780 + 22.080, y + 8.650), 0, 0, 0);
                    Poly36.AddVertexAt(++roomNumPoint, new Point2d(x - 25.620 + 22.080, y + 8.650), 0, 0, 0);
                    Poly36.AddVertexAt(++roomNumPoint, new Point2d(x - 25.620 + 22.080, y + 9.280), 0, 0, 0);
                    Poly36.AddVertexAt(++roomNumPoint, new Point2d(x - 25.430 + 22.080, y + 9.280), 0, 0, 0);
                    Poly36.AddVertexAt(++roomNumPoint, new Point2d(x - 25.240 + 22.080, y + 9.310), 0, 0, 0);
                    Poly36.AddVertexAt(++roomNumPoint, new Point2d(x - 25.130 + 22.080, y + 9.370), 0, 0, 0);
                    Poly36.AddVertexAt(++roomNumPoint, new Point2d(x - 25.030 + 22.080, y + 9.440), 0, 0, 0);
                    Poly36.AddVertexAt(++roomNumPoint, new Point2d(x - 24.960 + 22.080, y + 9.530), 0, 0, 0);
                    Poly36.AddVertexAt(++roomNumPoint, new Point2d(x - 24.930 + 22.080, y + 9.590), 0, 0, 0);
                    Poly36.AddVertexAt(++roomNumPoint, new Point2d(x - 24.930 + 22.080, y + 9.850), 0, 0, 0);
                    Poly36.AddVertexAt(++roomNumPoint, new Point2d(x - 24.990 + 22.080, y + 9.940), 0, 0, 0);
                    Poly36.AddVertexAt(++roomNumPoint, new Point2d(x - 25.060 + 22.080, y + 9.970), 0, 0, 0);
                    Poly36.AddVertexAt(++roomNumPoint, new Point2d(x - 25.120 + 22.080, y + 10.000), 0, 0, 0);
                    Poly36.AddVertexAt(++roomNumPoint, new Point2d(x - 25.310 + 22.080, y + 10.030), 0, 0, 0);
                    Poly36.AddVertexAt(++roomNumPoint, new Point2d(x - 25.530 + 22.080, y + 10.030), 0, 0, 0);
                    Poly36.AddVertexAt(++roomNumPoint, new Point2d(x - 25.780 + 22.080, y + 10.030), 0, 0, 0);
                    Poly36.Closed = true;
                    Poly36.LineWeight = 0;
                    Poly36.ConstantWidth = 0;
                    Poly36.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly36);
                    acTrans.AddNewlyCreatedDBObject(Poly36, true);

                    ObjectIdCollection acObjIdColl36 = new ObjectIdCollection();
                    acObjIdColl36.Add(Poly36.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly37 = new Polyline();
                    Poly37.SetDatabaseDefaults();
                    Poly37.AddVertexAt(roomNumPoint, new Point2d(x - 25.620 + 22.080, y + 9.940), 0, 0, 0);
                    Poly37.AddVertexAt(++roomNumPoint, new Point2d(x - 25.210 + 22.080, y + 9.940), 0, 0, 0);
                    Poly37.AddVertexAt(++roomNumPoint, new Point2d(x - 25.120 + 22.080, y + 9.850), 0, 0, 0);
                    Poly37.AddVertexAt(++roomNumPoint, new Point2d(x - 25.120 + 22.080, y + 9.590), 0, 0, 0);
                    Poly37.AddVertexAt(++roomNumPoint, new Point2d(x - 25.150 + 22.080, y + 9.530), 0, 0, 0);
                    Poly37.AddVertexAt(++roomNumPoint, new Point2d(x - 25.280 + 22.080, y + 9.410), 0, 0, 0);
                    Poly37.AddVertexAt(++roomNumPoint, new Point2d(x - 25.340 + 22.080, y + 9.370), 0, 0, 0);
                    Poly37.AddVertexAt(++roomNumPoint, new Point2d(x - 25.620 + 22.080, y + 9.370), 0, 0, 0);
                    Poly37.Closed = true;
                    Poly37.LineWeight = 0;
                    Poly37.ConstantWidth = 0;
                    Poly37.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly37);
                    acTrans.AddNewlyCreatedDBObject(Poly37, true);

                    ObjectIdCollection acObjIdColl37 = new ObjectIdCollection();
                    acObjIdColl37.Add(Poly37.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    // X
                    roomNumPoint = 0;
                    Polyline Poly38 = new Polyline();
                    Poly38.SetDatabaseDefaults();
                    Poly38.AddVertexAt(roomNumPoint, new Point2d(x - 24.150, y + 9.370), 0, 0, 0);
                    Poly38.AddVertexAt(++roomNumPoint, new Point2d(x - 24.590, y + 8.650), 0, 0, 0);
                    Poly38.AddVertexAt(++roomNumPoint, new Point2d(x - 24.430, y + 8.650), 0, 0, 0);
                    Poly38.AddVertexAt(++roomNumPoint, new Point2d(x - 24.370, y + 8.810), 0, 0, 0);
                    Poly38.AddVertexAt(++roomNumPoint, new Point2d(x - 24.180, y + 9.190), 0, 0, 0);
                    Poly38.AddVertexAt(++roomNumPoint, new Point2d(x - 24.050, y + 9.250), 0, 0, 0);
                    Poly38.AddVertexAt(++roomNumPoint, new Point2d(x - 23.740, y + 8.650), 0, 0, 0);
                    Poly38.AddVertexAt(++roomNumPoint, new Point2d(x - 23.520, y + 8.650), 0, 0, 0);
                    Poly38.AddVertexAt(++roomNumPoint, new Point2d(x - 23.990, y + 9.440), 0, 0, 0);
                    Poly38.AddVertexAt(++roomNumPoint, new Point2d(x - 23.580, y + 10.030), 0, 0, 0);
                    Poly38.AddVertexAt(++roomNumPoint, new Point2d(x - 23.670, y + 10.000), 0, 0, 0);
                    Poly38.AddVertexAt(++roomNumPoint, new Point2d(x - 23.740, y + 9.970), 0, 0, 0);
                    Poly38.AddVertexAt(++roomNumPoint, new Point2d(x - 23.800, y + 9.910), 0, 0, 0);
                    Poly38.AddVertexAt(++roomNumPoint, new Point2d(x - 23.830, y + 9.810), 0, 0, 0);
                    Poly38.AddVertexAt(++roomNumPoint, new Point2d(x - 23.930, y + 9.660), 0, 0, 0);
                    Poly38.AddVertexAt(++roomNumPoint, new Point2d(x - 24.050, y + 9.500), 0, 0, 0);
                    Poly38.AddVertexAt(++roomNumPoint, new Point2d(x - 24.300, y + 9.970), 0, 0, 0);
                    Poly38.AddVertexAt(++roomNumPoint, new Point2d(x - 24.400, y + 10.030), 0, 0, 0);
                    Poly38.AddVertexAt(++roomNumPoint, new Point2d(x - 24.550, y + 10.030), 0, 0, 0);
                    Poly38.Closed = true;
                    Poly38.LineWeight = 0;
                    Poly38.ConstantWidth = 0;
                    Poly38.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly38);
                    acTrans.AddNewlyCreatedDBObject(Poly38, true);

                    ObjectIdCollection acObjIdColl38 = new ObjectIdCollection();
                    acObjIdColl38.Add(Poly38.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    // И
                    roomNumPoint = 0;
                    Polyline Poly39 = new Polyline();
                    Poly39.SetDatabaseDefaults();
                    Poly39.AddVertexAt(roomNumPoint, new Point2d(x - 23.080, y + 8.650), 0, 0, 0);
                    Poly39.AddVertexAt(++roomNumPoint, new Point2d(x - 22.070, y + 9.750), 0, 0, 0);
                    Poly39.AddVertexAt(++roomNumPoint, new Point2d(x - 22.070, y + 8.650), 0, 0, 0);
                    Poly39.AddVertexAt(++roomNumPoint, new Point2d(x - 21.880, y + 8.650), 0, 0, 0);
                    Poly39.AddVertexAt(++roomNumPoint, new Point2d(x - 21.880, y + 10.030), 0, 0, 0);
                    Poly39.AddVertexAt(++roomNumPoint, new Point2d(x - 21.980, y + 10.000), 0, 0, 0);
                    Poly39.AddVertexAt(++roomNumPoint, new Point2d(x - 22.100, y + 9.910), 0, 0, 0);
                    Poly39.AddVertexAt(++roomNumPoint, new Point2d(x - 22.890, y + 9.000), 0, 0, 0);
                    Poly39.AddVertexAt(++roomNumPoint, new Point2d(x - 22.890, y + 10.030), 0, 0, 0);
                    Poly39.AddVertexAt(++roomNumPoint, new Point2d(x - 23.080, y + 10.030), 0, 0, 0);
                    Poly39.Closed = true;
                    Poly39.LineWeight = 0;
                    Poly39.ConstantWidth = 0;
                    Poly39.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly39);
                    acTrans.AddNewlyCreatedDBObject(Poly39, true);

                    ObjectIdCollection acObjIdColl39 = new ObjectIdCollection();
                    acObjIdColl39.Add(Poly39.ObjectId);

                    //////////////////////////////////////////////////////////////////////////////////////////
                    // Т
                    roomNumPoint = 0;
                    Polyline Poly41 = new Polyline();
                    Poly41.SetDatabaseDefaults();
                    Poly41.AddVertexAt(roomNumPoint, new Point2d(x - 21.440, y + 9.940), 0, 0, 0);
                    Poly41.AddVertexAt(++roomNumPoint, new Point2d(x - 21.010, y + 9.940), 0, 0, 0);
                    Poly41.AddVertexAt(++roomNumPoint, new Point2d(x - 21.010, y + 8.650), 0, 0, 0);
                    Poly41.AddVertexAt(++roomNumPoint, new Point2d(x - 20.850, y + 8.650), 0, 0, 0);
                    Poly41.AddVertexAt(++roomNumPoint, new Point2d(x - 20.850, y + 9.970), 0, 0, 0);
                    Poly41.AddVertexAt(++roomNumPoint, new Point2d(x - 20.440, y + 9.940), 0, 0, 0);
                    Poly41.AddVertexAt(++roomNumPoint, new Point2d(x - 20.440, y + 10.030), 0, 0, 0);
                    Poly41.AddVertexAt(++roomNumPoint, new Point2d(x - 21.440, y + 10.030), 0, 0, 0);
                    Poly41.Closed = true;
                    Poly41.LineWeight = 0;
                    Poly41.ConstantWidth = 0;
                    Poly41.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly41);
                    acTrans.AddNewlyCreatedDBObject(Poly41, true);

                    ObjectIdCollection acObjIdColl41 = new ObjectIdCollection();
                    acObjIdColl41.Add(Poly41.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly40 = new Polyline();
                    Poly40.SetDatabaseDefaults();
                    Poly40.AddVertexAt(roomNumPoint, new Point2d(x - 21.440 + 3.990, y + 9.940), 0, 0, 0);
                    Poly40.AddVertexAt(++roomNumPoint, new Point2d(x - 21.010 + 3.990, y + 9.940), 0, 0, 0);
                    Poly40.AddVertexAt(++roomNumPoint, new Point2d(x - 21.010 + 3.990, y + 8.650), 0, 0, 0);
                    Poly40.AddVertexAt(++roomNumPoint, new Point2d(x - 20.850 + 3.990, y + 8.650), 0, 0, 0);
                    Poly40.AddVertexAt(++roomNumPoint, new Point2d(x - 20.850 + 3.990, y + 9.970), 0, 0, 0);
                    Poly40.AddVertexAt(++roomNumPoint, new Point2d(x - 20.440 + 3.990, y + 9.940), 0, 0, 0);
                    Poly40.AddVertexAt(++roomNumPoint, new Point2d(x - 20.440 + 3.990, y + 10.030), 0, 0, 0);
                    Poly40.AddVertexAt(++roomNumPoint, new Point2d(x - 21.440 + 3.990, y + 10.030), 0, 0, 0);
                    Poly40.Closed = true;
                    Poly40.LineWeight = 0;
                    Poly40.ConstantWidth = 0;
                    Poly40.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly40);
                    acTrans.AddNewlyCreatedDBObject(Poly40, true);

                    ObjectIdCollection acObjIdColl40 = new ObjectIdCollection();
                    acObjIdColl40.Add(Poly40.ObjectId);

                    //////////////////////////////////////////////////////////////////////////////////////////
                    // E
                    roomNumPoint = 0;
                    Polyline Poly42 = new Polyline();
                    Poly42.SetDatabaseDefaults();
                    Poly42.AddVertexAt(roomNumPoint, new Point2d(x - 20.000, y + 8.650), 0, 0, 0);
                    Poly42.AddVertexAt(++roomNumPoint, new Point2d(x - 19.280, y + 8.650), 0, 0, 0);
                    Poly42.AddVertexAt(++roomNumPoint, new Point2d(x - 19.280, y + 8.810), 0, 0, 0);
                    Poly42.AddVertexAt(++roomNumPoint, new Point2d(x - 19.430, y + 8.810), 0, 0, 0);
                    Poly42.AddVertexAt(++roomNumPoint, new Point2d(x - 19.840, y + 8.750), 0, 0, 0);
                    Poly42.AddVertexAt(++roomNumPoint, new Point2d(x - 19.840, y + 9.370), 0, 0, 0);
                    Poly42.AddVertexAt(++roomNumPoint, new Point2d(x - 19.340, y + 9.370), 0, 0, 0);
                    Poly42.AddVertexAt(++roomNumPoint, new Point2d(x - 19.340, y + 9.470), 0, 0, 0);
                    Poly42.AddVertexAt(++roomNumPoint, new Point2d(x - 19.840, y + 9.470), 0, 0, 0);
                    Poly42.AddVertexAt(++roomNumPoint, new Point2d(x - 19.840, y + 9.940), 0, 0, 0);
                    Poly42.AddVertexAt(++roomNumPoint, new Point2d(x - 19.310, y + 9.940), 0, 0, 0);
                    Poly42.AddVertexAt(++roomNumPoint, new Point2d(x - 19.310, y + 10.030), 0, 0, 0);
                    Poly42.AddVertexAt(++roomNumPoint, new Point2d(x - 20.000, y + 10.030), 0, 0, 0);
                    Poly42.Closed = true;
                    Poly42.LineWeight = 0;
                    Poly42.ConstantWidth = 0;
                    Poly42.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly42);
                    acTrans.AddNewlyCreatedDBObject(Poly42, true);

                    ObjectIdCollection acObjIdColl42 = new ObjectIdCollection();
                    acObjIdColl42.Add(Poly42.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly43 = new Polyline();
                    Poly43.SetDatabaseDefaults();
                    Poly43.AddVertexAt(roomNumPoint, new Point2d(x - 20.000 + 10.580, y + 8.650), 0, 0, 0);
                    Poly43.AddVertexAt(++roomNumPoint, new Point2d(x - 19.280 + 10.580, y + 8.650), 0, 0, 0);
                    Poly43.AddVertexAt(++roomNumPoint, new Point2d(x - 19.280 + 10.580, y + 8.810), 0, 0, 0);
                    Poly43.AddVertexAt(++roomNumPoint, new Point2d(x - 19.430 + 10.580, y + 8.810), 0, 0, 0);
                    Poly43.AddVertexAt(++roomNumPoint, new Point2d(x - 19.840 + 10.580, y + 8.750), 0, 0, 0);
                    Poly43.AddVertexAt(++roomNumPoint, new Point2d(x - 19.840 + 10.580, y + 9.370), 0, 0, 0);
                    Poly43.AddVertexAt(++roomNumPoint, new Point2d(x - 19.340 + 10.580, y + 9.370), 0, 0, 0);
                    Poly43.AddVertexAt(++roomNumPoint, new Point2d(x - 19.340 + 10.580, y + 9.470), 0, 0, 0);
                    Poly43.AddVertexAt(++roomNumPoint, new Point2d(x - 19.840 + 10.580, y + 9.470), 0, 0, 0);
                    Poly43.AddVertexAt(++roomNumPoint, new Point2d(x - 19.840 + 10.580, y + 9.940), 0, 0, 0);
                    Poly43.AddVertexAt(++roomNumPoint, new Point2d(x - 19.310 + 10.580, y + 9.940), 0, 0, 0);
                    Poly43.AddVertexAt(++roomNumPoint, new Point2d(x - 19.310 + 10.580, y + 10.030), 0, 0, 0);
                    Poly43.AddVertexAt(++roomNumPoint, new Point2d(x - 20.000 + 10.580, y + 10.030), 0, 0, 0);
                    Poly43.Closed = true;
                    Poly43.LineWeight = 0;
                    Poly43.ConstantWidth = 0;
                    Poly43.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly43);
                    acTrans.AddNewlyCreatedDBObject(Poly43, true);

                    ObjectIdCollection acObjIdColl43 = new ObjectIdCollection();
                    acObjIdColl43.Add(Poly43.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    // К
                    roomNumPoint = 0;
                    Polyline Poly44 = new Polyline();
                    Poly44.SetDatabaseDefaults();
                    Poly44.AddVertexAt(roomNumPoint, new Point2d(x - 18.810, y + 8.650), 0, 0, 0);
                    Poly44.AddVertexAt(++roomNumPoint, new Point2d(x - 18.650, y + 8.650), 0, 0, 0);
                    Poly44.AddVertexAt(++roomNumPoint, new Point2d(x - 18.620, y + 9.370), 0, 0, 0);
                    Poly44.AddVertexAt(++roomNumPoint, new Point2d(x - 18.050, y + 8.650), 0, 0, 0);
                    Poly44.AddVertexAt(++roomNumPoint, new Point2d(x - 17.770, y + 8.680), 0, 0, 0);
                    Poly44.AddVertexAt(++roomNumPoint, new Point2d(x - 18.460, y + 9.440), 0, 0, 0);
                    Poly44.AddVertexAt(++roomNumPoint, new Point2d(x - 17.900, y + 10.030), 0, 0, 0);
                    Poly44.AddVertexAt(++roomNumPoint, new Point2d(x - 18.270, y + 9.750), 0, 0, 0);
                    Poly44.AddVertexAt(++roomNumPoint, new Point2d(x - 18.650, y + 9.410), 0, 0, 0);
                    Poly44.AddVertexAt(++roomNumPoint, new Point2d(x - 18.650, y + 10.030), 0, 0, 0);
                    Poly44.AddVertexAt(++roomNumPoint, new Point2d(x - 18.810, y + 10.030), 0, 0, 0);
                    Poly44.Closed = true;
                    Poly44.LineWeight = 0;
                    Poly44.ConstantWidth = 0;
                    Poly44.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly44);
                    acTrans.AddNewlyCreatedDBObject(Poly44, true);

                    ObjectIdCollection acObjIdColl44 = new ObjectIdCollection();
                    acObjIdColl44.Add(Poly44.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    // У
                    roomNumPoint = 0;
                    Polyline Poly45 = new Polyline();
                    Poly45.SetDatabaseDefaults();
                    Poly45.AddVertexAt(roomNumPoint, new Point2d(x - 15.450, y + 9.060), 0, 0, 0);
                    Poly45.AddVertexAt(++roomNumPoint, new Point2d(x - 15.670, y + 8.620), 0, 0, 0);
                    Poly45.AddVertexAt(++roomNumPoint, new Point2d(x - 15.540, y + 8.650), 0, 0, 0);
                    Poly45.AddVertexAt(++roomNumPoint, new Point2d(x - 14.850, y + 10.000), 0, 0, 0);
                    Poly45.AddVertexAt(++roomNumPoint, new Point2d(x - 14.980, y + 10.000), 0, 0, 0);
                    Poly45.AddVertexAt(++roomNumPoint, new Point2d(x - 15.320, y + 9.250), 0, 0, 0);
                    Poly45.AddVertexAt(++roomNumPoint, new Point2d(x - 15.350, y + 9.250), 0, 0, 0);
                    Poly45.AddVertexAt(++roomNumPoint, new Point2d(x - 15.850, y + 10.000), 0, 0, 0);
                    Poly45.AddVertexAt(++roomNumPoint, new Point2d(x - 16.110, y + 10.030), 0, 0, 0);
                    Poly45.Closed = true;
                    Poly45.LineWeight = 0;
                    Poly45.ConstantWidth = 0;
                    Poly45.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly45);
                    acTrans.AddNewlyCreatedDBObject(Poly45, true);

                    ObjectIdCollection acObjIdColl45 = new ObjectIdCollection();
                    acObjIdColl45.Add(Poly45.ObjectId);

                    //////////////////////////////////////////////////////////////////////////////////////////
                    // Н
                    roomNumPoint = 0;
                    Polyline Poly46 = new Polyline();
                    Poly46.SetDatabaseDefaults();
                    Poly46.AddVertexAt(roomNumPoint, new Point2d(x - 13.090, y + 8.650), 0, 0, 0);
                    Poly46.AddVertexAt(++roomNumPoint, new Point2d(x - 12.900, y + 8.650), 0, 0, 0);
                    Poly46.AddVertexAt(++roomNumPoint, new Point2d(x - 12.900, y + 9.370), 0, 0, 0);
                    Poly46.AddVertexAt(++roomNumPoint, new Point2d(x - 12.050, y + 9.370), 0, 0, 0);
                    Poly46.AddVertexAt(++roomNumPoint, new Point2d(x - 12.050, y + 8.650), 0, 0, 0);
                    Poly46.AddVertexAt(++roomNumPoint, new Point2d(x - 11.900, y + 8.650), 0, 0, 0);
                    Poly46.AddVertexAt(++roomNumPoint, new Point2d(x - 11.900, y + 10.030), 0, 0, 0);
                    Poly46.AddVertexAt(++roomNumPoint, new Point2d(x - 12.050, y + 10.030), 0, 0, 0);
                    Poly46.AddVertexAt(++roomNumPoint, new Point2d(x - 12.050, y + 9.470), 0, 0, 0);
                    Poly46.AddVertexAt(++roomNumPoint, new Point2d(x - 12.900, y + 9.470), 0, 0, 0);
                    Poly46.AddVertexAt(++roomNumPoint, new Point2d(x - 12.900, y + 10.030), 0, 0, 0);
                    Poly46.AddVertexAt(++roomNumPoint, new Point2d(x - 13.090, y + 10.030), 0, 0, 0);
                    Poly46.Closed = true;
                    Poly46.LineWeight = 0;
                    Poly46.ConstantWidth = 0;
                    Poly46.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly46);
                    acTrans.AddNewlyCreatedDBObject(Poly46, true);

                    ObjectIdCollection acObjIdColl46 = new ObjectIdCollection();
                    acObjIdColl46.Add(Poly46.ObjectId);

                    /////////////////////////////////////////////////////////////////////////////////////////
                    // архитектурнОе бюрО
                    roomNumPoint = 0;
                    Polyline Poly47 = new Polyline();
                    Poly47.SetDatabaseDefaults();
                    Poly47.AddVertexAt(roomNumPoint, new Point2d(x - 34.820 + 23.460, y + 9.270), 0, 0, 0);
                    Poly47.AddVertexAt(++roomNumPoint, new Point2d(x - 34.750 + 23.460, y + 9.030), 0, 0, 0);
                    Poly47.AddVertexAt(++roomNumPoint, new Point2d(x - 34.640 + 23.460, y + 8.870), 0, 0, 0);
                    Poly47.AddVertexAt(++roomNumPoint, new Point2d(x - 34.510 + 23.460, y + 8.760), 0, 0, 0);
                    Poly47.AddVertexAt(++roomNumPoint, new Point2d(x - 34.360 + 23.460, y + 8.680), 0, 0, 0);
                    Poly47.AddVertexAt(++roomNumPoint, new Point2d(x - 34.190 + 23.460, y + 8.640), 0, 0, 0);
                    Poly47.AddVertexAt(++roomNumPoint, new Point2d(x - 34.010 + 23.460, y + 8.640), 0, 0, 0);
                    Poly47.AddVertexAt(++roomNumPoint, new Point2d(x - 33.850 + 23.460, y + 8.680), 0, 0, 0);
                    Poly47.AddVertexAt(++roomNumPoint, new Point2d(x - 33.700 + 23.460, y + 8.760), 0, 0, 0);
                    Poly47.AddVertexAt(++roomNumPoint, new Point2d(x - 33.580 + 23.460, y + 8.870), 0, 0, 0);
                    Poly47.AddVertexAt(++roomNumPoint, new Point2d(x - 33.470 + 23.460, y + 9.030), 0, 0, 0);
                    Poly47.AddVertexAt(++roomNumPoint, new Point2d(x - 33.390 + 23.460, y + 9.230), 0, 0, 0);
                    Poly47.AddVertexAt(++roomNumPoint, new Point2d(x - 33.400 + 23.460, y + 9.520), 0, 0, 0);
                    Poly47.AddVertexAt(++roomNumPoint, new Point2d(x - 33.470 + 23.460, y + 9.710), 0, 0, 0);
                    Poly47.AddVertexAt(++roomNumPoint, new Point2d(x - 33.600 + 23.460, y + 9.870), 0, 0, 0);
                    Poly47.AddVertexAt(++roomNumPoint, new Point2d(x - 33.760 + 23.460, y + 10.000), 0, 0, 0);
                    Poly47.AddVertexAt(++roomNumPoint, new Point2d(x - 33.910 + 23.460, y + 10.030), 0, 0, 0);
                    Poly47.AddVertexAt(++roomNumPoint, new Point2d(x - 34.300 + 23.460, y + 10.030), 0, 0, 0);
                    Poly47.AddVertexAt(++roomNumPoint, new Point2d(x - 34.460 + 23.460, y + 10.000), 0, 0, 0);
                    Poly47.AddVertexAt(++roomNumPoint, new Point2d(x - 34.630 + 23.460, y + 9.870), 0, 0, 0);
                    Poly47.AddVertexAt(++roomNumPoint, new Point2d(x - 34.750 + 23.460, y + 9.710), 0, 0, 0);
                    Poly47.AddVertexAt(++roomNumPoint, new Point2d(x - 34.820 + 23.460, y + 9.490), 0, 0, 0);
                    Poly47.Closed = true;
                    Poly47.LineWeight = 0;
                    Poly47.ConstantWidth = 0;
                    Poly47.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly47);
                    acTrans.AddNewlyCreatedDBObject(Poly47, true);

                    ObjectIdCollection acObjIdColl47 = new ObjectIdCollection();
                    acObjIdColl47.Add(Poly47.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly48 = new Polyline();
                    Poly48.SetDatabaseDefaults();
                    Poly48.AddVertexAt(roomNumPoint, new Point2d(x - 34.640 + 23.460, y + 9.440), 0, 0, 0);
                    Poly48.AddVertexAt(++roomNumPoint, new Point2d(x - 34.600 + 23.460, y + 9.660), 0, 0, 0);
                    Poly48.AddVertexAt(++roomNumPoint, new Point2d(x - 34.520 + 23.460, y + 9.820), 0, 0, 0);
                    Poly48.AddVertexAt(++roomNumPoint, new Point2d(x - 34.440 + 23.460, y + 9.920), 0, 0, 0);
                    Poly48.AddVertexAt(++roomNumPoint, new Point2d(x - 34.330 + 23.460, y + 9.970), 0, 0, 0);
                    Poly48.AddVertexAt(++roomNumPoint, new Point2d(x - 34.110 + 23.460, y + 9.980), 0, 0, 0);
                    Poly48.AddVertexAt(++roomNumPoint, new Point2d(x - 33.880 + 23.460, y + 9.970), 0, 0, 0);
                    Poly48.AddVertexAt(++roomNumPoint, new Point2d(x - 33.770 + 23.460, y + 9.920), 0, 0, 0);
                    Poly48.AddVertexAt(++roomNumPoint, new Point2d(x - 33.680 + 23.460, y + 9.820), 0, 0, 0);
                    Poly48.AddVertexAt(++roomNumPoint, new Point2d(x - 33.600 + 23.460, y + 9.660), 0, 0, 0);
                    Poly48.AddVertexAt(++roomNumPoint, new Point2d(x - 33.560 + 23.460, y + 9.440), 0, 0, 0);
                    Poly48.AddVertexAt(++roomNumPoint, new Point2d(x - 33.560 + 23.460, y + 9.160), 0, 0, 0);
                    Poly48.AddVertexAt(++roomNumPoint, new Point2d(x - 33.630 + 23.460, y + 8.970), 0, 0, 0);
                    Poly48.AddVertexAt(++roomNumPoint, new Point2d(x - 33.760 + 23.460, y + 8.820), 0, 0, 0);
                    Poly48.AddVertexAt(++roomNumPoint, new Point2d(x - 33.920 + 23.460, y + 8.750), 0, 0, 0);
                    Poly48.AddVertexAt(++roomNumPoint, new Point2d(x - 34.070 + 23.460, y + 8.720), 0, 0, 0);
                    Poly48.AddVertexAt(++roomNumPoint, new Point2d(x - 34.140 + 23.460, y + 8.720), 0, 0, 0);
                    Poly48.AddVertexAt(++roomNumPoint, new Point2d(x - 34.300 + 23.460, y + 8.750), 0, 0, 0);
                    Poly48.AddVertexAt(++roomNumPoint, new Point2d(x - 34.450 + 23.460, y + 8.820), 0, 0, 0);
                    Poly48.AddVertexAt(++roomNumPoint, new Point2d(x - 34.570 + 23.460, y + 8.970), 0, 0, 0);
                    Poly48.AddVertexAt(++roomNumPoint, new Point2d(x - 34.640 + 23.460, y + 9.150), 0, 0, 0);
                    Poly48.Closed = true;
                    Poly48.LineWeight = 0;
                    Poly48.ConstantWidth = 0;
                    Poly48.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly48);
                    acTrans.AddNewlyCreatedDBObject(Poly48, true);

                    ObjectIdCollection acObjIdColl48 = new ObjectIdCollection();
                    acObjIdColl48.Add(Poly48.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    roomNumPoint = 0;
                    Polyline Poly49 = new Polyline();
                    Poly49.SetDatabaseDefaults();
                    Poly49.AddVertexAt(roomNumPoint, new Point2d(x - 34.820 + 32.410, y + 9.270), 0, 0, 0);
                    Poly49.AddVertexAt(++roomNumPoint, new Point2d(x - 34.750 + 32.410, y + 9.030), 0, 0, 0);
                    Poly49.AddVertexAt(++roomNumPoint, new Point2d(x - 34.640 + 32.410, y + 8.870), 0, 0, 0);
                    Poly49.AddVertexAt(++roomNumPoint, new Point2d(x - 34.510 + 32.410, y + 8.760), 0, 0, 0);
                    Poly49.AddVertexAt(++roomNumPoint, new Point2d(x - 34.360 + 32.410, y + 8.680), 0, 0, 0);
                    Poly49.AddVertexAt(++roomNumPoint, new Point2d(x - 34.190 + 32.410, y + 8.640), 0, 0, 0);
                    Poly49.AddVertexAt(++roomNumPoint, new Point2d(x - 34.010 + 32.410, y + 8.640), 0, 0, 0);
                    Poly49.AddVertexAt(++roomNumPoint, new Point2d(x - 33.850 + 32.410, y + 8.680), 0, 0, 0);
                    Poly49.AddVertexAt(++roomNumPoint, new Point2d(x - 33.700 + 32.410, y + 8.760), 0, 0, 0);
                    Poly49.AddVertexAt(++roomNumPoint, new Point2d(x - 33.580 + 32.410, y + 8.870), 0, 0, 0);
                    Poly49.AddVertexAt(++roomNumPoint, new Point2d(x - 33.470 + 32.410, y + 9.030), 0, 0, 0);
                    Poly49.AddVertexAt(++roomNumPoint, new Point2d(x - 33.390 + 32.410, y + 9.230), 0, 0, 0);
                    Poly49.AddVertexAt(++roomNumPoint, new Point2d(x - 33.400 + 32.410, y + 9.520), 0, 0, 0);
                    Poly49.AddVertexAt(++roomNumPoint, new Point2d(x - 33.470 + 32.410, y + 9.710), 0, 0, 0);
                    Poly49.AddVertexAt(++roomNumPoint, new Point2d(x - 33.600 + 32.410, y + 9.870), 0, 0, 0);
                    Poly49.AddVertexAt(++roomNumPoint, new Point2d(x - 33.760 + 32.410, y + 10.000), 0, 0, 0);
                    Poly49.AddVertexAt(++roomNumPoint, new Point2d(x - 33.910 + 32.410, y + 10.030), 0, 0, 0);
                    Poly49.AddVertexAt(++roomNumPoint, new Point2d(x - 34.300 + 32.410, y + 10.030), 0, 0, 0);
                    Poly49.AddVertexAt(++roomNumPoint, new Point2d(x - 34.460 + 32.410, y + 10.000), 0, 0, 0);
                    Poly49.AddVertexAt(++roomNumPoint, new Point2d(x - 34.630 + 32.410, y + 9.870), 0, 0, 0);
                    Poly49.AddVertexAt(++roomNumPoint, new Point2d(x - 34.750 + 32.410, y + 9.710), 0, 0, 0);
                    Poly49.AddVertexAt(++roomNumPoint, new Point2d(x - 34.820 + 32.410, y + 9.490), 0, 0, 0);
                    Poly49.Closed = true;
                    Poly49.LineWeight = 0;
                    Poly49.ConstantWidth = 0;
                    Poly49.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly49);
                    acTrans.AddNewlyCreatedDBObject(Poly49, true);

                    ObjectIdCollection acObjIdColl49 = new ObjectIdCollection();
                    acObjIdColl49.Add(Poly49.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly50 = new Polyline();
                    Poly50.SetDatabaseDefaults();
                    Poly50.AddVertexAt(roomNumPoint, new Point2d(x - 34.640 + 32.410, y + 9.440), 0, 0, 0);
                    Poly50.AddVertexAt(++roomNumPoint, new Point2d(x - 34.600 + 32.410, y + 9.660), 0, 0, 0);
                    Poly50.AddVertexAt(++roomNumPoint, new Point2d(x - 34.520 + 32.410, y + 9.820), 0, 0, 0);
                    Poly50.AddVertexAt(++roomNumPoint, new Point2d(x - 34.440 + 32.410, y + 9.920), 0, 0, 0);
                    Poly50.AddVertexAt(++roomNumPoint, new Point2d(x - 34.330 + 32.410, y + 9.970), 0, 0, 0);
                    Poly50.AddVertexAt(++roomNumPoint, new Point2d(x - 34.110 + 32.410, y + 9.980), 0, 0, 0);
                    Poly50.AddVertexAt(++roomNumPoint, new Point2d(x - 33.880 + 32.410, y + 9.970), 0, 0, 0);
                    Poly50.AddVertexAt(++roomNumPoint, new Point2d(x - 33.770 + 32.410, y + 9.920), 0, 0, 0);
                    Poly50.AddVertexAt(++roomNumPoint, new Point2d(x - 33.680 + 32.410, y + 9.820), 0, 0, 0);
                    Poly50.AddVertexAt(++roomNumPoint, new Point2d(x - 33.600 + 32.410, y + 9.660), 0, 0, 0);
                    Poly50.AddVertexAt(++roomNumPoint, new Point2d(x - 33.560 + 32.410, y + 9.440), 0, 0, 0);
                    Poly50.AddVertexAt(++roomNumPoint, new Point2d(x - 33.560 + 32.410, y + 9.160), 0, 0, 0);
                    Poly50.AddVertexAt(++roomNumPoint, new Point2d(x - 33.630 + 32.410, y + 8.970), 0, 0, 0);
                    Poly50.AddVertexAt(++roomNumPoint, new Point2d(x - 33.760 + 32.410, y + 8.820), 0, 0, 0);
                    Poly50.AddVertexAt(++roomNumPoint, new Point2d(x - 33.920 + 32.410, y + 8.750), 0, 0, 0);
                    Poly50.AddVertexAt(++roomNumPoint, new Point2d(x - 34.070 + 32.410, y + 8.720), 0, 0, 0);
                    Poly50.AddVertexAt(++roomNumPoint, new Point2d(x - 34.140 + 32.410, y + 8.720), 0, 0, 0);
                    Poly50.AddVertexAt(++roomNumPoint, new Point2d(x - 34.300 + 32.410, y + 8.750), 0, 0, 0);
                    Poly50.AddVertexAt(++roomNumPoint, new Point2d(x - 34.450 + 32.410, y + 8.820), 0, 0, 0);
                    Poly50.AddVertexAt(++roomNumPoint, new Point2d(x - 34.570 + 32.410, y + 8.970), 0, 0, 0);
                    Poly50.AddVertexAt(++roomNumPoint, new Point2d(x - 34.640 + 32.410, y + 9.150), 0, 0, 0);
                    Poly50.Closed = true;
                    Poly50.LineWeight = 0;
                    Poly50.ConstantWidth = 0;
                    Poly50.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly50);
                    acTrans.AddNewlyCreatedDBObject(Poly50, true);

                    ObjectIdCollection acObjIdColl50 = new ObjectIdCollection();
                    acObjIdColl50.Add(Poly50.ObjectId);

                    //////////////////////////////////////////////////////////////////////////////////////////
                    // Б
                    roomNumPoint = 0;
                    Polyline Poly51 = new Polyline();
                    Poly51.SetDatabaseDefaults();
                    Poly51.AddVertexAt(roomNumPoint, new Point2d(x - 7.380, y + 8.650), 0, 0, 0);
                    Poly51.AddVertexAt(++roomNumPoint, new Point2d(x - 7.03, y + 8.65), 0, 0, 0);
                    Poly51.AddVertexAt(++roomNumPoint, new Point2d(x - 6.81, y + 8.68), 0, 0, 0);
                    Poly51.AddVertexAt(++roomNumPoint, new Point2d(x - 6.65, y + 8.75), 0, 0, 0);
                    Poly51.AddVertexAt(++roomNumPoint, new Point2d(x - 6.53, y + 8.87), 0, 0, 0);
                    Poly51.AddVertexAt(++roomNumPoint, new Point2d(x - 6.50, y + 8.97), 0, 0, 0);
                    Poly51.AddVertexAt(++roomNumPoint, new Point2d(x - 6.50, y + 9.22), 0, 0, 0);
                    Poly51.AddVertexAt(++roomNumPoint, new Point2d(x - 6.53, y + 9.28), 0, 0, 0);
                    Poly51.AddVertexAt(++roomNumPoint, new Point2d(x - 6.59, y + 9.34), 0, 0, 0);
                    Poly51.AddVertexAt(++roomNumPoint, new Point2d(x - 6.65, y + 9.37), 0, 0, 0);
                    Poly51.AddVertexAt(++roomNumPoint, new Point2d(x - 6.81, y + 9.44), 0, 0, 0);
                    Poly51.AddVertexAt(++roomNumPoint, new Point2d(x - 7.00, y + 9.47), 0, 0, 0);
                    Poly51.AddVertexAt(++roomNumPoint, new Point2d(x - 7.22, y + 9.47), 0, 0, 0);
                    Poly51.AddVertexAt(++roomNumPoint, new Point2d(x - 7.22, y + 9.94), 0, 0, 0);
                    Poly51.AddVertexAt(++roomNumPoint, new Point2d(x - 6.59, y + 9.94), 0, 0, 0);
                    Poly51.AddVertexAt(++roomNumPoint, new Point2d(x - 6.59, y + 10.03), 0, 0, 0);
                    Poly51.AddVertexAt(++roomNumPoint, new Point2d(x - 7.38, y + 10.03), 0, 0, 0);
                    Poly51.Closed = true;
                    Poly51.LineWeight = 0;
                    Poly51.ConstantWidth = 0;
                    Poly51.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly51);
                    acTrans.AddNewlyCreatedDBObject(Poly51, true);

                    ObjectIdCollection acObjIdColl51 = new ObjectIdCollection();
                    acObjIdColl51.Add(Poly51.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly52 = new Polyline();
                    Poly52.SetDatabaseDefaults();
                    Poly52.AddVertexAt(roomNumPoint, new Point2d(x - 6.97, y + 8.75), 0, 0, 0);
                    Poly52.AddVertexAt(++roomNumPoint, new Point2d(x - 7.22, y + 8.75), 0, 0, 0);
                    Poly52.AddVertexAt(++roomNumPoint, new Point2d(x - 7.22, y + 9.37), 0, 0, 0);
                    Poly52.AddVertexAt(++roomNumPoint, new Point2d(x - 6.87, y + 9.37), 0, 0, 0);
                    Poly52.AddVertexAt(++roomNumPoint, new Point2d(x - 6.78, y + 9.34), 0, 0, 0);
                    Poly52.AddVertexAt(++roomNumPoint, new Point2d(x - 6.72, y + 9.28), 0, 0, 0);
                    Poly52.AddVertexAt(++roomNumPoint, new Point2d(x - 6.68, y + 9.19), 0, 0, 0);
                    Poly52.AddVertexAt(++roomNumPoint, new Point2d(x - 6.68, y + 8.97), 0, 0, 0);
                    Poly52.AddVertexAt(++roomNumPoint, new Point2d(x - 6.75, y + 8.87), 0, 0, 0);
                    Poly52.AddVertexAt(++roomNumPoint, new Point2d(x - 6.84, y + 8.78), 0, 0, 0);
                    Poly52.Closed = true;
                    Poly52.LineWeight = 0;
                    Poly52.ConstantWidth = 0;
                    Poly52.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly52);
                    acTrans.AddNewlyCreatedDBObject(Poly52, true);

                    ObjectIdCollection acObjIdColl52 = new ObjectIdCollection();
                    acObjIdColl52.Add(Poly52.ObjectId);

                    /////////////////////////////////////////////////////////////////////////////////////////
                    // Ю
                    roomNumPoint = 0;
                    Polyline Poly53 = new Polyline();
                    Poly53.SetDatabaseDefaults();
                    Poly53.AddVertexAt(roomNumPoint, new Point2d(x - 34.820 + 29.18, y + 9.270), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 34.750 + 29.18, y + 9.030), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 34.640 + 29.18, y + 8.870), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 34.510 + 29.18, y + 8.760), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 34.360 + 29.18, y + 8.680), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 34.190 + 29.18, y + 8.640), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 34.010 + 29.18, y + 8.640), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 33.850 + 29.18, y + 8.680), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 33.700 + 29.18, y + 8.760), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 33.580 + 29.18, y + 8.870), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 33.470 + 29.18, y + 9.030), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 33.390 + 29.18, y + 9.230), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 33.400 + 29.18, y + 9.520), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 33.470 + 29.18, y + 9.710), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 33.600 + 29.18, y + 9.870), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 33.760 + 29.18, y + 10.000), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 33.910 + 29.18, y + 10.030), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 34.300 + 29.18, y + 10.030), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 34.460 + 29.18, y + 10.000), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 34.630 + 29.18, y + 9.870), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 34.750 + 29.18, y + 9.710), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 34.820 + 29.18, y + 9.490), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 34.820 + 29.18, y + 9.460), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 35.01 + 29.18, y + 9.460), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 35.01 + 29.18, y + 10.04), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 35.20 + 29.18, y + 10.04), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 35.20 + 29.18, y + 8.64), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 35.01 + 29.18, y + 8.64), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 35.01 + 29.18, y + 9.37), 0, 0, 0);
                    Poly53.AddVertexAt(++roomNumPoint, new Point2d(x - 34.82 + 29.18, y + 9.37), 0, 0, 0);
                    Poly53.Closed = true;
                    Poly53.LineWeight = 0;
                    Poly53.ConstantWidth = 0;
                    Poly53.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly53);
                    acTrans.AddNewlyCreatedDBObject(Poly53, true);

                    ObjectIdCollection acObjIdColl53 = new ObjectIdCollection();
                    acObjIdColl53.Add(Poly53.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly54 = new Polyline();
                    Poly54.SetDatabaseDefaults();
                    Poly54.AddVertexAt(roomNumPoint, new Point2d(x - 34.640 + 29.18, y + 9.440), 0, 0, 0);
                    Poly54.AddVertexAt(++roomNumPoint, new Point2d(x - 34.600 + 29.18, y + 9.660), 0, 0, 0);
                    Poly54.AddVertexAt(++roomNumPoint, new Point2d(x - 34.520 + 29.18, y + 9.820), 0, 0, 0);
                    Poly54.AddVertexAt(++roomNumPoint, new Point2d(x - 34.440 + 29.18, y + 9.920), 0, 0, 0);
                    Poly54.AddVertexAt(++roomNumPoint, new Point2d(x - 34.330 + 29.18, y + 9.970), 0, 0, 0);
                    Poly54.AddVertexAt(++roomNumPoint, new Point2d(x - 34.110 + 29.18, y + 9.980), 0, 0, 0);
                    Poly54.AddVertexAt(++roomNumPoint, new Point2d(x - 33.880 + 29.18, y + 9.970), 0, 0, 0);
                    Poly54.AddVertexAt(++roomNumPoint, new Point2d(x - 33.770 + 29.18, y + 9.920), 0, 0, 0);
                    Poly54.AddVertexAt(++roomNumPoint, new Point2d(x - 33.680 + 29.18, y + 9.820), 0, 0, 0);
                    Poly54.AddVertexAt(++roomNumPoint, new Point2d(x - 33.600 + 29.18, y + 9.660), 0, 0, 0);
                    Poly54.AddVertexAt(++roomNumPoint, new Point2d(x - 33.560 + 29.18, y + 9.440), 0, 0, 0);
                    Poly54.AddVertexAt(++roomNumPoint, new Point2d(x - 33.560 + 29.18, y + 9.160), 0, 0, 0);
                    Poly54.AddVertexAt(++roomNumPoint, new Point2d(x - 33.630 + 29.18, y + 8.970), 0, 0, 0);
                    Poly54.AddVertexAt(++roomNumPoint, new Point2d(x - 33.760 + 29.18, y + 8.820), 0, 0, 0);
                    Poly54.AddVertexAt(++roomNumPoint, new Point2d(x - 33.920 + 29.18, y + 8.750), 0, 0, 0);
                    Poly54.AddVertexAt(++roomNumPoint, new Point2d(x - 34.070 + 29.18, y + 8.720), 0, 0, 0);
                    Poly54.AddVertexAt(++roomNumPoint, new Point2d(x - 34.140 + 29.18, y + 8.720), 0, 0, 0);
                    Poly54.AddVertexAt(++roomNumPoint, new Point2d(x - 34.300 + 29.18, y + 8.750), 0, 0, 0);
                    Poly54.AddVertexAt(++roomNumPoint, new Point2d(x - 34.450 + 29.18, y + 8.820), 0, 0, 0);
                    Poly54.AddVertexAt(++roomNumPoint, new Point2d(x - 34.570 + 29.18, y + 8.970), 0, 0, 0);
                    Poly54.AddVertexAt(++roomNumPoint, new Point2d(x - 34.640 + 29.18, y + 9.150), 0, 0, 0);
                    Poly54.Closed = true;
                    Poly54.LineWeight = 0;
                    Poly54.ConstantWidth = 0;
                    Poly54.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly54);
                    acTrans.AddNewlyCreatedDBObject(Poly54, true);

                    ObjectIdCollection acObjIdColl54 = new ObjectIdCollection();
                    acObjIdColl54.Add(Poly54.ObjectId);

                    //////////////////////////////////////////////////////////////////////////////////////////
                    // Кавычки " большие
                    roomNumPoint = 0;
                    Polyline Poly55 = new Polyline();
                    Poly55.SetDatabaseDefaults();
                    Poly55.AddVertexAt(roomNumPoint, new Point2d(x - 34.54, y + 6.42), 0, 0, 0);
                    Poly55.AddVertexAt(++roomNumPoint, new Point2d(x - 34.41, y + 6.40), 0, 0, 0);
                    Poly55.AddVertexAt(++roomNumPoint, new Point2d(x - 34.26, y + 6.55), 0, 0, 0);
                    Poly55.AddVertexAt(++roomNumPoint, new Point2d(x - 34.10, y + 7.02), 0, 0, 0);
                    Poly55.AddVertexAt(++roomNumPoint, new Point2d(x - 33.90, y + 7.67), 0, 0, 0);
                    Poly55.AddVertexAt(++roomNumPoint, new Point2d(x - 33.95, y + 7.69), 0, 0, 0);
                    Poly55.AddVertexAt(++roomNumPoint, new Point2d(x - 34.10, y + 7.55), 0, 0, 0);
                    Poly55.AddVertexAt(++roomNumPoint, new Point2d(x - 34.57, y + 6.58), 0, 0, 0);
                    Poly55.Closed = true;
                    Poly55.LineWeight = 0;
                    Poly55.ConstantWidth = 0;
                    Poly55.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly55);
                    acTrans.AddNewlyCreatedDBObject(Poly55, true);

                    ObjectIdCollection acObjIdColl55 = new ObjectIdCollection();
                    acObjIdColl55.Add(Poly55.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly56 = new Polyline();
                    Poly56.SetDatabaseDefaults();
                    Poly56.AddVertexAt(roomNumPoint, new Point2d(x - 34.54 + 0.6, y + 6.42), 0, 0, 0);
                    Poly56.AddVertexAt(++roomNumPoint, new Point2d(x - 34.41 + 0.6, y + 6.40), 0, 0, 0);
                    Poly56.AddVertexAt(++roomNumPoint, new Point2d(x - 34.26 + 0.6, y + 6.55), 0, 0, 0);
                    Poly56.AddVertexAt(++roomNumPoint, new Point2d(x - 34.10 + 0.6, y + 7.02), 0, 0, 0);
                    Poly56.AddVertexAt(++roomNumPoint, new Point2d(x - 33.90 + 0.6, y + 7.67), 0, 0, 0);
                    Poly56.AddVertexAt(++roomNumPoint, new Point2d(x - 33.95 + 0.6, y + 7.69), 0, 0, 0);
                    Poly56.AddVertexAt(++roomNumPoint, new Point2d(x - 34.10 + 0.6, y + 7.55), 0, 0, 0);
                    Poly56.AddVertexAt(++roomNumPoint, new Point2d(x - 34.57 + 0.6, y + 6.58), 0, 0, 0);
                    Poly56.Closed = true;
                    Poly56.LineWeight = 0;
                    Poly56.ConstantWidth = 0;
                    Poly56.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly56);
                    acTrans.AddNewlyCreatedDBObject(Poly56, true);

                    ObjectIdCollection acObjIdColl56 = new ObjectIdCollection();
                    acObjIdColl56.Add(Poly56.ObjectId);

                    //////////////////////////////////////////////////////////////////////////////////////////
                    // C
                    roomNumPoint = 0;
                    Polyline Poly57 = new Polyline();
                    Poly57.SetDatabaseDefaults();
                    Poly57.AddVertexAt(roomNumPoint, new Point2d(x - 32.20, y + 5.99), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 32.17, y + 5.65), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 31.95, y + 5.12), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 31.57, y + 4.71), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 31.03, y + 4.46), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 30.53, y + 4.42), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 30.00, y + 4.52), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 29.56, y + 4.71), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 29.56, y + 4.96), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 30.00, y + 4.74), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 30.53, y + 4.61), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 31.00, y + 4.74), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 31.38, y + 4.99), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 31.63, y + 5.37), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 31.73, y + 5.81), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 31.69, y + 6.47), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 31.54, y + 6.94), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 31.22, y + 7.28), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 30.85, y + 7.50), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 30.47, y + 7.50), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 30.12, y + 7.47), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 29.90, y + 7.35), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 29.59, y + 7.16), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 29.46, y + 7.50), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 29.97, y + 7.66), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 30.25, y + 7.69), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 30.94, y + 7.69), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 31.25, y + 7.60), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 31.54, y + 7.44), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 31.76, y + 7.25), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 31.95, y + 7.00), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 32.10, y + 6.72), 0, 0, 0);
                    Poly57.AddVertexAt(++roomNumPoint, new Point2d(x - 32.17, y + 6.37), 0, 0, 0);

                    Poly57.Closed = true;
                    Poly57.LineWeight = 0;
                    Poly57.ConstantWidth = 0;
                    Poly57.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly57);
                    acTrans.AddNewlyCreatedDBObject(Poly57, true);

                    ObjectIdCollection acObjIdColl57 = new ObjectIdCollection();
                    acObjIdColl57.Add(Poly57.ObjectId);

                    //////////////////////////////////////////////////////////////////////////////////////////
                    // Т
                    roomNumPoint = 0;
                    Polyline Poly58 = new Polyline();
                    Poly58.SetDatabaseDefaults();
                    Poly58.AddVertexAt(roomNumPoint, new Point2d(x - 27.51, y + 7.44), 0, 0, 0);
                    Poly58.AddVertexAt(++roomNumPoint, new Point2d(x - 27.51, y + 4.55), 0, 0, 0);
                    Poly58.AddVertexAt(++roomNumPoint, new Point2d(x - 27.13, y + 4.55), 0, 0, 0);
                    Poly58.AddVertexAt(++roomNumPoint, new Point2d(x - 27.13, y + 7.44), 0, 0, 0);
                    Poly58.AddVertexAt(++roomNumPoint, new Point2d(x - 26.16, y + 7.41), 0, 0, 0);
                    Poly58.AddVertexAt(++roomNumPoint, new Point2d(x - 26.16, y + 7.69), 0, 0, 0);
                    Poly58.AddVertexAt(++roomNumPoint, new Point2d(x - 28.51, y + 7.69), 0, 0, 0);
                    Poly58.AddVertexAt(++roomNumPoint, new Point2d(x - 28.51, y + 7.41), 0, 0, 0);
                    Poly58.Closed = true;
                    Poly58.LineWeight = 0;
                    Poly58.ConstantWidth = 0;
                    Poly58.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly58);
                    acTrans.AddNewlyCreatedDBObject(Poly58, true);

                    ObjectIdCollection acObjIdColl58 = new ObjectIdCollection();
                    acObjIdColl58.Add(Poly58.ObjectId);

                    //////////////////////////////////////////////////////////////////////////////////////////
                    // Y
                    roomNumPoint = 0;
                    Polyline Poly59 = new Polyline();
                    Poly59.SetDatabaseDefaults();
                    Poly59.AddVertexAt(roomNumPoint, new Point2d(x - 23.77, y + 5.49), 0, 0, 0);
                    Poly59.AddVertexAt(++roomNumPoint, new Point2d(x - 24.27, y + 4.55), 0, 0, 0);
                    Poly59.AddVertexAt(++roomNumPoint, new Point2d(x - 23.96, y + 4.55), 0, 0, 0);
                    Poly59.AddVertexAt(++roomNumPoint, new Point2d(x - 22.32, y + 7.69), 0, 0, 0);
                    Poly59.AddVertexAt(++roomNumPoint, new Point2d(x - 22.64, y + 7.69), 0, 0, 0);
                    Poly59.AddVertexAt(++roomNumPoint, new Point2d(x - 23.55, y + 5.93), 0, 0, 0);
                    Poly59.AddVertexAt(++roomNumPoint, new Point2d(x - 24.81, y + 7.69), 0, 0, 0);
                    Poly59.AddVertexAt(++roomNumPoint, new Point2d(x - 25.28, y + 7.69), 0, 0, 0);
                    Poly59.Closed = true;
                    Poly59.LineWeight = 0;
                    Poly59.ConstantWidth = 0;
                    Poly59.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly59);
                    acTrans.AddNewlyCreatedDBObject(Poly59, true);

                    ObjectIdCollection acObjIdColl59 = new ObjectIdCollection();
                    acObjIdColl59.Add(Poly59.ObjectId);

                    //////////////////////////////////////////////////////////////////////////////////////////
                    // Д
                    roomNumPoint = 0;
                    Polyline Poly60 = new Polyline();
                    Poly60.SetDatabaseDefaults();
                    Poly60.AddVertexAt(roomNumPoint, new Point2d(x - 22.01, y + 4.74), 0, 0, 0);
                    Poly60.AddVertexAt(++roomNumPoint, new Point2d(x - 22.14, y + 4.74), 0, 0, 0);
                    Poly60.AddVertexAt(++roomNumPoint, new Point2d(x - 22.14, y + 4.36), 0, 0, 0);
                    Poly60.AddVertexAt(++roomNumPoint, new Point2d(x - 21.98, y + 4.39), 0, 0, 0);
                    Poly60.AddVertexAt(++roomNumPoint, new Point2d(x - 21.88, y + 4.46), 0, 0, 0);
                    Poly60.AddVertexAt(++roomNumPoint, new Point2d(x - 21.70, y + 4.52), 0, 0, 0);
                    Poly60.AddVertexAt(++roomNumPoint, new Point2d(x - 21.41, y + 4.55), 0, 0, 0);
                    Poly60.AddVertexAt(++roomNumPoint, new Point2d(x - 19.72, y + 4.55), 0, 0, 0);
                    Poly60.AddVertexAt(++roomNumPoint, new Point2d(x - 19.43, y + 4.52), 0, 0, 0);
                    Poly60.AddVertexAt(++roomNumPoint, new Point2d(x - 19.28, y + 4.46), 0, 0, 0);
                    Poly60.AddVertexAt(++roomNumPoint, new Point2d(x - 19.18, y + 4.39), 0, 0, 0);
                    Poly60.AddVertexAt(++roomNumPoint, new Point2d(x - 19.06, y + 4.36), 0, 0, 0);
                    Poly60.AddVertexAt(++roomNumPoint, new Point2d(x - 19.06, y + 4.74), 0, 0, 0);
                    Poly60.AddVertexAt(++roomNumPoint, new Point2d(x - 19.18, y + 4.74), 0, 0, 0);
                    Poly60.AddVertexAt(++roomNumPoint, new Point2d(x - 19.31, y + 4.86), 0, 0, 0);
                    Poly60.AddVertexAt(++roomNumPoint, new Point2d(x - 20.41, y + 7.56), 0, 0, 0);
                    Poly60.AddVertexAt(++roomNumPoint, new Point2d(x - 20.53, y + 7.69), 0, 0, 0);
                    Poly60.AddVertexAt(++roomNumPoint, new Point2d(x - 20.63, y + 7.69), 0, 0, 0);
                    Poly60.AddVertexAt(++roomNumPoint, new Point2d(x - 20.91, y + 7.22), 0, 0, 0);
                    Poly60.AddVertexAt(++roomNumPoint, new Point2d(x - 21.82, y + 5.02), 0, 0, 0);
                    Poly60.AddVertexAt(++roomNumPoint, new Point2d(x - 21.88, y + 4.86), 0, 0, 0);
                    Poly60.Closed = true;
                    Poly60.LineWeight = 0;
                    Poly60.ConstantWidth = 0;
                    Poly60.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly60);
                    acTrans.AddNewlyCreatedDBObject(Poly60, true);
                    ObjectIdCollection acObjIdColl60 = new ObjectIdCollection();
                    acObjIdColl60.Add(Poly60.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly61 = new Polyline();
                    Poly61.SetDatabaseDefaults();
                    Poly61.AddVertexAt(roomNumPoint, new Point2d(x - 20.69, y + 7.16), 0, 0, 0);
                    Poly61.AddVertexAt(++roomNumPoint, new Point2d(x - 19.72, y + 4.74), 0, 0, 0);
                    Poly61.AddVertexAt(++roomNumPoint, new Point2d(x - 21.63, y + 4.74), 0, 0, 0);
                    Poly61.AddVertexAt(++roomNumPoint, new Point2d(x - 21.66, y + 4.77), 0, 0, 0);
                    Poly61.Closed = true;
                    Poly61.LineWeight = 0;
                    Poly61.ConstantWidth = 0;
                    Poly61.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly61);
                    acTrans.AddNewlyCreatedDBObject(Poly61, true);
                    ObjectIdCollection acObjIdColl61 = new ObjectIdCollection();
                    acObjIdColl61.Add(Poly61.ObjectId);

                    //////////////////////////////////////////////////////////////////////////////////////////
                    // И
                    roomNumPoint = 0;
                    Polyline Poly62 = new Polyline();
                    Poly62.SetDatabaseDefaults();
                    Poly62.AddVertexAt(roomNumPoint, new Point2d(x - 15.70, y + 7.06), 0, 0, 0);
                    Poly62.AddVertexAt(++roomNumPoint, new Point2d(x - 15.70, y + 4.61), 0, 0, 0);
                    Poly62.AddVertexAt(++roomNumPoint, new Point2d(x - 15.32, y + 4.61), 0, 0, 0);
                    Poly62.AddVertexAt(++roomNumPoint, new Point2d(x - 15.32, y + 7.69), 0, 0, 0);
                    Poly62.AddVertexAt(++roomNumPoint, new Point2d(x - 15.54, y + 7.69), 0, 0, 0);
                    Poly62.AddVertexAt(++roomNumPoint, new Point2d(x - 17.62, y + 5.34), 0, 0, 0);
                    Poly62.AddVertexAt(++roomNumPoint, new Point2d(x - 17.64, y + 5.37), 0, 0, 0);
                    Poly62.AddVertexAt(++roomNumPoint, new Point2d(x - 17.64, y + 7.69), 0, 0, 0);
                    Poly62.AddVertexAt(++roomNumPoint, new Point2d(x - 18.05, y + 7.69), 0, 0, 0);
                    Poly62.AddVertexAt(++roomNumPoint, new Point2d(x - 18.05, y + 4.61), 0, 0, 0);
                    Poly62.AddVertexAt(++roomNumPoint, new Point2d(x - 17.96, y + 4.61), 0, 0, 0);
                    Poly62.Closed = true;
                    Poly62.LineWeight = 0;
                    Poly62.ConstantWidth = 0;
                    Poly62.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly62);
                    acTrans.AddNewlyCreatedDBObject(Poly62, true);
                    ObjectIdCollection acObjIdColl62 = new ObjectIdCollection();
                    acObjIdColl62.Add(Poly62.ObjectId);

                    //////////////////////////////////////////////////////////////////////////////////////////
                    // Я
                    roomNumPoint = 0;
                    Polyline Poly63 = new Polyline();
                    Poly63.SetDatabaseDefaults();
                    Poly63.AddVertexAt(roomNumPoint, new Point2d(x - 13.94, y + 7.07), 0, 0, 0);
                    Poly63.AddVertexAt(++roomNumPoint, new Point2d(x - 13.94, y + 6.87), 0, 0, 0);
                    Poly63.AddVertexAt(++roomNumPoint, new Point2d(x - 13.84, y + 6.59), 0, 0, 0);
                    Poly63.AddVertexAt(++roomNumPoint, new Point2d(x - 13.62, y + 6.34), 0, 0, 0);
                    Poly63.AddVertexAt(++roomNumPoint, new Point2d(x - 13.31, y + 6.18), 0, 0, 0);
                    Poly63.AddVertexAt(++roomNumPoint, new Point2d(x - 13.00, y + 6.09), 0, 0, 0);
                    Poly63.AddVertexAt(++roomNumPoint, new Point2d(x - 14.22, y + 4.55), 0, 0, 0);
                    Poly63.AddVertexAt(++roomNumPoint, new Point2d(x - 13.78, y + 4.55), 0, 0, 0);
                    Poly63.AddVertexAt(++roomNumPoint, new Point2d(x - 12.56, y + 5.99), 0, 0, 0);
                    Poly63.AddVertexAt(++roomNumPoint, new Point2d(x - 12.46, y + 6.06), 0, 0, 0);
                    Poly63.AddVertexAt(++roomNumPoint, new Point2d(x - 12.34, y + 6.06), 0, 0, 0);
                    Poly63.AddVertexAt(++roomNumPoint, new Point2d(x - 12.34, y + 4.55), 0, 0, 0);
                    Poly63.AddVertexAt(++roomNumPoint, new Point2d(x - 11.96, y + 4.55), 0, 0, 0);
                    Poly63.AddVertexAt(++roomNumPoint, new Point2d(x - 11.96, y + 7.69), 0, 0, 0);
                    Poly63.AddVertexAt(++roomNumPoint, new Point2d(x - 12.53, y + 7.69), 0, 0, 0);
                    Poly63.AddVertexAt(++roomNumPoint, new Point2d(x - 12.97, y + 7.66), 0, 0, 0);
                    Poly63.AddVertexAt(++roomNumPoint, new Point2d(x - 13.32, y + 7.61), 0, 0, 0);
                    Poly63.AddVertexAt(++roomNumPoint, new Point2d(x - 13.62, y + 7.53), 0, 0, 0);
                    Poly63.AddVertexAt(++roomNumPoint, new Point2d(x - 13.84, y + 7.31), 0, 0, 0);
                    Poly63.Closed = true;
                    Poly63.LineWeight = 0;
                    Poly63.ConstantWidth = 0;
                    Poly63.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly63);
                    acTrans.AddNewlyCreatedDBObject(Poly63, true);
                    ObjectIdCollection acObjIdColl63 = new ObjectIdCollection();
                    acObjIdColl63.Add(Poly63.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly64 = new Polyline();
                    Poly64.SetDatabaseDefaults();
                    Poly64.AddVertexAt(roomNumPoint, new Point2d(x - 13.53, y + 6.89), 0, 0, 0);
                    Poly64.AddVertexAt(++roomNumPoint, new Point2d(x - 13.44, y + 7.24), 0, 0, 0);
                    Poly64.AddVertexAt(++roomNumPoint, new Point2d(x - 13.22, y + 7.43), 0, 0, 0);
                    Poly64.AddVertexAt(++roomNumPoint, new Point2d(x - 12.87, y + 7.49), 0, 0, 0);
                    Poly64.AddVertexAt(++roomNumPoint, new Point2d(x - 12.46, y + 7.52), 0, 0, 0);
                    Poly64.AddVertexAt(++roomNumPoint, new Point2d(x - 12.34, y + 7.52), 0, 0, 0);
                    Poly64.AddVertexAt(++roomNumPoint, new Point2d(x - 12.34, y + 6.20), 0, 0, 0);
                    Poly64.AddVertexAt(++roomNumPoint, new Point2d(x - 12.75, y + 6.20), 0, 0, 0);
                    Poly64.AddVertexAt(++roomNumPoint, new Point2d(x - 13.03, y + 6.27), 0, 0, 0);
                    Poly64.AddVertexAt(++roomNumPoint, new Point2d(x - 13.28, y + 6.39), 0, 0, 0);
                    Poly64.AddVertexAt(++roomNumPoint, new Point2d(x - 13.47, y + 6.61), 0, 0, 0);
                    Poly64.Closed = true;
                    Poly64.LineWeight = 0;
                    Poly64.ConstantWidth = 0;
                    Poly64.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly64);
                    acTrans.AddNewlyCreatedDBObject(Poly64, true);
                    ObjectIdCollection acObjIdColl64 = new ObjectIdCollection();
                    acObjIdColl64.Add(Poly64.ObjectId);

                    //////////////////////////////////////////////////////////////////////////////////////////
                    // 44
                    roomNumPoint = 0;
                    Polyline Poly65 = new Polyline();
                    Poly65.SetDatabaseDefaults();
                    Poly65.AddVertexAt(roomNumPoint, new Point2d(x - 8.88, y + 5.49), 0, 0, 0);
                    Poly65.AddVertexAt(++roomNumPoint, new Point2d(x - 7.44, y + 5.49), 0, 0, 0);
                    Poly65.AddVertexAt(++roomNumPoint, new Point2d(x - 7.44, y + 4.64), 0, 0, 0);
                    Poly65.AddVertexAt(++roomNumPoint, new Point2d(x - 7.09, y + 4.64), 0, 0, 0);
                    Poly65.AddVertexAt(++roomNumPoint, new Point2d(x - 7.09, y + 5.49), 0, 0, 0);
                    Poly65.AddVertexAt(++roomNumPoint, new Point2d(x - 6.68, y + 5.49), 0, 0, 0);
                    Poly65.AddVertexAt(++roomNumPoint, new Point2d(x - 6.68, y + 5.74), 0, 0, 0);
                    Poly65.AddVertexAt(++roomNumPoint, new Point2d(x - 7.09, y + 5.74), 0, 0, 0);
                    Poly65.AddVertexAt(++roomNumPoint, new Point2d(x - 7.09, y + 7.69), 0, 0, 0);
                    Poly65.AddVertexAt(++roomNumPoint, new Point2d(x - 7.41, y + 7.69), 0, 0, 0);
                    Poly65.Closed = true;
                    Poly65.LineWeight = 0;
                    Poly65.ConstantWidth = 0;
                    Poly65.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly65);
                    acTrans.AddNewlyCreatedDBObject(Poly65, true);
                    ObjectIdCollection acObjIdColl65 = new ObjectIdCollection();
                    acObjIdColl65.Add(Poly65.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly66 = new Polyline();
                    Poly66.SetDatabaseDefaults();
                    Poly66.AddVertexAt(roomNumPoint, new Point2d(x - 7.47, y + 7.22), 0, 0, 0);
                    Poly66.AddVertexAt(++roomNumPoint, new Point2d(x - 7.44, y + 7.22), 0, 0, 0);
                    Poly66.AddVertexAt(++roomNumPoint, new Point2d(x - 7.44, y + 5.74), 0, 0, 0);
                    Poly66.AddVertexAt(++roomNumPoint, new Point2d(x - 8.47, y + 5.74), 0, 0, 0);
                    Poly66.Closed = true;
                    Poly66.LineWeight = 0;
                    Poly66.ConstantWidth = 0;
                    Poly66.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly66);
                    acTrans.AddNewlyCreatedDBObject(Poly66, true);
                    ObjectIdCollection acObjIdColl66 = new ObjectIdCollection();
                    acObjIdColl66.Add(Poly66.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly67 = new Polyline();
                    Poly67.SetDatabaseDefaults();
                    Poly67.AddVertexAt(roomNumPoint, new Point2d(x - 8.88 + 3.2, y + 5.49), 0, 0, 0);
                    Poly67.AddVertexAt(++roomNumPoint, new Point2d(x - 7.44 + 3.2, y + 5.49), 0, 0, 0);
                    Poly67.AddVertexAt(++roomNumPoint, new Point2d(x - 7.44 + 3.2, y + 4.64), 0, 0, 0);
                    Poly67.AddVertexAt(++roomNumPoint, new Point2d(x - 7.09 + 3.2, y + 4.64), 0, 0, 0);
                    Poly67.AddVertexAt(++roomNumPoint, new Point2d(x - 7.09 + 3.2, y + 5.49), 0, 0, 0);
                    Poly67.AddVertexAt(++roomNumPoint, new Point2d(x - 6.68 + 3.2, y + 5.49), 0, 0, 0);
                    Poly67.AddVertexAt(++roomNumPoint, new Point2d(x - 6.68 + 3.2, y + 5.74), 0, 0, 0);
                    Poly67.AddVertexAt(++roomNumPoint, new Point2d(x - 7.09 + 3.2, y + 5.74), 0, 0, 0);
                    Poly67.AddVertexAt(++roomNumPoint, new Point2d(x - 7.09 + 3.2, y + 7.69), 0, 0, 0);
                    Poly67.AddVertexAt(++roomNumPoint, new Point2d(x - 7.41 + 3.2, y + 7.69), 0, 0, 0);
                    Poly67.Closed = true;
                    Poly67.LineWeight = 0;
                    Poly67.ConstantWidth = 0;
                    Poly67.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly67);
                    acTrans.AddNewlyCreatedDBObject(Poly67, true);
                    ObjectIdCollection acObjIdColl67 = new ObjectIdCollection();
                    acObjIdColl67.Add(Poly67.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly68 = new Polyline();
                    Poly68.SetDatabaseDefaults();
                    Poly68.AddVertexAt(roomNumPoint, new Point2d(x - 7.47 + 3.2, y + 7.22), 0, 0, 0);
                    Poly68.AddVertexAt(++roomNumPoint, new Point2d(x - 7.44 + 3.2, y + 7.22), 0, 0, 0);
                    Poly68.AddVertexAt(++roomNumPoint, new Point2d(x - 7.44 + 3.2, y + 5.74), 0, 0, 0);
                    Poly68.AddVertexAt(++roomNumPoint, new Point2d(x - 8.47 + 3.2, y + 5.74), 0, 0, 0);
                    Poly68.Closed = true;
                    Poly68.LineWeight = 0;
                    Poly68.ConstantWidth = 0;
                    Poly68.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly68);
                    acTrans.AddNewlyCreatedDBObject(Poly68, true);
                    ObjectIdCollection acObjIdColl68 = new ObjectIdCollection();
                    acObjIdColl68.Add(Poly68.ObjectId);

                    //////////////////////////////////////////////////////////////////////////////////////////
                    // Кавычки большие в конце "
                    roomNumPoint = 0;
                    Polyline Poly69 = new Polyline();
                    Poly69.SetDatabaseDefaults();
                    Poly69.AddVertexAt(roomNumPoint, new Point2d(x - 2.48, y + 6.39), 0, 0, 0);
                    Poly69.AddVertexAt(++roomNumPoint, new Point2d(x - 2.32, y + 6.57), 0, 0, 0);
                    Poly69.AddVertexAt(++roomNumPoint, new Point2d(x - 2.10, y + 6.95), 0, 0, 0);
                    Poly69.AddVertexAt(++roomNumPoint, new Point2d(x - 1.88, y + 7.33), 0, 0, 0);
                    Poly69.AddVertexAt(++roomNumPoint, new Point2d(x - 1.82, y + 7.52), 0, 0, 0);
                    Poly69.AddVertexAt(++roomNumPoint, new Point2d(x - 1.82, y + 7.60), 0, 0, 0);
                    Poly69.AddVertexAt(++roomNumPoint, new Point2d(x - 1.88, y + 7.67), 0, 0, 0);
                    Poly69.AddVertexAt(++roomNumPoint, new Point2d(x - 1.97, y + 7.69), 0, 0, 0);
                    Poly69.AddVertexAt(++roomNumPoint, new Point2d(x - 2.07, y + 7.67), 0, 0, 0);
                    Poly69.AddVertexAt(++roomNumPoint, new Point2d(x - 2.16, y + 7.52), 0, 0, 0);
                    Poly69.AddVertexAt(++roomNumPoint, new Point2d(x - 2.35, y + 7.11), 0, 0, 0);
                    Poly69.AddVertexAt(++roomNumPoint, new Point2d(x - 2.48, y + 6.64), 0, 0, 0);
                    Poly69.AddVertexAt(++roomNumPoint, new Point2d(x - 2.49, y + 6.52), 0, 0, 0);
                    Poly69.Closed = true;
                    Poly69.LineWeight = 0;
                    Poly69.ConstantWidth = 0;
                    Poly69.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly69);
                    acTrans.AddNewlyCreatedDBObject(Poly69, true);
                    ObjectIdCollection acObjIdColl69 = new ObjectIdCollection();
                    acObjIdColl69.Add(Poly69.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly70 = new Polyline();
                    Poly70.SetDatabaseDefaults();
                    Poly70.AddVertexAt(roomNumPoint, new Point2d(x - 2.48 + 0.77, y + 6.39), 0, 0, 0);
                    Poly70.AddVertexAt(++roomNumPoint, new Point2d(x - 2.32 + 0.77, y + 6.57), 0, 0, 0);
                    Poly70.AddVertexAt(++roomNumPoint, new Point2d(x - 2.10 + 0.77, y + 6.95), 0, 0, 0);
                    Poly70.AddVertexAt(++roomNumPoint, new Point2d(x - 1.88 + 0.77, y + 7.33), 0, 0, 0);
                    Poly70.AddVertexAt(++roomNumPoint, new Point2d(x - 1.82 + 0.77, y + 7.52), 0, 0, 0);
                    Poly70.AddVertexAt(++roomNumPoint, new Point2d(x - 1.82 + 0.77, y + 7.60), 0, 0, 0);
                    Poly70.AddVertexAt(++roomNumPoint, new Point2d(x - 1.88 + 0.77, y + 7.67), 0, 0, 0);
                    Poly70.AddVertexAt(++roomNumPoint, new Point2d(x - 1.97 + 0.77, y + 7.69), 0, 0, 0);
                    Poly70.AddVertexAt(++roomNumPoint, new Point2d(x - 2.07 + 0.77, y + 7.67), 0, 0, 0);
                    Poly70.AddVertexAt(++roomNumPoint, new Point2d(x - 2.16 + 0.77, y + 7.52), 0, 0, 0);
                    Poly70.AddVertexAt(++roomNumPoint, new Point2d(x - 2.35 + 0.77, y + 7.11), 0, 0, 0);
                    Poly70.AddVertexAt(++roomNumPoint, new Point2d(x - 2.48 + 0.77, y + 6.64), 0, 0, 0);
                    Poly70.AddVertexAt(++roomNumPoint, new Point2d(x - 2.49 + 0.77, y + 6.52), 0, 0, 0);
                    Poly70.Closed = true;
                    Poly70.LineWeight = 0;
                    Poly70.ConstantWidth = 0;
                    Poly70.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly70);
                    acTrans.AddNewlyCreatedDBObject(Poly70, true);
                    ObjectIdCollection acObjIdColl70 = new ObjectIdCollection();
                    acObjIdColl70.Add(Poly70.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////

                    Hatch acHatch = new Hatch();
                    acBlkTblRec.AppendEntity(acHatch);

                    acTrans.AddNewlyCreatedDBObject(acHatch, true);

                    acHatch.SetDatabaseDefaults();
                    acHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                    acHatch.Associative = true;

                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl1);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl2);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl3);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl4);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl5);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl6);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl7);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl8);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl9);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl10);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl_10);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl11);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl12);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl13);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl14);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl15);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl16);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl17);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl18);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl19);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl20);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl21);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl22);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl23);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl24);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl25);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl26);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl27);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl28);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl29);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl30);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl31);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl32);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl33);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl34);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl35);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl36);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl37);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl38);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl39);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl40);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl41);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl42);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl43);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl44);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl45);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl46);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl47);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl48);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl49);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl50);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl51);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl52);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl53);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl54);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl55);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl56);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl57);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl58);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl59);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl60);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl61);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl62);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl63);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl64);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl65);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl66);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl67);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl68);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl69);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl70);

                    acHatch.EvaluateHatch(true);
                    acHatch.Layer = "Z-STMP";
                    acTrans.Commit();
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n [AddLogo] Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        //Регистрация приложения в системе, не придётся постоянно вызывать Netload для загрузки dll.

        [CommandMethod("RegisterTDMSApp")]
        public void RegisterTdmsApp()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                // Из реестра получаем ключ AutoCAD
                string sProdKey = Autodesk.AutoCAD.DatabaseServices.HostApplicationServices.Current.UserRegistryProductRootKey;
                string sAppName = "MyApp";
                Microsoft.Win32.RegistryKey regAcadProdKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(sProdKey);
                Microsoft.Win32.RegistryKey regAcadAppKey = regAcadProdKey.OpenSubKey("Applications", true);
                // Проверяем, существует ли в ветке реестра раздел "MyApp"
                string[] subKeys = regAcadAppKey.GetSubKeyNames();
                foreach (string subKey in subKeys)
                {
                    // Если приложение уже зарегистрировано - завершаем работу
                    if (subKey.Equals(sAppName))
                    {
                        regAcadAppKey.Close();
                        return;
                    }
                }
                // Получаем строковое представление пути к каталогу, в котором находится текущий модуль
                string sAssemblyPath = Assembly.GetExecutingAssembly().Location;
                // Регистрируем приложение
                Microsoft.Win32.RegistryKey regAppAddInKey = regAcadAppKey.CreateSubKey(sAppName);
                regAppAddInKey.SetValue("DESCRIPTION", sAppName, RegistryValueKind.String);
                regAppAddInKey.SetValue("LOADCTRLS", 2, RegistryValueKind.DWord);
                regAppAddInKey.SetValue("LOADER", sAssemblyPath, RegistryValueKind.String);
                regAppAddInKey.SetValue("MANAGED", 1, RegistryValueKind.DWord);
                regAcadAppKey.Close();
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        ////Очистка регистров в системе, вызвать Netload для загрузки dll.
        [CommandMethod("UnregisterTDMSApp")]
        public void UnregisterTDMSApp()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                // Из реестра получаем ключ AutoCAD
                string sProdKey = Autodesk.AutoCAD.DatabaseServices.HostApplicationServices.Current.UserRegistryProductRootKey;
                string sAppName = "MyApp";

                RegistryKey regAcadProdKey = Registry.CurrentUser.OpenSubKey(sProdKey);
                RegistryKey regAcadAppKey = regAcadProdKey.OpenSubKey("Applications", true);
                //// Удаляем ключ приложения
                regAcadAppKey.DeleteSubKeyTree(sAppName);
                regAcadAppKey.Close();
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /////////////////////////////////////////////////////////////////////////////////

        public static string GetUserName()
        {
            System.Security.Principal.WindowsIdentity win = null;
            win = System.Security.Principal.WindowsIdentity.GetCurrent();
            return win.Name.Substring(win.Name.IndexOf("\\") + 1);
        }

        [CommandMethod("PACKAGE", CommandFlags.NoBlockEditor)]
        public static void Etrans()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            CurrentVersion cvPath = new CurrentVersion();
            SaveOptions Save = new SaveOptions();

            try
            {
                Save.SaveActiveDrawing();
                Document doc = Application.DocumentManager.MdiActiveDocument;

                Microsoft.Win32.Registry.CurrentUser.CreateSubKey(cvPath.pathEtransmit());
                Microsoft.Win32.RegistryKey myKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(cvPath.pathEtransmit(), true);

                myKey.SetValue("AEC_EXPLODE_DWG", "1", RegistryValueKind.DWord);
                myKey.SetValue("AEC_INCLUDE_DWTS", "0", RegistryValueKind.DWord);
                myKey.SetValue("AEC_INCLUDE_PROJ_INFO", "0", RegistryValueKind.DWord);
                myKey.SetValue("AEC_TEMP_SETUP", "0", RegistryValueKind.DWord);

                myKey.SetValue("BindType", "1", RegistryValueKind.DWord);
                myKey.SetValue("BindXref", "1", RegistryValueKind.DWord);
                myKey.SetValue("Description", "", RegistryValueKind.String);
                myKey.SetValue("DestFile", "", RegistryValueKind.String);
                myKey.SetValue("DestFileAction", "2", RegistryValueKind.DWord);
                myKey.SetValue("DestFolder", "D:\\ETRANSMIT", RegistryValueKind.String);
                myKey.SetValue("FilePathOption", "1", RegistryValueKind.DWord);

                myKey.SetValue("IncludeDataLinkFile", "0", RegistryValueKind.DWord);
                myKey.SetValue("IncludeFont", "0", RegistryValueKind.DWord);
                myKey.SetValue("IncludeMaterialTextures", "0", RegistryValueKind.DWord);
                myKey.SetValue("IncludeFotometricWebFile", "0", RegistryValueKind.DWord);
                myKey.SetValue("IncludeSSFiles", "1", RegistryValueKind.DWord);
                myKey.SetValue("IncludeUnloadedReferences", "0", RegistryValueKind.DWord);
                myKey.SetValue("Name", "Enterprise", RegistryValueKind.String);
                myKey.SetValue("PackageType", "1", RegistryValueKind.DWord);
                myKey.SetValue("PurgeDatabase", "1", RegistryValueKind.DWord);
                myKey.SetValue("RootFolder", "", RegistryValueKind.String);
                myKey.SetValue("SaveDrawingFormat", "4", RegistryValueKind.DWord);
                myKey.SetValue("SendMail", "0", RegistryValueKind.DWord);
                myKey.SetValue("SetPlotterNone", "1", RegistryValueKind.DWord);
                myKey.SetValue("UsePassword", "0", RegistryValueKind.DWord);
                myKey.SetValue("VisualFidelity", "1", RegistryValueKind.DWord);

                doc.SendStringToExecute("._-ETRANSMIT" + "\n" + "CHoose setup" + "\n" + "Enterprise" + "\n" + "Create" + "\n", true, false, false);

                doc.CloseAndDiscard();
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }
    }
}