namespace Auto
{
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.EditorInput;
    using Autodesk.AutoCAD.Runtime;

    /// <summary>
    /// Класс предназначен для обновления атрибутов в блоках и базе данных чертежа
    /// </summary>
    public class Commands
    {
        [CommandMethod("UA")]
        public void UpdateAttribute()
        {
            Condition Check = new Condition();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor editor = doc.Editor;

            try
            {
                if (Check.CheckPath())
                {
                    if (Check.CheckTdmsProcess())
                    {
                        Technical Refresh = new Technical();
                        string moduleName = "CMD_SYSLIB";
                        string _GDFOFA = "GETDATAFROMOBJFORACAD";
                        string _GDFPOFA = "GETDATAFROMPAROBJFORACAD";
                        string _GDFARTFA = "GETDATAFROMACTIVEROUTETABLEFORACAD";
                        string _GDFSPFA = "GETDATAFROMSYSPROPSFORACAD";
                        string _GOA = "GETOPADDRESS";
                        string _GON = "GETOPNAME";
                        //Наименования атрибутов
                        string nameObj = "A_DESIGN_OBJ_REF";
                        string oboznach = "A_OBOZN_DOC";
                        string inventNum = "A_ARCH_SIGN";
                        string pageNum = "A_PAGE_NUM";
                        string pageCount = "A_PAGE_COUNT";
                        string tomPage = "A_TOM_PAGE_NUM";

                        //Обновление обычных атрибутов
                        Refresh.RefreshAttribute(nameObj);
                        Refresh.RefreshAttribute(oboznach);
                        Refresh.RefreshAttribute(inventNum);
                        //Refresh.RefreshAttribute(pageNum);
                        Refresh.RefreshAttribute(pageCount);
                        Refresh.RefreshAttribute(tomPage);

                        //Обновление атрибутов функций с разным количеством параметров
                        Refresh.RefreshAttribute(moduleName, _GDFOFA, "A_STAGE_CLSF"); // Стадия
                        Refresh.RefreshAttribute(moduleName, _GDFPOFA, "A_INSTEAD_OF_NUM"); // взамен инв №
                        Refresh.RefreshAttribute(moduleName, _GDFSPFA, "GKAB_"); //Гл.констр.АБ
                        Refresh.RefreshAttribute(moduleName, _GOA);// Адрес
                        Refresh.RefreshAttribute(moduleName, _GON);// Наименование объекта

                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "DEVELOP", "A_User"); // Разработал
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "CHECK", "A_User"); //Проверил
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "NORMKL", "A_User"); //Нормоконтроль
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GR_HEAD", "A_User"); //Руководитель группы
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GIP_", "A_User"); //ГИП
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GAP_", "A_User"); //ГAП
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GKP_", "A_User"); //Главный конструктор проекта

                        //Обновление дат
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "DEVELOP", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "CHECK", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "NORMKL", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GR_HEAD", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GIP_", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GAP_", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GKP_", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GKAB_", "A_DATE");

                        editor.Regen();

#if _DEBUG_UpdateAttribute
                        editor.WriteMessage("\n ---------------- START DEBUG INFO ----------------");
                        editor.WriteMessage("\n ---------------- UpdateAttribute  ----------------");
                        editor.WriteMessage("\n -----moduleName: " + moduleName );
                        editor.WriteMessage("\n -----_GDFOFA:    " + _GDFOFA);
                        editor.WriteMessage("\n -----_GDFPOFA:   " + _GDFPOFA);
                        editor.WriteMessage("\n -----_GDFARTFA:  " + _GDFARTFA);
                        editor.WriteMessage("\n -----_GDFSPFA:   " + _GDFSPFA);
                        editor.WriteMessage("\n -----_GOA:       " + _GOA);
                        editor.WriteMessage("\n -----_GON:       " + _GON);
                        editor.WriteMessage("\n ---------------- Attribute name  -----------------");
                        editor.WriteMessage("\n -----nameObj:    " + nameObj);
                        editor.WriteMessage("\n -----oboznach:   " + oboznach);
                        editor.WriteMessage("\n -----inventNum:  " + inventNum);
                        editor.WriteMessage("\n -----pageNum:    " + pageNum);
                        editor.WriteMessage("\n -----pageCount:  " + pageCount);
                        editor.WriteMessage("\n -----tomPage:    " + tomPage);
                        editor.WriteMessage("\n ----------------- STOP DEBUG INFO ----------------");
#endif
                    }
                    else
                    {
                        Application.ShowAlertDialog("\n Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                    }
                }
                else
                {
                    Application.ShowAlertDialog("Документ не принадлежит TDMS!");
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace);
            }
        }

        public void UpdateAttributeWithoutMsg()
        {
            Condition Check = new Condition();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor editor = doc.Editor;

            try
            {
                if (Check.CheckPath())
                {
                    if (Check.CheckTdmsProcess())
                    {
                        Technical Refresh = new Technical();
                        string moduleName = "CMD_SYSLIB";
                        string _GDFOFA = "GETDATAFROMOBJFORACAD";
                        string _GDFPOFA = "GETDATAFROMPAROBJFORACAD";
                        string _GDFARTFA = "GETDATAFROMACTIVEROUTETABLEFORACAD";
                        string _GDFSPFA = "GETDATAFROMSYSPROPSFORACAD";
                        string _GOA = "GETOPADDRESS";
                        string _GON = "GETOPNAME";

                        //Наименования атрибутов
                        string nameObj = "A_DESIGN_OBJ_REF";
                        string oboznach = "A_OBOZN_DOC";
                        string inventNum = "A_ARCH_SIGN";
                        string pageNum = "A_PAGE_NUM";
                        string pageCount = "A_PAGE_COUNT";
                        string tomPage = "A_TOM_PAGE_NUM";

                        //Обновление обычных атрибутов
                        Refresh.RefreshAttribute(nameObj);
                        Refresh.RefreshAttribute(oboznach);
                        Refresh.RefreshAttribute(inventNum);
                        Refresh.RefreshAttribute(pageNum);
                        Refresh.RefreshAttribute(pageCount);
                        Refresh.RefreshAttribute(tomPage);

                        //Обновление атрибутов функций с разным количеством параметров
                        Refresh.RefreshAttribute(moduleName, _GDFOFA, "A_STAGE_CLSF"); // Стадия
                        Refresh.RefreshAttribute(moduleName, _GDFPOFA, "A_INSTEAD_OF_NUM"); // взамен инв №
                        Refresh.RefreshAttribute(moduleName, _GDFSPFA, "GKAB_"); //Гл.констр.АБ
                        Refresh.RefreshAttribute(moduleName, _GOA);// Адрес
                        Refresh.RefreshAttribute(moduleName, _GON);// Наименование объекта

                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "DEVELOP", "A_User"); // Разработал
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "CHECK", "A_User"); //Проверил
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "NORMKL", "A_User"); //Нормоконтроль
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GR_HEAD", "A_User"); //Руководитель группы
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GIP_", "A_User"); //ГИП
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GAP_", "A_User"); //ГAП
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GKP_", "A_User"); //Главный конструктор проекта

                        //Обновление дат
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "DEVELOP", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "CHECK", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "NORMKL", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GR_HEAD", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GIP_", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GAP_", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GKP_", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GKAB_", "A_DATE");

                        editor.Regen();

#if _DEBUG_UpdateAttributeWithoutMsg
                        editor.WriteMessage("\n ---------------- START DEBUG INFO ----------------");
                        editor.WriteMessage("\n ---------------- UpdateAttribute  ----------------");
                        editor.WriteMessage("\n -----moduleName: " + moduleName);
                        editor.WriteMessage("\n -----_GDFOFA:    " + _GDFOFA);
                        editor.WriteMessage("\n -----_GDFPOFA:   " + _GDFPOFA);
                        editor.WriteMessage("\n -----_GDFARTFA:  " + _GDFARTFA);
                        editor.WriteMessage("\n -----_GDFSPFA:   " + _GDFSPFA);
                        editor.WriteMessage("\n -----_GOA:       " + _GOA);
                        editor.WriteMessage("\n -----_GON:       " + _GON);
                        editor.WriteMessage("\n ---------------- Attribute name  -----------------");
                        editor.WriteMessage("\n -----nameObj:    " + nameObj);
                        editor.WriteMessage("\n -----oboznach:   " + oboznach);
                        editor.WriteMessage("\n -----inventNum:  " + inventNum);
                        editor.WriteMessage("\n -----pageNum:    " + pageNum);
                        editor.WriteMessage("\n -----pageCount:  " + pageCount);
                        editor.WriteMessage("\n -----tomPage:    " + tomPage);
                        editor.WriteMessage("\n ----------------- STOP DEBUG INFO ----------------");
#endif
                    }
                    else
                    {
                        editor.WriteMessage("\n Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                    }
                }
                else
                {
                    editor.WriteMessage("Документ не принадлежит TDMS!");
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace);
            }
        }

        public int UpdateAttributesInDatabase(Database db, string blockName, string attbName, string attbValue)
        {
            //ed.WriteMessage("\n имя блока:         " + blockName);
            //ed.WriteMessage("\n имя атрибута:      " + attbName);
            //ed.WriteMessage("\n значение атрибута: " + attbValue);
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            ObjectId msId, psId;
            Transaction tr = db.TransactionManager.StartTransaction();

            using (tr)
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                //msId = bt[BlockTableRecord.ModelSpace];
                psId = bt[BlockTableRecord.PaperSpace];
                tr.Commit();
            }
            //int msCount = UpdateAttributesInBlock(msId, blockName, attbName, attbValue);
            int psCount = UpdateAttributesInBlock(psId, blockName, attbName, attbValue);

#if _DEBUG_UpdateAttributesInDatabase
            editor.WriteMessage("\n ---------------- START DEBUG INFO ----------------");
            editor.WriteMessage("\n ----------- UpdateAttributesInDatabase -----------");
            editor.WriteMessage("\n -----имя блока:         " + blockName);
            editor.WriteMessage("\n -----имя атрибута:      " + attbName);
            editor.WriteMessage("\n -----значение атрибута: " + attbValue);
            editor.WriteMessage("\n -----psCount:           " + psCount);
            editor.WriteMessage("\n ----------------- STOP DEBUG INFO ----------------");
#endif

            return /*msCount +*/  psCount;
        }

        public int UpdateAttributesInBlock(ObjectId btrId, string blockName, string attbName, string attbValue)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor editor = doc.Editor;
            int changedCount = 0;
            Database db = doc.Database;

#if _DEBUG_UpdateAttributesInBlock
            editor.WriteMessage("\n ---------------- START DEBUG INFO ----------------");
            editor.WriteMessage("\n ------------ UpdateAttributesInBlock -------------");
#endif

            Transaction tr = doc.TransactionManager.StartTransaction();
            using (tr)
            {
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                foreach (ObjectId entId in btr)
                {
                    Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                    if (ent != null)
                    {
                        BlockReference br = ent as BlockReference;
                        if (br != null)
                        {
                            BlockTableRecord bd = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                            foreach (ObjectId arId in br.AttributeCollection)
                            {
                                DBObject obj = tr.GetObject(arId, OpenMode.ForRead);
                                AttributeReference ar = obj as AttributeReference;
                                if (ar != null)
                                {
                                    if (ar.Tag.ToUpper() == attbName.ToUpper())
                                    {
#if _DEBUG_UpdateAttributesInBlock
                                        editor.WriteMessage("\n Имя блока:         " + blockName);
                                        editor.WriteMessage("\n Имя атрибута:      " + attbName);
                                        editor.WriteMessage("\n Значение атрибута: " + attbName);
#endif

                                        ar.UpgradeOpen();
                                        ar.TextString = attbValue;

                                        ar.DowngradeOpen();

                                        changedCount++;

#if _DEBUG_UpdateAttributesInBlock
                                        editor.WriteMessage("\n changedCount:    " + changedCount);
#endif
                                    }
                                }
                            }
                            changedCount += UpdateAttributesInBlock(br.BlockTableRecord, blockName, attbName, attbValue);
                        }
                    }
                }
                tr.Commit();
            }
            return changedCount;
        }
    }
}