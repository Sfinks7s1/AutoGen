namespace Auto
{
    using Autodesk.AutoCAD.ApplicationServices.Core;
    using Autodesk.AutoCAD.Runtime;

    /// <summary>
    /// Класс содержит методы, предоставляющие доступ к справочной информации
    /// </summary>
    public sealed class Help
    {
        public void Initialize()
        { }

        public void Terminate()
        { }

        /// <summary>
        /// Метод HelpCommands предназначен для вывода справочной информации о дополнительных командах в AutoCAD
        /// </summary>
        [CommandMethod("HelpCommands", CommandFlags.NoBlockEditor)]
        public void HelpCommands()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                editor.WriteMessage("_This plugin to developed by Lev Kozhaev aka Sfinks7s1 in 2014._" + "\n");
                editor.WriteMessage("\n");
                //Функции сохранения
                editor.WriteMessage("---Функции сохранения" + "\n");
                editor.WriteMessage("_SaveActiveDrawing" + "\n");
                editor.WriteMessage("_SaveAndClose" + "\n");
                editor.WriteMessage("_CloseAndDiscard" + "\n");
                editor.WriteMessage("_RelativePath" + "\n");
                editor.WriteMessage("_ExportXrefFromTDMS" + "\n");
                editor.WriteMessage("_Package" + "\n");

                editor.WriteMessage("\n");
                //Внешние ссылки из TDMS
                editor.WriteMessage("---Внешние ссылки из TDMS" + "\n");
                editor.WriteMessage("_TDMSXREFADD" + "\n");
                editor.WriteMessage("_UpdateXrefAttr" + "\n");
                editor.WriteMessage("_AddExternalSheduleTable" + "\n");
                editor.WriteMessage("_UpdateTable" + "\n");

                editor.WriteMessage("\n");
                //Рамки по формату
                editor.WriteMessage("---Рамки по формату" + "\n");
                editor.WriteMessage("_A0H" + "\n");
                editor.WriteMessage("_A0V" + "\n");
                editor.WriteMessage("_A1HL" + "\n");
                editor.WriteMessage("_A1VL" + "\n");
                editor.WriteMessage("_A1H" + "\n");
                editor.WriteMessage("_A1V" + "\n");
                editor.WriteMessage("_A2H" + "\n");
                editor.WriteMessage("_A2V" + "\n");
                editor.WriteMessage("_A3H" + "\n");
                editor.WriteMessage("_A3V" + "\n");
                editor.WriteMessage("_A3H3" + "\n");
                editor.WriteMessage("_A4H" + "\n");
                editor.WriteMessage("_A4V" + "\n");
                editor.WriteMessage("_A4H3" + "\n");
                editor.WriteMessage("_A4H4" + "\n");
                editor.WriteMessage("_FrameCustom" + "\n");

                editor.WriteMessage("\n");
                //Рамки для оформления
                editor.WriteMessage("---Рамки для оформления" + "\n");
                editor.WriteMessage("_A4HA" + "\n");
                editor.WriteMessage("_A4VA" + "\n");
                editor.WriteMessage("_A3HA" + "\n");
                editor.WriteMessage("_A3VA" + "\n");
                editor.WriteMessage("_A2HA" + "\n");
                editor.WriteMessage("_A2VA" + "\n");
                editor.WriteMessage("_A1HA" + "\n");
                editor.WriteMessage("_A1VA" + "\n");
                editor.WriteMessage("_A0HA" + "\n");
                editor.WriteMessage("_A0VA" + "\n");

                editor.WriteMessage("\n");
                //Штампы
                editor.WriteMessage("---Штампы" + "\n");
                editor.WriteMessage("_StampAR" + "\n");
                editor.WriteMessage("_StampK" + "\n");
                editor.WriteMessage("_StampKAB" + "\n");

                editor.WriteMessage("\n");
                //Слои
                editor.WriteMessage("---Слои" + "\n");
                editor.WriteMessage("_LDLNP" + "\n");
                editor.WriteMessage("_LDLF" + "\n");
                editor.WriteMessage("_LDLO" + "\n");
                editor.WriteMessage("_LDLC" + "\n");
                editor.WriteMessage("_LayerArchitect" + "\n");
                editor.WriteMessage("_LayerConstruct" + "\n");
                editor.WriteMessage("_LayerRestorer" + "\n");
                editor.WriteMessage("_LayerRestorerElements" + "\n");
                editor.WriteMessage("_LayerRestorerHatch" + "\n");
                editor.WriteMessage("_LayerRestorerFurniture" + "\n");
                editor.WriteMessage("_LayerRestorerFundament" + "\n");
                editor.WriteMessage("_LayerRestorerFasads" + "\n");
                editor.WriteMessage("_LayerRestorerNode" + "\n");
                editor.WriteMessage("_LayerRestorerSize" + "\n");
                editor.WriteMessage("_LayerRestorerLayerPlace" + "\n");
                editor.WriteMessage("_LayerRestorerFloor" + "\n");
                editor.WriteMessage("_LayerRestorerWindow" + "\n");
                editor.WriteMessage("_LayerRestorerStairs" + "\n");
                editor.WriteMessage("_LayerRestorerDoors" + "\n");

                editor.WriteMessage("\n");
                //Листы
                editor.WriteMessage("---Листы" + "\n");
                editor.WriteMessage("_DivideLayout" + "\n");

                editor.WriteMessage("\n");
                //Помощь
                editor.WriteMessage("---Помощь" + "\n");
                editor.WriteMessage("_OpenHelp" + "\n");
                editor.WriteMessage("_OpenStandart" + "\n");
                editor.WriteMessage("_Questions" + "\n");

                editor.WriteMessage("\n");
                //Системные команды
                editor.WriteMessage("---Системные команды" + "\n");
                editor.WriteMessage("_RegisterTDMSApp" + "\n");
                editor.WriteMessage("_UnregisterTDMSApp" + "\n");
                editor.WriteMessage("_MyRibbon" + "\n");
                editor.WriteMessage("_HelpCommands" + "\n");
                editor.WriteMessage("_FrameCustom" + "\n");
                editor.WriteMessage("_СustomVar" + "\n");
                editor.WriteMessage("_UPurge" + "\n");
                editor.WriteMessage("\n");
                editor.WriteMessage("Все команды TDMS выведены, для просмотра нажмите кнопку F2 на клавиатуре." + "\n");

#if _DEBUG_Help
                editor.WriteMessage("\n ---------------- START DEBUG INFO ----------------");
                editor.WriteMessage("\n ----------------   HelpCommands   ----------------");

                editor.WriteMessage("\n All commands displayed is correct");

                editor.WriteMessage("\n ---------------- STOP DEBUG INFO -----------------");
#endif
            }
            catch (Exception ex)
            {
                editor.WriteMessage("\n Exception caught: \n" + ex);
            }
        }

        /// <summary>
        /// Метод OpenHelp открывает каталог HELP, расположенный на сервере (V:\_HELP\)
        /// </summary>
        [CommandMethod("OpenHelp", CommandFlags.NoBlockEditor)]
        public static void OpenHelp()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                System.Diagnostics.Process.Start(@"V:\_HELP\");
            }
            catch
            {
                editor.WriteMessage("\n Path not found");
            }
        }

        [CommandMethod("OpenStandart", CommandFlags.NoBlockEditor)]
        public static void OpenStandart()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                System.Diagnostics.Process.Start(@"V:\_HELP\STANDART.pdf");
            }
            catch
            {
                editor.WriteMessage("\n Path not found");
            }
        }

        [CommandMethod("OpenManual", CommandFlags.NoBlockEditor)]
        public static void OpenManual()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                System.Diagnostics.Process.Start(@"D:\_TDMSHELP\TDMSHELP.chm");
            }
            catch
            {
                editor.WriteMessage("\n Path not found");
            }
        }

        [CommandMethod("_Questions", CommandFlags.NoBlockEditor)]
        public static void Questions()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                System.Diagnostics.Process.Start(@"V:\_HELP\Часто задаваемые вопросы.pdf");
            }
            catch
            {
                editor.WriteMessage("\n Path not found");
            }
        }
    }
}