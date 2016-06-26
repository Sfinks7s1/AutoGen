namespace Auto
{
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.Runtime;
    using Autodesk.AutoCAD.Windows;
    using System;

    public class ContextMenu : IExtensionApplication
    {
        public void Terminate()
        {
            try
            {
                DefaultContextMenu.RemoveMe();
            }
            catch (System.Exception ex)
            {
            }
        }

        public void Initialize()
        {
            try
            {
                DefaultContextMenu.AddMe();
            }
            catch (System.Exception ex)
            {
            }
        }
    }

    public class DefaultContextMenu
    {
        private static ContextMenuExtension s_cme;

        public static void RemoveMe()
        {
            Application.RemoveDefaultContextMenuExtension(s_cme);
        }

        public static void AddMe()
        {
            try
            {
                s_cme = new ContextMenuExtension();
                s_cme.Title = "TDMS Настройки";

                MenuItem mi = new MenuItem("Настройки");
                mi.Click += new EventHandler(callback_OnClick);

                s_cme.MenuItems.Add(mi);

                Application.AddDefaultContextMenuExtension(s_cme);
            }
            catch (System.Exception ex)
            {
            }
        }

        private static void callback_OnClick(Object o, EventArgs e)
        {
            try
            {
                Property PropertyForm = new Property();
                Application.ShowModalDialog(Application.MainWindow.Handle, PropertyForm);
            }
            catch (System.Exception ex)
            {
            }
        }
    }
}