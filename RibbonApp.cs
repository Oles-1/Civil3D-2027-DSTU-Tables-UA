using System;
using System.Linq; // Додано для перевірки наявності Civil 3D
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;
using System.Windows.Media.Imaging;

[assembly: ExtensionApplication(typeof(CivilRibbonTest.RibbonApp))]

namespace CivilRibbonTest
{
    public class RibbonApp : IExtensionApplication
    {
        public void Initialize()
        {
            if (ComponentManager.Ribbon == null)
            {
                ComponentManager.ItemInitialized += ComponentManager_ItemInitialized;
            }
            else
            {
                CreateRibbon();
            }
        }

        public void Terminate() { }

        private void ComponentManager_ItemInitialized(object sender, RibbonItemEventArgs e)
        {
            if (ComponentManager.Ribbon != null)
            {
                ComponentManager.ItemInitialized -= ComponentManager_ItemInitialized;
                CreateRibbon();
            }
        }

        private void CreateRibbon()
        {
            try
            {
                // ==========================================================
                // ЗАПОБІЖНИК: Працює ТІЛЬКИ в Civil 3D
                // ==========================================================
                bool isCivil3D = System.AppDomain.CurrentDomain.GetAssemblies()
                    .Any(a => a.GetName().Name.StartsWith("AeccDbMgd"));

                // Якщо це звичайний AutoCAD (бібліотек Civil 3D немає) — виходимо
                if (!isCivil3D)
                {
                    return;
                }
                // ==========================================================

                RibbonControl ribbon = ComponentManager.Ribbon;
                if (ribbon == null) return;

                RibbonTab tab = new RibbonTab();
                tab.Title = "+Іструменти";
                tab.Id = "MY_CIVIL_TOOLS_TAB";

                RibbonPanelSource panelSource = new RibbonPanelSource();
                panelSource.Title = "Інструменти проектувальника, таблиці,..."; // Змінив назву панелі на більш загальну
                RibbonPanel panel = new RibbonPanel();
                panel.Source = panelSource;

                string dllFolder = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

                // --- Кнопка 1: Мала таблиця ---
                RibbonButton btnSmallTable = new RibbonButton();
                btnSmallTable.Text = "В.К.П.\n(мала)";
                btnSmallTable.ShowText = true;
                btnSmallTable.ShowImage = true;
                btnSmallTable.Size = RibbonItemSize.Large;
                btnSmallTable.LargeImage = LoadImage(System.IO.Path.Combine(dllFolder, "1.png"));
                btnSmallTable.CommandParameter = "DSTU_TABLE ";
                btnSmallTable.CommandHandler = new RibbonCommandHandler();

                // --- Кнопка 2: Велика таблиця ---
                RibbonButton btnFullTable = new RibbonButton();
                btnFullTable.Text = "В.К.П.\n(повна)";
                btnFullTable.ShowText = true;
                btnFullTable.ShowImage = true;
                btnFullTable.Size = RibbonItemSize.Large;
                btnFullTable.LargeImage = LoadImage(System.IO.Path.Combine(dllFolder, "2.png"));
                btnFullTable.CommandParameter = "DSTU_TABLE_FULL ";
                btnFullTable.CommandHandler = new RibbonCommandHandler();

                // --- Кнопка 3: Калькулятор (НОВА) ---
                RibbonButton btnShowCalc = new RibbonButton();
                btnShowCalc.Text = "Калькулятор\nсум";
                btnShowCalc.ShowText = true;
                btnShowCalc.ShowImage = true;
                btnShowCalc.Size = RibbonItemSize.Large;
                // Переконайся, що файл calc_icon.png лежить у папці з DLL
                btnShowCalc.LargeImage = LoadImage(System.IO.Path.Combine(dllFolder, "calc_icon.png"));
                btnShowCalc.CommandParameter = "ShowCalc ";
                btnShowCalc.CommandHandler = new RibbonCommandHandler();

                // Збираємо кнопки на панель (додаємо всі три)
                panelSource.Items.Add(btnSmallTable);
                panelSource.Items.Add(btnFullTable);
                panelSource.Items.Add(btnShowCalc); // Додали третю кнопку

                tab.Panels.Add(panel);
                ribbon.Tabs.Add(tab);

                // Рядок tab.IsActive = true; видалено, щоб вкладка не відкривалася примусово
            }
            catch (System.Exception ex)
            {
                Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    doc.Editor.WriteMessage("\nПомилка створення меню: " + ex.Message);
                }
            }
        }

        private BitmapImage LoadImage(string path)
        {
            try
            {
                return new BitmapImage(new Uri(path, UriKind.Absolute));
            }
            catch
            {
                return null;
            }
        }
    }

    public class RibbonCommandHandler : System.Windows.Input.ICommand
    {
        public bool CanExecute(object parameter) => true;
        public event EventHandler CanExecuteChanged;

        public void Execute(object parameter)
        {
            RibbonButton btn = parameter as RibbonButton;
            if (btn != null && btn.CommandParameter != null)
            {
                Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    doc.SendStringToExecute(btn.CommandParameter.ToString(), true, false, true);
                }
            }
        }
    }
}