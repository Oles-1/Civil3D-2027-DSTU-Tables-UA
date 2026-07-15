using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows; // Для роботи зі стрічкою (Ribbon)
using System.Windows.Media.Imaging; // Для роботи з картинками (потребує PresentationCore)

// Вказуємо AutoCAD, що цей клас треба запустити автоматично при завантаженні DLL
[assembly: ExtensionApplication(typeof(CivilRibbonTest.RibbonApp))]

namespace CivilRibbonTest
{
    // IExtensionApplication - це інтерфейс, який має два обов'язкових методи: Initialize та Terminate
    public class RibbonApp : IExtensionApplication
    {
        // Метод Initialize спрацьовує один раз в момент команди NETLOAD
        public void Initialize()
        {
            // Перевіряємо, чи вже завантажився графічний інтерфейс Civil 3D
            if (ComponentManager.Ribbon == null)
            {
                // Якщо ще не завантажився, ми "підписуємося" на подію його створення
                // Це захищає нас від крашу при старті програми
                ComponentManager.ItemInitialized += ComponentManager_ItemInitialized;
            }
            else
            {
                // Якщо інтерфейс вже готовий, одразу створюємо нашу вкладку
                CreateRibbon();
            }
        }

        // Метод Terminate спрацьовує при закритті AutoCAD. Тут ми нічого не робимо.
        public void Terminate() { }

        // Цей метод викличеться автоматично, коли AutoCAD закінчить малювати свій інтерфейс
        private void ComponentManager_ItemInitialized(object sender, RibbonItemEventArgs e)
        {
            if (ComponentManager.Ribbon != null)
            {
                // Відписуємося від події, щоб вона не спрацювала двічі
                ComponentManager.ItemInitialized -= ComponentManager_ItemInitialized;
                CreateRibbon();
            }
        }

        // Головний метод створення нашого меню
        private void CreateRibbon()
        {
            try
            {
                // Отримуємо доступ до головної стрічки (Ribbon) AutoCAD
                RibbonControl ribbon = ComponentManager.Ribbon;

                // ==========================================
                // 1. СТВОРЕННЯ ВКАДКИ (RibbonTab)
                // ==========================================
                RibbonTab tab = new RibbonTab();
                tab.Title = "Олеся ніщаки"; // Назва, яку ти бачиш зверху
                tab.Id = "MY_CIVIL_TOOLS_TAB"; // Унікальний внутрішній ідентифікатор

                // ==========================================
                // 2. СТВОРЕННЯ ПАНЕЛІ (RibbonPanel)
                // ==========================================
                RibbonPanelSource panelSource = new RibbonPanelSource();
                panelSource.Title = "Таблиці ДСТУ 2008 (2026 актуальні)"; // Назва знизу під кнопками
                RibbonPanel panel = new RibbonPanel();
                panel.Source = panelSource;

                // ==========================================
                // 3. СТВОРЕННЯ КНОПОК
                // ==========================================
                string dllFolder = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                // Кнопка 1: Мала таблиця
                RibbonButton btnSmallTable = new RibbonButton();
                btnSmallTable.Text = "В.К.П.\n(мала)"; // \n - це новий рядок
                btnSmallTable.ShowText = true;
                btnSmallTable.ShowImage = true; // Тепер ми хочемо бачити картинку
                btnSmallTable.Size = RibbonItemSize.Large; // Робимо кнопку великою
                // Вказуємо шлях до твоєї картинки (заміни на свій реальний шлях)
                btnSmallTable.LargeImage = LoadImage(System.IO.Path.Combine(dllFolder, "1.png"));
                // Команда з пробілом в кінці (пробіл імітує натискання Enter)
                btnSmallTable.CommandParameter = "DSTU_TABLE ";
                btnSmallTable.CommandHandler = new RibbonCommandHandler();

                // Кнопка 2: Велика таблиця
                RibbonButton btnFullTable = new RibbonButton();
                btnFullTable.Text = "В.К.П.\n(повна)";
                btnFullTable.ShowText = true;
                btnFullTable.ShowImage = true;
                btnFullTable.Size = RibbonItemSize.Large;
                // Вказуємо шлях до картинки для другої кнопки
                btnFullTable.LargeImage = LoadImage(System.IO.Path.Combine(dllFolder, "2.png"));
                btnFullTable.CommandParameter = "DSTU_TABLE_FULL ";
                btnFullTable.CommandHandler = new RibbonCommandHandler();

                // ==========================================
                // 4. ЗБИРАННЯ ВСЬОГО ДОКУПИ
                // ==========================================
                // Додаємо кнопки на панель
                panelSource.Items.Add(btnSmallTable);
                panelSource.Items.Add(btnFullTable);

                // Додаємо панель у вкладку
                tab.Panels.Add(panel);

                // Додаємо вкладку у головну стрічку Civil 3D
                ribbon.Tabs.Add(tab);

                // Робимо нашу вкладку активною (щоб вона відкрилася сама після завантаження)
                tab.IsActive = true;
            }
            catch (System.Exception ex) // Явно вказуємо System.Exception, щоб уникнути конфлікту
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    doc.Editor.WriteMessage("\nПомилка створення меню: " + ex.Message);
                }
            }
        }

        // ==========================================
        // ДОПОМІЖНИЙ МЕТОД ДЛЯ ЗАВАНТАЖЕННЯ ІКОНОК
        // ==========================================
        // Цей метод безпечно намагається завантажити картинку з жорсткого диска
        private BitmapImage LoadImage(string path)
        {
            try
            {
                // UriKind.Absolute означає, що ми передаємо повний шлях (наприклад C:\...)
                return new BitmapImage(new Uri(path, UriKind.Absolute));
            }
            catch
            {
                // Якщо файлу картинки немає за вказаним шляхом, програма не крашнеться,
                // а просто поверне null (кнопка буде без картинки)
                return null;
            }
        }
    }

    // ==========================================
    // КЛАС-ОБРОБНИК НАТИСКАННЯ КНОПОК
    // ==========================================
    // Цей клас відповідає за те, що відбувається, коли ти клікаєш на кнопку в стрічці
    public class RibbonCommandHandler : System.Windows.Input.ICommand
    {
        // Дозволяє натискати кнопку завжди
        public bool CanExecute(object parameter) => true;

        // Обов'язкова подія для інтерфейсу ICommand (ми її не використовуємо, але вона має бути)
        public event EventHandler CanExecuteChanged;

        // Саме цей метод спрацьовує при кліку
        public void Execute(object parameter)
        {
            // Перевіряємо, чи параметр дійсно є нашою кнопкою
            RibbonButton btn = parameter as RibbonButton;
            if (btn != null && btn.CommandParameter != null)
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    // Відправляємо текст (CommandParameter) у командний рядок Civil 3D.
                    // Наприклад, він відправить "DSTU_TABLE " - і твоя таблиця почне будуватися.
                    doc.SendStringToExecute(btn.CommandParameter.ToString(), true, false, true);
                }
            }
        }
    }
}