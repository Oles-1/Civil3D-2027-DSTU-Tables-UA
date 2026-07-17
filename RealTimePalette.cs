using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using Document = Autodesk.AutoCAD.ApplicationServices.Document;
using Timer = System.Windows.Forms.Timer;
using Region = Autodesk.AutoCAD.DatabaseServices.Region;

namespace AutoCAD_CalcTools
{
    public class RealTimePalette : IExtensionApplication
    {
        static PaletteSet _ps = null;
        public static Label _lblLength;
        public static Label _lblArea;
        public static ComboBox _cmbLayers; // Наш новий список шарів

        static double _lastLength = 0;
        static double _lastArea = 0;
        static int _decimals = 2;

        static ObjectId[] _idsToRestore = null;

        // Створюємо структуру для зберігання даних у пам'яті
        class ShapeData
        {
            public string LayerName { get; set; }
            public double Length { get; set; }
            public double Area { get; set; }
        }

        // "Табличка" в пам'яті для миттєвої фільтрації
        static List<ShapeData> _cachedData = new List<ShapeData>();

        public void Initialize() { }
        public void Terminate() { }

        [CommandMethod("ShowCalc")]
        public void ShowPalette()
        {
            if (_ps == null)
            {
                _ps = new PaletteSet("Геометрія (Суми)");
                _ps.Size = new System.Drawing.Size(280, 240); // Трохи збільшили висоту для списку
                _ps.DockEnabled = DockSides.Left | DockSides.Right;
                _ps.KeepFocus = false;

                System.Drawing.Color darkBg = System.Drawing.Color.FromArgb(34, 34, 36);
                System.Drawing.Color btnColor = System.Drawing.Color.FromArgb(60, 60, 60);
                System.Drawing.Color whiteFg = System.Drawing.Color.White;

                Panel host = new Panel();
                host.Dock = DockStyle.Fill;
                host.BackColor = darkBg;

                var table = new TableLayoutPanel();
                table.Dock = DockStyle.Fill;
                table.BackColor = darkBg;
                table.Padding = new Padding(10);
                table.ColumnCount = 2;
                table.RowCount = 4; // Тепер у нас 4 рядки
                table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
                table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 45F));
                table.RowStyles.Add(new RowStyle(SizeType.Absolute, 45F)); // Кнопка розрахунку
                table.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F)); // Рядок для випадаючого списку
                table.RowStyles.Add(new RowStyle(SizeType.Percent, 50F)); // Довжина
                table.RowStyles.Add(new RowStyle(SizeType.Percent, 50F)); // Площа

                // --- ГОЛОВНА КНОПКА ---
                Button btnUpdate = new Button();
                btnUpdate.Text = "Розрахувати виділене";
                btnUpdate.Dock = DockStyle.Fill;
                btnUpdate.FlatStyle = FlatStyle.Flat;
                btnUpdate.BackColor = btnColor;
                btnUpdate.ForeColor = whiteFg;
                btnUpdate.Font = new System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold);
                btnUpdate.Cursor = Cursors.Hand;

                btnUpdate.Click += (s, e) =>
                {
                    Document doc = Application.DocumentManager.MdiActiveDocument;
                    if (doc != null) doc.SendStringToExecute("CMD_UPDATE_PALETTE ", true, false, false);
                };

                // --- ВИПАДАЮЧИЙ СПИСОК ШАРІВ ---
                _cmbLayers = new ComboBox();
                _cmbLayers.Dock = DockStyle.Fill;
                _cmbLayers.DropDownStyle = ComboBoxStyle.DropDownList; // Тільки вибір, без вводу тексту
                _cmbLayers.BackColor = btnColor;
                _cmbLayers.ForeColor = whiteFg;
                _cmbLayers.Font = new System.Drawing.Font("Segoe UI", 9);
                _cmbLayers.Items.Add("Всі шари");
                _cmbLayers.SelectedIndex = 0;

                // Що робити, коли ми обираємо інший шар у списку
                _cmbLayers.SelectedIndexChanged += (s, e) => FilterDataByLayer();

                var font = new System.Drawing.Font("Segoe UI", 11, System.Drawing.FontStyle.Bold);

                _lblLength = new Label() { Text = "Сума довжин: 0", Dock = DockStyle.Fill, TextAlign = System.Drawing.ContentAlignment.MiddleLeft, Font = font, ForeColor = whiteFg, BackColor = darkBg };
                _lblArea = new Label() { Text = "Сума площ: 0", Dock = DockStyle.Fill, TextAlign = System.Drawing.ContentAlignment.MiddleLeft, Font = font, ForeColor = whiteFg, BackColor = darkBg };

                Button btnCopyLength = CreateCopyButton(btnColor);
                Button btnCopyArea = CreateCopyButton(btnColor);

                btnCopyLength.Click += (s, e) => CopyToClipboard(_lastLength.ToString("F" + _decimals), btnCopyLength, btnColor);
                btnCopyArea.Click += (s, e) => CopyToClipboard(_lastArea.ToString("F" + _decimals), btnCopyArea, btnColor);

                // Збираємо все до купи
                table.Controls.Add(btnUpdate, 0, 0);
                table.SetColumnSpan(btnUpdate, 2);

                table.Controls.Add(_cmbLayers, 0, 1);
                table.SetColumnSpan(_cmbLayers, 2); // Розтягуємо список на всю ширину

                table.Controls.Add(_lblLength, 0, 2);
                table.Controls.Add(btnCopyLength, 1, 2);

                table.Controls.Add(_lblArea, 0, 3);
                table.Controls.Add(btnCopyArea, 1, 3);

                host.Controls.Add(table);
                _ps.Add("CalcPanel", host);
            }

            _ps.Visible = true;
        }

        private Button CreateCopyButton(System.Drawing.Color color)
        {
            Button btn = new Button();
            btn.Text = "📋";
            btn.Font = new System.Drawing.Font("Segoe UI Emoji", 14);
            btn.Size = new System.Drawing.Size(35, 35);
            btn.Anchor = AnchorStyles.None;
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
            btn.Cursor = Cursors.Hand;
            btn.BackColor = color;
            btn.ForeColor = System.Drawing.Color.White;
            return btn;
        }

        private void CopyToClipboard(string value, Button btn, System.Drawing.Color originalColor)
        {
            if (!string.IsNullOrEmpty(value) && value != "0")
            {
                System.Windows.Forms.Clipboard.SetText(value);
                btn.Text = "✔️";
                btn.BackColor = System.Drawing.Color.FromArgb(40, 167, 69);
                var timer = new Timer { Interval = 1000 };
                timer.Tick += (ts, te) =>
                {
                    btn.Text = "📋";
                    btn.BackColor = originalColor;
                    timer.Stop();
                    timer.Dispose();
                };
                timer.Start();
            }
        }

        // ========================================================
        // ЛОГІКА ФІЛЬТРАЦІЇ З ПАМ'ЯТІ
        // ========================================================
        private static void FilterDataByLayer()
        {
            if (_cmbLayers.SelectedItem == null || _lblLength == null || _lblArea == null) return;

            string selectedLayer = _cmbLayers.SelectedItem.ToString();
            double totalLength = 0;
            double totalArea = 0;

            // Перебираємо дані в пам'яті
            foreach (var item in _cachedData)
            {
                // Додаємо, якщо вибрано "Всі шари" або шар збігається
                if (selectedLayer == "Всі шари" || item.LayerName == selectedLayer)
                {
                    totalLength += item.Length;
                    totalArea += item.Area;
                }
            }

            _lastLength = totalLength;
            _lastArea = totalArea;

            _lblLength.Text = $"Сума довжин: {_lastLength.ToString("F" + _decimals)}";
            _lblArea.Text = $"Сума площ: {_lastArea.ToString("F" + _decimals)}";
        }

        [CommandMethod("CMD_UPDATE_PALETTE", CommandFlags.UsePickSet | CommandFlags.Transparent)]
        public void UpdatePaletteCommand()
        {
            if (_lblLength == null || _lblArea == null || _cmbLayers == null) return;

            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            try { _decimals = Convert.ToInt32(Application.GetSystemVariable("LUPREC")); } catch { _decimals = 2; }

            PromptSelectionResult psr = doc.Editor.SelectImplied();

            if (psr.Status == PromptStatus.OK && psr.Value != null)
            {
                _idsToRestore = psr.Value.GetObjectIds();
                _cachedData.Clear(); // Очищаємо стару пам'ять
                HashSet<string> uniqueLayers = new HashSet<string>(); // Для збору унікальних шарів

                using (doc.LockDocument())
                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    foreach (SelectedObject so in psr.Value)
                    {
                        DBObject obj = tr.GetObject(so.ObjectId, OpenMode.ForRead);
                        Entity ent = obj as Entity;
                        if (ent == null) continue;

                        string layer = ent.Layer;
                        uniqueLayers.Add(layer);

                        double len = 0;
                        double area = 0;

                        if (obj is Curve curve)
                        {
                            try { len = curve.GetDistanceAtParameter(curve.EndParam); } catch { }
                            try { area = curve.Area; } catch { }
                        }
                        else if (obj is Region reg)
                        {
                            try { area = reg.Area; } catch { }
                        }

                        // Зберігаємо об'єкт у пам'ять
                        _cachedData.Add(new ShapeData { LayerName = layer, Length = len, Area = area });
                    }
                    tr.Commit();
                }

                // Оновлюємо випадаючий список
                _cmbLayers.Items.Clear();
                _cmbLayers.Items.Add("Всі шари");

                // Сортуємо шари за алфавітом і додаємо в список
                var sortedLayers = uniqueLayers.ToList();
                sortedLayers.Sort();
                foreach (var l in sortedLayers) _cmbLayers.Items.Add(l);

                // Автоматично вибираємо "Всі шари" (це викличе метод FilterDataByLayer)
                _cmbLayers.SelectedIndex = 0;

                Application.Idle -= RestoreSelection;
                Application.Idle += RestoreSelection;
            }
        }

        private static void RestoreSelection(object sender, EventArgs e)
        {
            Application.Idle -= RestoreSelection;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc != null && _idsToRestore != null)
            {
                try { doc.Editor.SetImpliedSelection(_idsToRestore); }
                catch { }
                finally { _idsToRestore = null; }
            }
        }
    }
}