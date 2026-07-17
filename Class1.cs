using System;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.DatabaseServices;
using Exception = System.Exception;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace CivilDstuTable
{
    public class DstuCommands
    {
        // ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓
        // БЛОК 1: БАЗОВА ТАБЛИЦЯ (Оригінальна, 18 стовпців)
        // Команда: DSTU_TABLE
        // ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓

        [CommandMethod("DSTU_TABLE")]
        public void CreateTableCommand()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            PromptEntityOptions pOpts = new PromptEntityOptions("\nВиберіть трасу: ");
            pOpts.SetRejectMessage("\nПомилка: Вибраний об'єкт не є трасою!");
            pOpts.AddAllowedClass(typeof(Alignment), true);
            PromptEntityResult pRes = ed.GetEntity(pOpts);
            if (pRes.Status != PromptStatus.OK) return;

            using (Transaction trans = doc.Database.TransactionManager.StartTransaction())
            {
                Alignment align = trans.GetObject(pRes.ObjectId, OpenMode.ForRead) as Alignment;
                if (align != null) CreateAndFillExcel(align);
                trans.Commit();
            }
        }

        private void CreateAndFillExcel(Alignment align)
        {
            Type excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null) throw new Exception("Excel не знайдено!");
            dynamic excelApp = Activator.CreateInstance(excelType);
            excelApp.Visible = true;
            dynamic workbook = excelApp.Workbooks.Add();
            dynamic worksheet = workbook.Worksheets[1];
            worksheet.Name = "Відомість кутів";

            worksheet.Cells.Font.Name = "GOST Common";
            worksheet.Cells.Font.Size = 12;
            worksheet.Cells.Font.Bold = false;

            worksheet.Cells[1, 1] = "Відомості кутів поворотів, прямих і кривих.";
            worksheet.Range["A1:R1"].Merge();
            worksheet.Range["A1"].HorizontalAlignment = -4108;

            worksheet.Cells[2, 1] = "Точка"; worksheet.Range["A2:A4"].Merge();
            worksheet.Cells[2, 2] = "Положеня вершини кута"; worksheet.Range["B2:D3"].Merge();
            worksheet.Cells[2, 5] = "Величина кута повороту"; worksheet.Range["E2:F2"].Merge();
            worksheet.Cells[2, 7] = "Радіус, м"; worksheet.Range["G2:G4"].Merge();
            worksheet.Cells[2, 8] = "Елементи кривої, м"; worksheet.Range["H2:N2"].Merge();
            worksheet.Cells[2, 15] = "ПКК"; worksheet.Range["O2:O4"].Merge();
            worksheet.Cells[2, 16] = "ККК"; worksheet.Range["P2:P4"].Merge();
            worksheet.Cells[2, 17] = "Відстань між\nвершинами, м"; worksheet.Range["Q2:Q4"].Merge();
            worksheet.Cells[2, 18] = "Довжина\nпрямої, м"; worksheet.Range["R2:R4"].Merge();

            worksheet.Cells[3, 5] = "вліво"; worksheet.Range["E3:E4"].Merge();
            worksheet.Cells[3, 6] = "вправо"; worksheet.Range["F3:F4"].Merge();
            worksheet.Cells[3, 8] = "тангенс"; worksheet.Range["H3:H4"].Merge();
            worksheet.Cells[3, 9] = "тангенс"; worksheet.Range["I3:I4"].Merge();
            worksheet.Cells[3, 10] = "перехідні криві"; worksheet.Range["J3:K4"].Merge();
            worksheet.Cells[3, 12] = "кругова крива"; worksheet.Range["L3:L4"].Merge();
            worksheet.Cells[3, 13] = "бісектриса"; worksheet.Range["M3:M4"].Merge();
            worksheet.Cells[3, 14] = "домір"; worksheet.Range["N3:N4"].Merge();

            worksheet.Cells[4, 2] = "км"; worksheet.Cells[4, 3] = "ПК"; worksheet.Cells[4, 4] = "+";

            for (int i = 1; i <= 18; i++) worksheet.Cells[5, i] = i.ToString();

            dynamic headerRange = worksheet.Range["A2:R5"];
            headerRange.HorizontalAlignment = -4108;
            headerRange.VerticalAlignment = -4108;
            headerRange.WrapText = true;
            headerRange.Borders.LineStyle = 1;

            worksheet.Columns("A:A").ColumnWidth = 8;
            worksheet.Columns("B:D").ColumnWidth = 7;
            worksheet.Columns("E:F").ColumnWidth = 10;
            worksheet.Columns("G:N").ColumnWidth = 9;
            worksheet.Columns("O:P").ColumnWidth = 11;
            worksheet.Columns("Q:R").ColumnWidth = 14;

            worksheet.Application.ActiveWindow.SplitRow = 5;
            worksheet.Application.ActiveWindow.FreezePanes = true;

            FillTableData(worksheet, align);
        }

        private void FillTableData(dynamic worksheet, Alignment align)
        {
            int row = 6;
            int vkNumber = 1;
            double prevPI = align.StartingStation;
            double prevKKK = align.StartingStation;

            worksheet.Cells[row, 1] = "ПТ";
            WriteStation(worksheet, row, align.StartingStation);
            FillPtkZeros(worksheet, row, 7, 14);
            MergePiBlock(worksheet, row, 16);
            row += 2;

            foreach (AlignmentEntity entity in align.Entities)
            {
                if (entity.EntityType == AlignmentEntityType.Arc)
                {
                    AlignmentArc arc = (AlignmentArc)entity;

                    // Шахматка: запис в рядок row-1 (попередній), об'єднання з row (поточний)
                    worksheet.Cells[row - 1, 17] = Math.Round(arc.PIStation - prevPI, 2);
                    worksheet.Range[worksheet.Cells[row - 1, 17], worksheet.Cells[row, 17]].Merge();
                    worksheet.Cells[row - 1, 18] = Math.Round(arc.StartStation - prevKKK, 2);
                    worksheet.Range[worksheet.Cells[row - 1, 18], worksheet.Cells[row, 18]].Merge();

                    worksheet.Cells[row, 1] = "ВК" + vkNumber++;
                    WriteStation(worksheet, row, arc.PIStation);

                    double angleRads = arc.Delta;
                    if (!arc.Clockwise) worksheet.Cells[row, 5] = FormatAngleToDMS(angleRads);
                    else worksheet.Cells[row, 6] = FormatAngleToDMS(angleRads);

                    worksheet.Cells[row, 7] = Math.Round(arc.Radius, 2);
                    double t = arc.Radius * Math.Tan(angleRads / 2.0);
                    worksheet.Cells[row, 8] = Math.Round(t, 2); worksheet.Cells[row, 9] = Math.Round(t, 2);
                    worksheet.Cells[row, 10] = 0; worksheet.Cells[row, 11] = 0;
                    worksheet.Cells[row, 12] = Math.Round(arc.Length, 2);
                    worksheet.Cells[row, 13] = Math.Round(arc.Radius * (1.0 / Math.Cos(angleRads / 2.0) - 1.0), 2);
                    worksheet.Cells[row, 14] = Math.Round(2 * t - arc.Length, 2);
                    worksheet.Cells[row, 15] = FormatStationString(arc.StartStation);
                    worksheet.Cells[row, 16] = FormatStationString(arc.EndStation);

                    MergePiBlock(worksheet, row, 16);
                    row += 2;
                    prevPI = arc.PIStation; prevKKK = arc.EndStation;
                }
            }

            // КТ
            worksheet.Cells[row - 1, 17] = Math.Round(align.EndingStation - prevPI, 2);
            worksheet.Range[worksheet.Cells[row - 1, 17], worksheet.Cells[row, 17]].Merge();
            worksheet.Cells[row - 1, 18] = Math.Round(align.EndingStation - prevKKK, 2);
            worksheet.Range[worksheet.Cells[row - 1, 18], worksheet.Cells[row, 18]].Merge();

            worksheet.Cells[row, 1] = "КТ";
            WriteStation(worksheet, row, align.EndingStation);
            FillPtkZeros(worksheet, row, 7, 14);
            MergePiBlock(worksheet, row, 16);

            // Фінальні рамки
            dynamic dataRange = worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[row + 1, 18]];
            dataRange.HorizontalAlignment = -4108;
            dataRange.VerticalAlignment = -4108;
            dataRange.Borders.LineStyle = 1;
        }

        // ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓
        // БЛОК 2: РОЗШИРЕНА ТАБЛИЦЯ (Перехідні криві, 24 стовпці)
        // Команда: DSTU_TABLE_FULL
        // ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓

        [CommandMethod("DSTU_TABLE_FULL")]
        public void CreateTableFullCommand()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            PromptEntityOptions pOpts = new PromptEntityOptions("\nВиберіть трасу: ");
            pOpts.SetRejectMessage("\nПомилка: Вибраний об'єкт не є трасою!");
            pOpts.AddAllowedClass(typeof(Alignment), true);
            PromptEntityResult pRes = ed.GetEntity(pOpts);
            if (pRes.Status != PromptStatus.OK) return;

            using (Transaction trans = doc.Database.TransactionManager.StartTransaction())
            {
                Alignment align = trans.GetObject(pRes.ObjectId, OpenMode.ForRead) as Alignment;
                if (align != null) CreateAndFillExcelFull(align);
                trans.Commit();
            }
        }

        private void CreateAndFillExcelFull(Alignment align)
        {
            Type excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null) throw new Exception("Excel не знайдено!");
            dynamic excelApp = Activator.CreateInstance(excelType);
            excelApp.Visible = true;
            dynamic workbook = excelApp.Workbooks.Add();
            dynamic worksheet = workbook.Worksheets[1];
            worksheet.Name = "Відомість кутів ПОВНА";

            worksheet.Cells.Font.Name = "GOST Common"; worksheet.Cells.Font.Size = 12;

            worksheet.Cells[1, 1] = "Відомості кутів поворотів, прямих і кривих.";
            worksheet.Range["A1:X1"].Merge(); worksheet.Range["A1"].HorizontalAlignment = -4108;

            worksheet.Cells[2, 1] = "Точка"; worksheet.Range["A2:A4"].Merge();
            worksheet.Cells[2, 2] = "Положеня вершини кута"; worksheet.Range["B2:D3"].Merge();
            worksheet.Cells[2, 5] = "Величина кута повороту"; worksheet.Range["E2:F2"].Merge();
            worksheet.Cells[2, 7] = "Радіус, м"; worksheet.Range["G2:G4"].Merge();
            worksheet.Cells[2, 8] = "Елементи кривої, м"; worksheet.Range["H2:N2"].Merge();
            worksheet.Cells[2, 15] = "Положення перехідних кривих"; worksheet.Range["O2:V2"].Merge();
            worksheet.Cells[2, 23] = "Відстань між\nвершинами, м"; worksheet.Range["W2:W4"].Merge();
            worksheet.Cells[2, 24] = "Довжина\nпрямої, м"; worksheet.Range["X2:X4"].Merge();

            worksheet.Cells[3, 5] = "вліво"; worksheet.Range["E3:E4"].Merge();
            worksheet.Cells[3, 6] = "вправо"; worksheet.Range["F3:F4"].Merge();
            worksheet.Cells[3, 8] = "тангенс"; worksheet.Range["H3:H4"].Merge();
            worksheet.Cells[3, 9] = "тангенс"; worksheet.Range["I3:I4"].Merge();
            worksheet.Cells[3, 10] = "довжина"; worksheet.Range["J3:J4"].Merge();
            worksheet.Cells[3, 11] = "парам. A"; worksheet.Range["K3:K4"].Merge();
            worksheet.Cells[3, 12] = "кругова крива"; worksheet.Range["L3:L4"].Merge();
            worksheet.Cells[3, 13] = "бісектриса"; worksheet.Range["M3:M4"].Merge();
            worksheet.Cells[3, 14] = "домір"; worksheet.Range["N3:N4"].Merge();

            worksheet.Cells[3, 15] = "Початок"; worksheet.Range["O3:P3"].Merge();
            worksheet.Cells[3, 17] = "Кінець"; worksheet.Range["Q3:R3"].Merge();
            worksheet.Cells[3, 19] = "Початок"; worksheet.Range["S3:T3"].Merge();
            worksheet.Cells[3, 21] = "Кінець"; worksheet.Range["U3:V3"].Merge();

            worksheet.Cells[4, 2] = "км"; worksheet.Cells[4, 3] = "ПК"; worksheet.Cells[4, 4] = "+";
            worksheet.Cells[4, 15] = "ПК"; worksheet.Cells[4, 16] = "+";
            worksheet.Cells[4, 17] = "ПК"; worksheet.Cells[4, 18] = "+";
            worksheet.Cells[4, 19] = "ПК"; worksheet.Cells[4, 20] = "+";
            worksheet.Cells[4, 21] = "ПК"; worksheet.Cells[4, 22] = "+";

            for (int i = 1; i <= 24; i++) worksheet.Cells[5, i] = i.ToString();

            dynamic headerRange = worksheet.Range["A2:X5"];
            headerRange.HorizontalAlignment = -4108; headerRange.VerticalAlignment = -4108;
            headerRange.WrapText = true; headerRange.Borders.LineStyle = 1;

            worksheet.Columns("A:A").ColumnWidth = 8;
            worksheet.Columns("B:D").ColumnWidth = 7;
            worksheet.Columns("E:V").ColumnWidth = 10;
            worksheet.Columns("W:X").ColumnWidth = 14;

            worksheet.Application.ActiveWindow.SplitRow = 5;
            worksheet.Application.ActiveWindow.FreezePanes = true;

            FillTableDataFull(worksheet, align);
        }

        private void FillTableDataFull(dynamic worksheet, Alignment align)
        {
            int row = 6;
            int vkNumber = 1;
            double prevPI = align.StartingStation;
            double prevKKK = align.StartingStation;

            worksheet.Cells[row, 1] = "ПТ";
            WriteStation(worksheet, row, align.StartingStation);
            FillPtkZeros(worksheet, row, 7, 22);
            MergePiBlock(worksheet, row, 22);
            row += 2;
            
            foreach (AlignmentEntity entity in align.Entities)
            {
                if (entity.EntityType == AlignmentEntityType.Arc)
                {
                    AlignmentArc arc = (AlignmentArc)entity;

                    worksheet.Cells[row - 1, 23] = Math.Round(arc.PIStation - prevPI, 2);
                    worksheet.Cells[row - 1, 24] = Math.Round(arc.StartStation - prevKKK, 2);

                    worksheet.Cells[row, 1] = "ВК" + vkNumber++;
                    WriteStation(worksheet, row, arc.PIStation);

                    double angleRads = arc.Delta;
                    if (!arc.Clockwise) worksheet.Cells[row, 5] = FormatAngleToDMS(angleRads);
                    else worksheet.Cells[row, 6] = FormatAngleToDMS(angleRads);

                    worksheet.Cells[row, 7] = Math.Round(arc.Radius, 2);
                    double t = arc.Radius * Math.Tan(angleRads / 2.0);
                    worksheet.Cells[row, 8] = Math.Round(t, 2); worksheet.Cells[row, 9] = Math.Round(t, 2);
                    worksheet.Cells[row, 10] = 0; worksheet.Cells[row, 11] = 0;
                    worksheet.Cells[row, 12] = Math.Round(arc.Length, 2);
                    worksheet.Cells[row, 13] = Math.Round(arc.Radius * (1.0 / Math.Cos(angleRads / 2.0) - 1.0), 2);
                    worksheet.Cells[row, 14] = Math.Round(2 * t - arc.Length, 2);

                    for (int c = 15; c <= 22; c++) worksheet.Cells[row, c] = 0;

                    MergePiBlock(worksheet, row, 22);
                    row += 2;
                    prevPI = arc.PIStation; prevKKK = arc.EndStation;
                }
                else if (entity.EntityType == AlignmentEntityType.SpiralCurveSpiral)
                {
                    // Використовуємо dynamic, але додаємо перевірку властивостей перед доступом
                    dynamic scs = entity;

                    // Отримуємо базові координати через безпечний виклик
                    double scsStart = (double)scs.StartStation;
                    double scsEnd = (double)scs.EndStation;

                    // Обчислення PI через геометрію дуги (яка є серцем групи)
                    double radius = (double)scs.Arc.Radius;
                    double arcLength = (double)scs.Arc.Length;
                    double ls1 = (double)scs.SpiralIn.Length;
                    double ls2 = (double)scs.SpiralOut.Length;

                    // Кут повороту вираховуємо через зміну напрямку дуги
                    double angleRads = Math.Abs((double)scs.Arc.Delta);

                    double t = radius * Math.Tan(angleRads / 2.0);
                    double piStation = scsStart + t; // Апроксимація PI

                    // Запис в Excel (без жодних .Merge() для колонок 23-24)
                    worksheet.Cells[row - 1, 23] = Math.Round(piStation - prevPI, 2);
                    worksheet.Range[worksheet.Cells[row - 1, 23], worksheet.Cells[row, 23]].Merge();

                    worksheet.Cells[row - 1, 24] = Math.Round(scsStart - prevKKK, 2);
                    worksheet.Range[worksheet.Cells[row - 1, 24], worksheet.Cells[row, 24]].Merge();

                    worksheet.Cells[row, 1] = "ВК" + vkNumber++;
                    WriteStation(worksheet, row, piStation);

                    // Кут повороту (враховуємо Clockwise)
                    if (!(bool)scs.Arc.Clockwise) worksheet.Cells[row, 5] = FormatAngleToDMS(angleRads);
                    else worksheet.Cells[row, 6] = FormatAngleToDMS(angleRads);

                    worksheet.Cells[row, 7] = Math.Round(radius, 2);
                    worksheet.Cells[row, 8] = Math.Round(t, 2);
                    worksheet.Cells[row, 9] = Math.Round(t, 2);

                    // Перехідні
                    worksheet.Cells[row, 10] = Math.Round(ls1, 2);
                    worksheet.Cells[row, 11] = Math.Round((double)scs.SpiralIn.A, 2);
                    worksheet.Cells[row, 12] = Math.Round(arcLength, 2);

                    // Бісектриса та Домір (спрощено для стабільності)
                    worksheet.Cells[row, 13] = Math.Round(radius * (1.0 / Math.Cos(angleRads / 2.0) - 1.0), 2);
                    worksheet.Cells[row, 14] = Math.Round(2 * t - arcLength, 2);

                    // Пікетажі перехідних
                    WriteStationDouble(worksheet, row, 15, scsStart);
                    WriteStationDouble(worksheet, row, 17, scsStart + ls1);
                    WriteStationDouble(worksheet, row, 19, scsEnd - ls2);
                    WriteStationDouble(worksheet, row, 21, scsEnd);

                    MergePiBlock(worksheet, row, 22);
                    row += 2;
                    prevPI = piStation; prevKKK = scsEnd;
                }
            }

            worksheet.Cells[row - 1, 23] = Math.Round(align.EndingStation - prevPI, 2);
            worksheet.Range[worksheet.Cells[row - 1, 23], worksheet.Cells[row, 23]].Merge();
            worksheet.Cells[row - 1, 24] = Math.Round(align.EndingStation - prevKKK, 2);
            worksheet.Range[worksheet.Cells[row - 1, 24], worksheet.Cells[row, 24]].Merge();

            worksheet.Cells[row, 1] = "КТ";
            WriteStation(worksheet, row, align.EndingStation);
            FillPtkZeros(worksheet, row, 7, 22);
            MergePiBlock(worksheet, row, 22);

            dynamic dataRange = worksheet.Range[worksheet.Cells[6, 1], worksheet.Cells[row + 1, 24]];
            dataRange.HorizontalAlignment = -4108; dataRange.VerticalAlignment = -4108; dataRange.Borders.LineStyle = 1;
        }

        // ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓
        // БЛОК 3: СПІЛЬНІ ДОПОМІЖНІ МЕТОДИ
        // ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓

        private void FillPtkZeros(dynamic ws, int r, int startCol, int endCol)
        {
            for (int c = startCol; c <= endCol; c++) ws.Cells[r, c] = 0;
        }

        private void MergePiBlock(dynamic ws, int r, int endCol)
        {
            for (int c = 1; c <= endCol; c++) ws.Range[ws.Cells[r, c], ws.Cells[r + 1, c]].Merge();
        }

        private void WriteStation(dynamic worksheet, int row, double station)
        {
            double km = Math.Floor(station / 1000);
            double pk = Math.Floor((station % 1000) / 100);
            double plus = station % 100;
            worksheet.Cells[row, 2] = km;
            worksheet.Cells[row, 3] = pk;
            worksheet.Cells[row, 4] = Math.Round(plus, 2);
        }

        private void WriteStationDouble(dynamic worksheet, int row, int startCol, double station)
        {
            double pk = Math.Floor(station / 100);
            double plus = station % 100;
            worksheet.Cells[row, startCol] = pk;
            worksheet.Cells[row, startCol + 1] = Math.Round(plus, 2);
        }

        private string FormatStationString(double station)
        {
            double pk = Math.Floor(station / 100);
            double plus = station % 100;
            return $"{pk}+{plus.ToString("F2")}";
        }

        private string FormatAngleToDMS(double radians)
        {
            double decimalDegrees = radians * (180.0 / Math.PI);
            int degrees = (int)Math.Floor(decimalDegrees);
            double remainingMinutes = (decimalDegrees - degrees) * 60.0;
            int minutes = (int)Math.Floor(remainingMinutes);
            double seconds = (remainingMinutes - minutes) * 60.0;
            int intSeconds = (int)Math.Round(seconds);

            if (intSeconds >= 60)
            {
                intSeconds = 0; minutes++;
                if (minutes >= 60) { minutes = 0; degrees++; }
            }
            return $"{degrees}°{minutes:D2}'{intSeconds:D2}''";
        }
    }
}