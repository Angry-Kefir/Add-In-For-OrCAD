using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace AddInForOrcad
{
    public partial class ThisAddIn
    {
        public struct Page
        {
            public Excel.Range allPage, // страница целиком
                partList,               // только таблица
                partTable,              // таблица без шапки (только список)
                frameTable,             // основная подпись
                addFrameTable;          // дополнительная таблица слева
        }

        public struct Part
        {
            public string Availability, // наличие компонента
                Current,                // максимальный ток (?)
                Description,            // словесное описание
                Distributor,            // фирма поставщик
                DistributorPartNumber,  // обозначение у поставщика
                ESR,                    // Equivalent Series Resistance конденсатора (активные потери в цепи перем. тока)
                ItemNumber,             // номер в схеме
                Manufacturer,           // фирма изготовитель
                ManufacturerPartNumber, // обозначение у производителя
                Model,                  // модель
                Package,                // нечто для ICs (Integrated Circuit) (?)
                PartName,               // имя группы к которой принадлежит элемент (конденсатор, резистор и пр.)
                PartNumber,             // номер в каталоге производителя
                PartReference,          // имя компонента
                PartType,               // тип компонента
                PcbFootprint,           // наименование корпуса
                Power,                  // мощность
                Price,                  // цена
                Quantity,               // количество
                SchematicPart,          // наименование в библиотеке
                Speed,                  // скорость для ICs (Integrated Circuit) (?)
                T,                      // некая величина времени (?)
                Temperature,            // температура
                TKE,                    // температурный коэффициент ёмкости конденсаторов
                Tolerance,              // допустимое отклонение от номинала
                Value,                  // номинал
                ValueBOM,               // русский номинал для таблицы
                ValueRus,               // русский номинал для схем
                Voltage;                // вольтаж
            public int group;           // номер группы элемента
        }

        public List<Page> pages = new List<Page>();

        void PrepareSheetForTable(Excel.Worksheet currentSheet)
        {
            Excel.Range allCells = currentSheet.Range["A1", "XFD1048576"];
            allCells.Clear();

            // шрифт по умолчанию для всего листа
            allCells.Font.Name = "GOST Common";
            allCells.Font.ColorIndex = 1;
            allCells.Font.Italic = true;
            allCells.Font.Bold = false;
            allCells.Font.Size = 12;
            allCells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            allCells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            // для удобства разобьем лист на маленькие квадратные ячейки
            allCells.ColumnWidth = 2;
            allCells.RowHeight = 15.75;

            // исключения по ширине столбцов ячеек
            currentSheet.Range["A1"].ColumnWidth = 1.86;
            currentSheet.Range["D1"].ColumnWidth = 3.57;
            currentSheet.Range["K1"].ColumnWidth = 1.00;
            currentSheet.Range["AK1"].ColumnWidth = 3.57;

            // установка параметров листа для печати
            var cmInPoints = 0.5 / 0.035;
            currentSheet.PageSetup.RightMargin = cmInPoints;
            currentSheet.PageSetup.FooterMargin = cmInPoints;
            currentSheet.PageSetup.LeftMargin = cmInPoints;
            currentSheet.PageSetup.TopMargin = cmInPoints;
            currentSheet.PageSetup.BottomMargin = cmInPoints;
            currentSheet.PageSetup.HeaderMargin = cmInPoints;
            currentSheet.PageSetup.Zoom = 99;
        }

        void GenerateEmptyTable(Excel.Worksheet currentSheet, bool firstSheet)
        {
            var startCell = currentSheet.Range["A1"]; // первая страница всегда будет от А1
            //var startCell = currentSheet.Cells[1 + 53 * pages.Count, 1];

            Excel.Range temp;   // переменная для промежуточного хранения диапазонов, т.к. невозможна обработка операции в одной строке
                                // приходится разбить на 2 :(

            // основные диапазоны страницы относительно начальной клетки
            Excel.Range allPage = currentSheet.Range[startCell,
                currentSheet.Cells[startCell.Row + 52, startCell.Column + 36]]; // страница целиком

            Excel.Range partList = currentSheet.Range[currentSheet.Cells[startCell.Row, startCell.Column + 3],
                currentSheet.Cells[startCell.Row + 44, startCell.Column + 36]]; // только таблица

            Excel.Range partTable = currentSheet.Range[currentSheet.Cells[startCell.Row + 3, startCell.Column + 3],
                currentSheet.Cells[startCell.Row + 44, startCell.Column + 36]]; // таблица без шапки (только список)

            Excel.Range frameTable = currentSheet.Range[currentSheet.Cells[startCell.Row + 45, startCell.Column + 3],
                currentSheet.Cells[startCell.Row + 52, startCell.Column + 36]]; // основная подпись

            Excel.Range addFrameTable = currentSheet.Range[currentSheet.Cells[startCell.Row + 24, startCell.Column + 1],
                currentSheet.Cells[startCell.Row + 52, startCell.Column + 2]];  // дополнительная таблица слева

            #region Создание сетки (границ) рамки и таблицы относительно диапазонов
            partList.Borders.Weight = Excel.XlBorderWeight.xlMedium;                              // внутренние границы таблицы
            partList.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick); // толстый контур таблицы

            frameTable.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;                          // внутренние границы подписи
            frameTable.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick); // толстый контур подписи

            addFrameTable.Borders.Weight = Excel.XlBorderWeight.xlThick; // все границы таблицы слева


            // много мелких границ для подписи
            temp = currentSheet.Range[currentSheet.Cells[frameTable.Row, frameTable.Column],
                currentSheet.Cells[frameTable.Row + 7, frameTable.Column + 33]];
            temp.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            temp = currentSheet.Range[currentSheet.Cells[frameTable.Row, frameTable.Column + 1],
                currentSheet.Cells[frameTable.Row + 2, frameTable.Column + 2]];
            temp.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            temp = currentSheet.Range[currentSheet.Cells[frameTable.Row, frameTable.Column + 3],
                currentSheet.Cells[frameTable.Row + 7, frameTable.Column + 7]];
            temp.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            temp = currentSheet.Range[currentSheet.Cells[frameTable.Row, frameTable.Column + 11],
                currentSheet.Cells[frameTable.Row + 7, frameTable.Column + 12]];
            temp.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);


            temp = currentSheet.Range[currentSheet.Cells[frameTable.Row + 2, frameTable.Column],
                currentSheet.Cells[frameTable.Row + 2, frameTable.Column + 12]];
            temp.Borders.Weight = Excel.XlBorderWeight.xlThick;


            temp = currentSheet.Range[currentSheet.Cells[frameTable.Row + 3, frameTable.Column + 13],
                currentSheet.Cells[frameTable.Row + 7, frameTable.Column + 24]];
            temp.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);


            temp = currentSheet.Range[currentSheet.Cells[frameTable.Row + 3, frameTable.Column + 25],
                currentSheet.Cells[frameTable.Row + 3, frameTable.Column + 33]];
            temp.Borders.Weight = Excel.XlBorderWeight.xlThick;

            temp = currentSheet.Range[currentSheet.Cells[frameTable.Row + 4, frameTable.Column + 28],
                currentSheet.Cells[frameTable.Row + 4, frameTable.Column + 33]];
            temp.Borders.Weight = Excel.XlBorderWeight.xlThick;


            temp = currentSheet.Range[currentSheet.Cells[frameTable.Row + 5, frameTable.Column + 25],
                currentSheet.Cells[frameTable.Row + 7, frameTable.Column + 33]];
            temp.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);
            #endregion

            #region Объединение и заполнение ячеек

            // --- ДЛЯ СПИСКА КОМПОНЕНТОВ ---

            temp = currentSheet.Cells[partList.Row, partList.Column];
            temp.Value2 = "Поз. обозна-чение";
            temp = currentSheet.Range[currentSheet.Cells[partList.Row, partList.Column],
                currentSheet.Cells[partList.Row + 2, partList.Column + 2]];
            temp.WrapText = true;
            temp.Merge();

            temp = currentSheet.Cells[partList.Row, partList.Column + 3];
            temp.Value2 = "Наименование";
            temp = currentSheet.Range[currentSheet.Cells[partList.Row, partList.Column + 3],
                currentSheet.Cells[partList.Row + 2, partList.Column + 24]];
            temp.Merge();

            temp = currentSheet.Cells[partList.Row, partList.Column + 25];
            temp.Value2 = "Кол.";
            temp = currentSheet.Range[currentSheet.Cells[partList.Row, partList.Column + 25],
                currentSheet.Cells[partList.Row + 2, partList.Column + 26]];
            temp.Merge();

            temp = currentSheet.Cells[partList.Row, partList.Column + 27];
            temp.Value2 = "Примечание";
            temp = currentSheet.Range[currentSheet.Cells[partList.Row, partList.Column + 27],
                currentSheet.Cells[partList.Row + 2, partList.Column + 33]];
            temp.Merge();

            for (int i = 3; i <= 44; i++) // строки таблицы
            {
                temp = currentSheet.Range[currentSheet.Cells[partList.Row + i, partList.Column],
                    currentSheet.Cells[partList.Row + i, partList.Column + 2]];
                temp.Merge();

                temp = currentSheet.Range[currentSheet.Cells[partList.Row + i, partList.Column + 3],
                    currentSheet.Cells[partList.Row + i, partList.Column + 24]];
                temp.Merge();

                temp = currentSheet.Range[currentSheet.Cells[partList.Row + i, partList.Column + 25],
                    currentSheet.Cells[partList.Row + i, partList.Column + 26]];
                temp.Merge();

                temp = currentSheet.Range[currentSheet.Cells[partList.Row + i, partList.Column + 27],
                    currentSheet.Cells[partList.Row + i, partList.Column + 33]];
                temp.Merge();
            }

            // выравнивание по левому краю для наименований и примечаний
            temp = currentSheet.Range[currentSheet.Cells[partList.Row + 3, partList.Column + 3],
                currentSheet.Cells[partList.Row + 44, partList.Column + 24]];
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            temp = currentSheet.Range[currentSheet.Cells[partList.Row + 3, partList.Column + 27],
                currentSheet.Cells[partList.Row + 44, partList.Column + 33]];
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;


            // --- ДЛЯ РАМКИ (ПОДПИСЬ) ---

            temp = currentSheet.Range[currentSheet.Cells[frameTable.Row, frameTable.Column],
                currentSheet.Cells[frameTable.Row + 7, frameTable.Column + 12]];
            temp.Font.Size = 10;

            temp = currentSheet.Cells[frameTable.Row + 2, frameTable.Column];
            temp.Value2 = "Изм.";

            temp = currentSheet.Cells[frameTable.Row + 2, frameTable.Column + 1];
            temp.Value2 = "Лист";

            temp = currentSheet.Cells[frameTable.Row + 3, frameTable.Column];
            temp.Value2 = "Разраб.";

            temp = currentSheet.Cells[frameTable.Row + 4, frameTable.Column];
            temp.Value2 = "Пров.";

            temp = currentSheet.Cells[frameTable.Row + 6, frameTable.Column];
            temp.Value2 = "Н. контр.";

            temp = currentSheet.Cells[frameTable.Row + 7, frameTable.Column];
            temp.Value2 = "Утв.";

            for (int i = 0; i <= 2; i++)
            {
                temp = currentSheet.Range[currentSheet.Cells[frameTable.Row + i, frameTable.Column + 1],
                    currentSheet.Cells[frameTable.Row + i, frameTable.Column + 2]];
                temp.Merge();
            }

            for (int i = 3; i <= 7; i++)
            {
                temp = currentSheet.Range[currentSheet.Cells[frameTable.Row + i, frameTable.Column],
                    currentSheet.Cells[frameTable.Row + i, frameTable.Column + 2]];
                temp.Merge();
                temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            }

            temp = currentSheet.Cells[frameTable.Row + 2, frameTable.Column + 3];
            temp.Value2 = "№ докум.";

            temp = currentSheet.Cells[frameTable.Row + 2, frameTable.Column + 8];
            temp.Value2 = "Подп.";

            temp = currentSheet.Cells[frameTable.Row + 2, frameTable.Column + 11];
            temp.Value2 = "Дата";

            for (int i = 0; i <= 7; i++)
            {
                temp = currentSheet.Range[currentSheet.Cells[frameTable.Row + i, frameTable.Column + 3],
                    currentSheet.Cells[frameTable.Row + i, frameTable.Column + 7]];
                temp.Merge();

                temp = currentSheet.Range[currentSheet.Cells[frameTable.Row + i, frameTable.Column + 8],
                    currentSheet.Cells[frameTable.Row + i, frameTable.Column + 10]];
                temp.Merge();

                temp = currentSheet.Range[currentSheet.Cells[frameTable.Row + i, frameTable.Column + 11],
                    currentSheet.Cells[frameTable.Row + i, frameTable.Column + 12]];
                temp.Merge();
            }


            temp = currentSheet.Range[currentSheet.Cells[frameTable.Row, frameTable.Column + 13],
                currentSheet.Cells[frameTable.Row + 2, frameTable.Column + 33]];
            temp.Merge();
            temp.Font.Size = 20;

            temp = currentSheet.Range[currentSheet.Cells[frameTable.Row + 3, frameTable.Column + 13],
                currentSheet.Cells[frameTable.Row + 7, frameTable.Column + 24]];
            temp.Merge();
            temp.Font.Size = 12;

            temp = currentSheet.Range[currentSheet.Cells[frameTable.Row + 5, frameTable.Column + 25],
                currentSheet.Cells[frameTable.Row + 7, frameTable.Column + 33]];
            temp.Merge();

            temp = currentSheet.Range[currentSheet.Cells[frameTable.Row + 3, frameTable.Column + 25],
                currentSheet.Cells[frameTable.Row + 7, frameTable.Column + 33]];
            temp.Font.Size = 10;


            temp = currentSheet.Cells[frameTable.Row + 3, frameTable.Column + 25];
            temp.Value2 = "Лит.";

            temp = currentSheet.Cells[frameTable.Row + 3, frameTable.Column + 28];
            temp.Value2 = "Лист";

            temp = currentSheet.Cells[frameTable.Row + 3, frameTable.Column + 31];
            temp.Value2 = "Листов";

            temp = currentSheet.Range[currentSheet.Cells[frameTable.Row + 3, frameTable.Column + 25],
                currentSheet.Cells[frameTable.Row + 3, frameTable.Column + 27]];
            temp.Merge();

            for (int i = 3; i <= 4; i++)
            {
                temp = currentSheet.Range[currentSheet.Cells[frameTable.Row + i, frameTable.Column + 28],
                    currentSheet.Cells[frameTable.Row + i, frameTable.Column + 30]];
                temp.Merge();

                temp = currentSheet.Range[currentSheet.Cells[frameTable.Row + i, frameTable.Column + 31],
                    currentSheet.Cells[frameTable.Row + i, frameTable.Column + 33]];
                temp.Merge();
            }


            // --- ДЛЯ ДОП. РАМКИ СЛЕВА ---

            addFrameTable.Font.Size = 10;

            addFrameTable.Orientation = 90;

            temp = currentSheet.Cells[addFrameTable.Row, addFrameTable.Column];
            temp.Value2 = "Подп. и дата";

            temp = currentSheet.Cells[addFrameTable.Row + 7, addFrameTable.Column];
            temp.Value2 = "Инв. № дубл.";

            temp = currentSheet.Cells[addFrameTable.Row + 12, addFrameTable.Column];
            temp.Value2 = "Взам. инв. №";

            temp = currentSheet.Cells[addFrameTable.Row + 17, addFrameTable.Column];
            temp.Value2 = "Подп. и дата";

            temp = currentSheet.Cells[addFrameTable.Row + 24, addFrameTable.Column];
            temp.Value2 = "Инв. № подл.";

            for (int i = 0; i <= 1; i++)
            {
                temp = currentSheet.Range[currentSheet.Cells[addFrameTable.Row, addFrameTable.Column + i],
                    currentSheet.Cells[addFrameTable.Row + 6, addFrameTable.Column + i]];
                temp.Merge();

                temp = currentSheet.Range[currentSheet.Cells[addFrameTable.Row + 7, addFrameTable.Column + i],
                    currentSheet.Cells[addFrameTable.Row + 11, addFrameTable.Column + i]];
                temp.Merge();

                temp = currentSheet.Range[currentSheet.Cells[addFrameTable.Row + 12, addFrameTable.Column + i],
                    currentSheet.Cells[addFrameTable.Row + 16, addFrameTable.Column + i]];
                temp.Merge();

                temp = currentSheet.Range[currentSheet.Cells[addFrameTable.Row + 17, addFrameTable.Column + i],
                    currentSheet.Cells[addFrameTable.Row + 23, addFrameTable.Column + i]];
                temp.Merge();

                temp = currentSheet.Range[currentSheet.Cells[addFrameTable.Row + 24, addFrameTable.Column + i],
                    currentSheet.Cells[addFrameTable.Row + 28, addFrameTable.Column + i]];
                temp.Merge();
            }

            #endregion

            pages.Add(new Page()
            {
                allPage = allPage,
                partList = partList,
                partTable = partTable,
                frameTable = frameTable,
                addFrameTable = addFrameTable
            });
        }

        void GenerateEmptyTable(Excel.Worksheet currentSheet)
        {
            var startCell = currentSheet.Cells[1 + 53 * pages.Count, 1];

            Excel.Range temp;   // переменная для промежуточного хранения диапазонов, т.к. невозможна обработка операции в одной строке
                                // приходится разбить на 2 :(

            // основные диапазоны страницы относительно начальной клетки
            Excel.Range allPage = currentSheet.Range[startCell,
                currentSheet.Cells[startCell.Row + 52, startCell.Column + 36]]; // страница целиком

            Excel.Range partList = currentSheet.Range[currentSheet.Cells[startCell.Row, startCell.Column + 3],
                currentSheet.Cells[startCell.Row + 49, startCell.Column + 36]]; // только таблица

            Excel.Range partTable = currentSheet.Range[currentSheet.Cells[startCell.Row + 3, startCell.Column + 3],
                currentSheet.Cells[startCell.Row + 49, startCell.Column + 36]]; // таблица без шапки (только список)

            Excel.Range frameTable = currentSheet.Range[currentSheet.Cells[startCell.Row + 50, startCell.Column + 3],
                currentSheet.Cells[startCell.Row + 52, startCell.Column + 36]]; // основная подпись

            Excel.Range addFrameTable = null;  // дополнительная таблица слева

            #region Создание сетки (границ) рамки и таблицы относительно диапазонов
            partList.Borders.Weight = Excel.XlBorderWeight.xlMedium;                              // внутренние границы таблицы
            partList.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick); // толстый контур таблицы

            frameTable.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;                          // внутренние границы подписи
            frameTable.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick); // толстый контур подписи


            // много мелких границ для подписи
            temp = currentSheet.Range[currentSheet.Cells[frameTable.Row, frameTable.Column + 1],
                currentSheet.Cells[frameTable.Row + 2, frameTable.Column + 2]];
            temp.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            temp = currentSheet.Range[currentSheet.Cells[frameTable.Row, frameTable.Column + 3],
                currentSheet.Cells[frameTable.Row + 2, frameTable.Column + 7]];
            temp.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            temp = currentSheet.Range[currentSheet.Cells[frameTable.Row, frameTable.Column + 11],
                currentSheet.Cells[frameTable.Row + 2, frameTable.Column + 12]];
            temp.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);


            temp = currentSheet.Range[currentSheet.Cells[frameTable.Row + 2, frameTable.Column],
                currentSheet.Cells[frameTable.Row + 2, frameTable.Column + 12]];
            temp.Borders.Weight = Excel.XlBorderWeight.xlThick;


            temp = currentSheet.Range[currentSheet.Cells[frameTable.Row, frameTable.Column + 32],
                currentSheet.Cells[frameTable.Row + 2, frameTable.Column + 33]];
            temp.Borders.Weight = Excel.XlBorderWeight.xlThick;
            #endregion

            #region Объединение и заполнение ячеек

            // --- ДЛЯ СПИСКА КОМПОНЕНТОВ ---

            temp = currentSheet.Cells[partList.Row, partList.Column];
            temp.Value2 = "Поз. обозна-чение";
            temp = currentSheet.Range[currentSheet.Cells[partList.Row, partList.Column],
                currentSheet.Cells[partList.Row + 2, partList.Column + 2]];
            temp.WrapText = true;
            temp.Merge();

            temp = currentSheet.Cells[partList.Row, partList.Column + 3];
            temp.Value2 = "Наименование";
            temp = currentSheet.Range[currentSheet.Cells[partList.Row, partList.Column + 3],
                currentSheet.Cells[partList.Row + 2, partList.Column + 24]];
            temp.Merge();

            temp = currentSheet.Cells[partList.Row, partList.Column + 25];
            temp.Value2 = "Кол.";
            temp = currentSheet.Range[currentSheet.Cells[partList.Row, partList.Column + 25],
                currentSheet.Cells[partList.Row + 2, partList.Column + 26]];
            temp.Merge();

            temp = currentSheet.Cells[partList.Row, partList.Column + 27];
            temp.Value2 = "Примечание";
            temp = currentSheet.Range[currentSheet.Cells[partList.Row, partList.Column + 27],
                currentSheet.Cells[partList.Row + 2, partList.Column + 33]];
            temp.Merge();

            for (int i = 3; i <= 49; i++) // строки таблицы
            {
                temp = currentSheet.Range[currentSheet.Cells[partList.Row + i, partList.Column],
                    currentSheet.Cells[partList.Row + i, partList.Column + 2]];
                temp.Merge();

                temp = currentSheet.Range[currentSheet.Cells[partList.Row + i, partList.Column + 3],
                    currentSheet.Cells[partList.Row + i, partList.Column + 24]];
                temp.Merge();

                temp = currentSheet.Range[currentSheet.Cells[partList.Row + i, partList.Column + 25],
                    currentSheet.Cells[partList.Row + i, partList.Column + 26]];
                temp.Merge();

                temp = currentSheet.Range[currentSheet.Cells[partList.Row + i, partList.Column + 27],
                    currentSheet.Cells[partList.Row + i, partList.Column + 33]];
                temp.Merge();
            }

            // выравнивание по левому краю для наименований и примечаний
            temp = currentSheet.Range[currentSheet.Cells[partList.Row + 3, partList.Column + 3],
                currentSheet.Cells[partList.Row + 49, partList.Column + 24]];
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            temp = currentSheet.Range[currentSheet.Cells[partList.Row + 3, partList.Column + 27],
                currentSheet.Cells[partList.Row + 49, partList.Column + 33]];
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;


            // --- ДЛЯ РАМКИ (ПОДПИСЬ) ---

            temp = currentSheet.Range[currentSheet.Cells[frameTable.Row, frameTable.Column],
                currentSheet.Cells[frameTable.Row + 2, frameTable.Column + 12]];
            temp.Font.Size = 10;

            temp = currentSheet.Cells[frameTable.Row + 2, frameTable.Column];
            temp.Value2 = "Изм.";

            temp = currentSheet.Cells[frameTable.Row + 2, frameTable.Column + 1];
            temp.Value2 = "Лист";

            for (int i = 0; i <= 2; i++)
            {
                temp = currentSheet.Range[currentSheet.Cells[frameTable.Row + i, frameTable.Column + 1],
                    currentSheet.Cells[frameTable.Row + i, frameTable.Column + 2]];
                temp.Merge();
            }

            temp = currentSheet.Cells[frameTable.Row + 2, frameTable.Column + 3];
            temp.Value2 = "№ докум.";

            temp = currentSheet.Cells[frameTable.Row + 2, frameTable.Column + 8];
            temp.Value2 = "Подп.";

            temp = currentSheet.Cells[frameTable.Row + 2, frameTable.Column + 11];
            temp.Value2 = "Дата";

            for (int i = 0; i <= 2; i++)
            {
                temp = currentSheet.Range[currentSheet.Cells[frameTable.Row + i, frameTable.Column + 3],
                    currentSheet.Cells[frameTable.Row + i, frameTable.Column + 7]];
                temp.Merge();

                temp = currentSheet.Range[currentSheet.Cells[frameTable.Row + i, frameTable.Column + 8],
                    currentSheet.Cells[frameTable.Row + i, frameTable.Column + 10]];
                temp.Merge();

                temp = currentSheet.Range[currentSheet.Cells[frameTable.Row + i, frameTable.Column + 11],
                    currentSheet.Cells[frameTable.Row + i, frameTable.Column + 12]];
                temp.Merge();
            }


            temp = currentSheet.Range[currentSheet.Cells[frameTable.Row, frameTable.Column + 13],
                currentSheet.Cells[frameTable.Row + 2, frameTable.Column + 31]];
            temp.Merge();
            temp.Font.Size = 20;


            temp = currentSheet.Cells[frameTable.Row, frameTable.Column + 32];
            temp.Value2 = "Лист";
            temp = currentSheet.Range[currentSheet.Cells[frameTable.Row, frameTable.Column + 32],
                currentSheet.Cells[frameTable.Row, frameTable.Column + 33]];
            temp.Merge();
            temp.Font.Size = 10;

            temp = currentSheet.Range[currentSheet.Cells[frameTable.Row + 1, frameTable.Column + 32],
                currentSheet.Cells[frameTable.Row + 2, frameTable.Column + 33]];
            temp.Merge();
            temp.Font.Size = 20;
            #endregion

            pages.Add(new Page()
            {
                allPage = allPage,
                partList = partList,
                partTable = partTable,
                frameTable = frameTable,
                addFrameTable = addFrameTable
            });
        }

        List<Part> GeneratePartList(Excel.Worksheet currentSheet, int tableWidth, int tableHeight)
        {
            string[] header = new string[tableWidth];   // имена столбцов таблицы

            List<Part> listParts = new List<Part>();    // список всех компонентов

            var defaultRange = currentSheet.Range[currentSheet.Cells[1000, 1000],
                currentSheet.Cells[1000 + tableHeight - 1, 1000]];

            #region Объявление Excel Range
            Excel.Range rAvailability = defaultRange,   // наличие компонента
                rCurrent = defaultRange,                // (?)
                rDescription = defaultRange,            // словесное описание
                rDistributor = defaultRange,            // фирма поставщик
                rDistributorPartNumber = defaultRange,  // обозначение у поставщика
                rESR = defaultRange,                    //
                rItemNumber = defaultRange,             // номер в схеме
                rManufacturer = defaultRange,           // фирма изготовитель
                rManufacturerPartNumber = defaultRange, // обозначение у производителя
                rModel = defaultRange,                  // модель
                rPackage = defaultRange,                // (?)
                rPartName = defaultRange,               // имя группы к которой принадлежит элемент (конденсатор, резистор и пр.)
                rPartNumber = defaultRange,             // номер в каталоге производителя
                rPartReference = defaultRange,          // имя компонента
                rPartType = defaultRange,               // тип компонента
                rPcbFootprint = defaultRange,           // наименование корпуса
                rPower = defaultRange,                  // мощность
                rPrice = defaultRange,                  // цена
                rQuantity = defaultRange,               // количество
                rSchematicPart = defaultRange,          // наименование в библиотеке
                rSpeed = defaultRange,                  // (?)
                rT = defaultRange,                      // (?)
                rTemperature = defaultRange,            // температура
                rTKE = defaultRange,                    // температурный коэффициент ёмкости конденсаторов
                rTolerance = defaultRange,              // допустимое отклонение от номинала
                rValue = defaultRange,                  // номинал
                rValueBOM = defaultRange,               // русский номинал для таблицы
                rValueRus = defaultRange,               // русский номинал для схем
                rVoltage = defaultRange;                // вольтаж
            #endregion

            for (int i = 1; i < header.Length; i++)         // считывание заголовков в массив
            {
                if (currentSheet.Cells[1, i].Value2 != null)
                {
                    header[i] = currentSheet.Cells[1, i].Value2.ToString();
                }
            }

            for (int i = 1; i < header.Length; i++)         // поиск столбцов с определенными заголовками
            {
                var columnRange = currentSheet.Range[currentSheet.Cells[2, i], currentSheet.Cells[tableHeight, i]];

                #region Запись диапазонов в Excel Range
                switch (header[i])
                {
                    case "Availability":
                        rAvailability = columnRange;
                        break;
                    case "Current":
                        rCurrent = columnRange;
                        break;
                    case "Description":
                        rDescription = columnRange;
                        break;
                    case "Distributor":
                        rDistributor = columnRange;
                        break;
                    case "Distributor Part Number":
                        rDistributorPartNumber = columnRange;
                        break;
                    case "ESR":
                        rESR = columnRange;
                        break;
                    case "Item Number":
                        rItemNumber = columnRange;
                        break;
                    case "Manufacturer":
                        rManufacturer = columnRange;
                        break;
                    case "Manufacturer Part Number":
                        rManufacturerPartNumber = columnRange;
                        break;
                    case "Model":
                        rModel = columnRange;
                        break;
                    case "Package":
                        rPackage = columnRange;
                        break;
                    case "Part Name":
                        rPartName = columnRange;
                        break;
                    case "Part Number":
                        rPartNumber = columnRange;
                        break;
                    case "Part Reference":
                        rPartReference = columnRange;
                        break;
                    case "Part Type":
                        rPartType = columnRange;
                        break;
                    case "PCB Footprint":
                        rPcbFootprint = columnRange;
                        break;
                    case "Power":
                        rPower = columnRange;
                        break;
                    case "Price":
                        rPrice = columnRange;
                        break;
                    case "Quantity":
                        rQuantity = columnRange;
                        break;
                    case "Schematic Part":
                        rSchematicPart = columnRange;
                        break;
                    case "Speed":
                        rSpeed = columnRange;
                        break;
                    case "T":
                        rT = columnRange;
                        break;
                    case "Temperature":
                        rTemperature = columnRange;
                        break;
                    case "TKE":
                        rTKE = columnRange;
                        break;
                    case "Tolerance":
                        rTolerance = columnRange;
                        break;
                    case "Value":
                        rValue = columnRange;
                        break;
                    case "ValueBOM":
                        rValueBOM = columnRange;
                        break;
                    case "ValueRus":
                        rValueRus = columnRange;
                        break;
                    case "Voltage":
                        rVoltage = columnRange;
                        break;
                }
                #endregion
            }

            for (int i = 1; i < tableHeight; i++)               // создание списка (List) компонентов
            {
                if (rPartReference.Cells[i, 1].Value2 != null)
                {
                    listParts.Add(new Part()
                    {
                        PartReference = rPartReference.Cells[i, 1].Value2.ToString(),

                        #region Запись прочих свойств компонента
                        Availability = (rAvailability.Cells[i, 1].Value2 != null) ? rAvailability.Cells[i, 1].Value2.ToString() : "",
                        Current = (rCurrent.Cells[i, 1].Value2 != null) ? rCurrent.Cells[i, 1].Value2.ToString() : "",
                        Description = (rDescription.Cells[i, 1].Value2 != null) ? rDescription.Cells[i, 1].Value2.ToString() : "",
                        Distributor = (rDistributor.Cells[i, 1].Value2 != null) ? rDistributor.Cells[i, 1].Value2.ToString() : "",
                        DistributorPartNumber = (rDistributorPartNumber.Cells[i, 1].Value2 != null) ? rDistributorPartNumber.Cells[i, 1].Value2.ToString() : "",
                        ESR = (rESR.Cells[i, 1].Value2 != null) ? rESR.Cells[i, 1].Value2.ToString() : "",
                        ItemNumber = (rItemNumber.Cells[i, 1].Value2 != null) ? rItemNumber.Cells[i, 1].Value2.ToString() : "",
                        Manufacturer = (rManufacturer.Cells[i, 1].Value2 != null) ? rManufacturer.Cells[i, 1].Value2.ToString() : "",
                        ManufacturerPartNumber = (rManufacturerPartNumber.Cells[i, 1].Value2 != null) ? rManufacturerPartNumber.Cells[i, 1].Value2.ToString() : "",
                        Model = (rModel.Cells[i, 1].Value2 != null) ? rModel.Cells[i, 1].Value2.ToString() : "",
                        Package = (rPackage.Cells[i, 1].Value2 != null) ? rPackage.Cells[i, 1].Value2.ToString() : "",
                        PartName = (rPartName.Cells[i, 1].Value2 != null) ? rPartName.Cells[i, 1].Value2.ToString() : "",
                        PartNumber = (rPartNumber.Cells[i, 1].Value2 != null) ? rPartNumber.Cells[i, 1].Value2.ToString() : "",
                        PartType = (rPartType.Cells[i, 1].Value2 != null) ? rPartType.Cells[i, 1].Value2.ToString() : "",
                        PcbFootprint = (rPcbFootprint.Cells[i, 1].Value2 != null) ? rPcbFootprint.Cells[i, 1].Value2.ToString() : "",
                        Power = (rPower.Cells[i, 1].Value2 != null) ? rPower.Cells[i, 1].Value2.ToString() : "",
                        Price = (rPrice.Cells[i, 1].Value2 != null) ? rPrice.Cells[i, 1].Value2.ToString() : "",
                        Quantity = (rQuantity.Cells[i, 1].Value2 != null) ? rQuantity.Cells[i, 1].Value2.ToString() : "",
                        SchematicPart = (rSchematicPart.Cells[i, 1].Value2 != null) ? rSchematicPart.Cells[i, 1].Value2.ToString() : "",
                        Speed = (rSpeed.Cells[i, 1].Value2 != null) ? rSpeed.Cells[i, 1].Value2.ToString() : "",
                        T = (rT.Cells[i, 1].Value2 != null) ? rT.Cells[i, 1].Value2.ToString() : "",
                        Temperature = (rTemperature.Cells[i, 1].Value2 != null) ? rTemperature.Cells[i, 1].Value2.ToString() : "",
                        TKE = (rTKE.Cells[i, 1].Value2 != null) ? rTKE.Cells[i, 1].Value2.ToString() : "",
                        Tolerance = (rTolerance.Cells[i, 1].Value2 != null) ? rTolerance.Cells[i, 1].Value2.ToString() : "",
                        Voltage = (rVoltage.Cells[i, 1].Value2 != null) ? rVoltage.Cells[i, 1].Value2.ToString() : "",
                        ValueBOM = (rValueBOM.Cells[i, 1].Value2 != null) ? rValueBOM.Cells[i, 1].Value2.ToString() : "",
                        ValueRus = (rValueRus.Cells[i, 1].Value2 != null) ? rValueRus.Cells[i, 1].Value2.ToString() : "",
                        #endregion

                        Value = (rValue.Cells[i, 1].Value2 != null) ? rValue.Cells[i, 1].Value2.ToString() : "",
                        group = 0
                    });
                }
            }

            return listParts;
        }

        void CreateGroupForParts(List<Part> partList)
        {
            int groupCounter = 0;

            for (int i = 0; i < partList.Count; i++)
            {
                if (partList[i].group == 0)
                {
                    groupCounter++;

                    var part = partList[i];     // делаем копию
                    part.group = groupCounter;  // выполняем изменение
                    partList[i] = part;         // записываем
                    //эти 3 строчки делают то же самое что и
                    //partList[i].group = groupCounter;

                    for (int j = i + 1; j < partList.Count; j++)
                    {
                        if (partList[j].group == 0) // сравнение элементов
                        {
                            if ((partList[i].PartType == partList[j].PartType) &&
                                (partList[i].PartName == partList[j].PartName))
                            {

                                part = partList[j];             // делаем копию
                                part.group = partList[i].group; // выполняем изменение
                                partList[j] = part;             // записываем
                                //эти 3 строчки делают то же самое что и
                                //partList[j].group = partList[i].group;
                            }
                        }
                    }
                }
            }
        }

        void GenerateTable(List<Part> partList, Excel.Worksheet currentSheet)
        {
            CreateGroupForParts(partList);              // формирование групп компонентов

            PrepareSheetForTable(currentSheet);         // подготовка текущего листа книги для создания таблицы
            GenerateEmptyTable(currentSheet, true);     // создание пустой таблицы (ГОСТ)

            //partList.Sort(delegate (Part p1, Part p2)   // сортировка по номеру группы
            //{
            //    return p1.group.CompareTo(p2.group);
            //});

            // сортировка по номеру группы и обозначению компонента
            var sorted = partList.OrderBy(x => x.group).ThenBy(x => x.PartReference).ToList(); 
            partList = sorted;

            bool takeOutTitle = false;
            bool printedTitle = false;

            int l = 0; int r = 1;                      // l - счетчик страниц, r - счетчик строк
            for (int i = 0; i < partList.Count; i++)   // проход по всем элементам списка
            {
                if (i != 0)
                    if (partList[i].group != partList[i - 1].group)
                    {
                        DropDown(pages[l].partTable.Cells[r, 1], currentSheet);
                        r++;
                        takeOutTitle = false;
                        printedTitle = false;
                    }

                if ((i != partList.Count - 1)&&(!printedTitle))
                    if (partList[i].group == partList[i + 1].group)
                    {
                        DropDown(pages[l].partTable.Cells[r, 1], currentSheet);
                        r++;
                        takeOutTitle = true;
                    }

                if (r > pages[l].partTable.Rows.Count)  // если в таблицу не помещается 
                {
                    GenerateEmptyTable(currentSheet);   // создаем новую страницу
                    l++;                                // инкрементируем номер листа
                    r = 1;                              // обнуляем счетчик строки 
                }

                #region Заполнение столбцов таблицы

                if (!takeOutTitle)
                {
                    if (partList[i].PartName != "")                                     // Part Name
                    {
                        pages[l].partTable.Cells[r, 4].Value2 += partList[i].PartName;
                        pages[l].partTable.Cells[r, 4].Value2 += " ";
                    }

                    pages[l].partTable.Cells[r, 4].Value2 += partList[i].PartType;      // Part Type
                    pages[l].partTable.Cells[r, 4].Value2 += " ";
                }

                else
                {
                    if (!printedTitle)
                    {
                        if (partList[i].PartName != "")                                     // Part Name
                        {
                            pages[l].partTable.Cells[r - 1, 4].Value2 += partList[i].PartName;
                            pages[l].partTable.Cells[r - 1, 4].Value2 += " ";
                        }

                        pages[l].partTable.Cells[r - 1, 4].Value2 += partList[i].PartType;      // Part Type

                        printedTitle = true;
                    }
                }

                bool comma = false;

                CheckComma(pages[l].partTable.Cells[r, 4], partList[i].Manufacturer, ref comma);   // Manufacturer
                CheckComma(pages[l].partTable.Cells[r, 4], partList[i].Model, ref comma);          // Model
                CheckComma(pages[l].partTable.Cells[r, 4], partList[i].ValueBOM, ref comma);       // ValueBOM
                CheckComma(pages[l].partTable.Cells[r, 4], partList[i].Tolerance, ref comma);      // Tolerance


                pages[l].partTable.Cells[r, 1].Value2 = partList[i].PartReference;  // позиционное обозначение
                pages[l].partTable.Cells[r, 26].Value2 = partList[i].Quantity;      // количество эл-тов в третий столбик


                comma = false; // позволяет избежать лишней запятой в начале строки

                CheckComma(pages[l].partTable.Cells[r, 28], partList[i].Description, ref comma);// Description
                CheckComma(pages[l].partTable.Cells[r, 28], partList[i].Voltage, ref comma);    // Voltage
                CheckComma(pages[l].partTable.Cells[r, 28], partList[i].Power, ref comma);      // Power
                CheckComma(pages[l].partTable.Cells[r, 28], partList[i].TKE, ref comma);        // TKE
                CheckComma(pages[l].partTable.Cells[r, 28], partList[i].ESR, ref comma);        // ESR
                CheckComma(pages[l].partTable.Cells[r, 28], partList[i].Temperature, ref comma);// Temperature
                #endregion

                r++;    // инкремент номера строки
            }

        }

        void CheckComma(Excel.Range cell, string val)
        {
            if (val != "")
            {
                cell.Value2 += ", ";
                cell.Value2 += val;
            }
        }
        void CheckComma(Excel.Range cell, string val, ref bool comma)
        {
            if (val != "")
            {
                if (comma) cell.Value2 += ", ";
                cell.Value2 += val;
                comma = true;
            }
        }

        void DropDown(Excel.Range activeCell, Excel.Worksheet currentSheet)
        {
            if (pages.Count != 0)
            {
                int rowNum = activeCell.Row;        // номер строки выделенной ячейки
                int rowNumInPage = -1;              // номер строки выделенной ячейки в СК её страницы
                int pageNum = -1;                   // сюда запишем номер страницы с выделенной ячейкой

                int endPageNum = -1;                // номер последней страницы 
                int endRowNum = -1;                 // номер последней строки
                int endRowNumInPage = -1;           // номер последней строки в СК её страницы

                for (int i = 0; i < pages.Count; i++)                           // определим какой странице принадлежит выделенная ячейка
                {
                    if ((rowNum >= pages[i].partTable.Row) &&
                        (rowNum < pages[i].partTable.Row + pages[i].partTable.Rows.Count))
                    {
                        pageNum = i;
                        rowNumInPage = rowNum - (pages[i].partTable.Row - 1);   // находим смещение данного листа (списка) относительно 
                                                                                // A1 и вычитаем его из координат активной ячейки
                        break;
                    }
                }

                if (pageNum > -1)                                               // если активная ячейка принадлежит какой-нибудь странице
                {
                    bool _break = false;
                    for (int i = pages.Count - 1; i >= 0; i--)                  // найдем последнюю заполненную ячейку
                    {
                        for (int j = pages[i].partTable.Rows.Count; j > 0; j--)
                        {
                            if ((pages[i].partTable[j, 1].Value2 != null) || (pages[i].partTable[j, 4].Value2 != null)) // если ячейка "поз. обозначение" или "наименование" не пуста
                            {
                                endPageNum = i;
                                endRowNumInPage = j;                            // в системе координат страницы
                                endRowNum = j + (pages[i].partTable.Row - 1);   // в системе координат документа
                                _break = true;
                                break;
                            }
                        }
                        if (_break) break;
                    }

                    if (endPageNum > -1)
                    {
                        if (endPageNum == pageNum)  // опустим содержимое вниз (при условии что не происходит межстраничного сдвига)
                        {
                            if (endRowNumInPage == pages[pageNum].partTable.Rows.Count)
                            {
                                if (endPageNum == pages.Count - 1)
                                {
                                    GenerateEmptyTable(currentSheet);
                                }

                                pages[pageNum + 1].partTable[1, 1].Value2 = pages[pageNum].partTable[endRowNumInPage, 1].Value2;
                                pages[pageNum + 1].partTable[1, 4].Value2 = pages[pageNum].partTable[endRowNumInPage, 4].Value2;
                                pages[pageNum + 1].partTable[1, 26].Value2 = pages[pageNum].partTable[endRowNumInPage, 26].Value2;
                                pages[pageNum + 1].partTable[1, 28].Value2 = pages[pageNum].partTable[endRowNumInPage, 28].Value2;

                                for (int i = endRowNumInPage - 1; i >= rowNumInPage; i--)
                                {
                                    pages[pageNum].partTable[i + 1, 1].Value2 = pages[pageNum].partTable[i, 1].Value2;
                                    pages[pageNum].partTable[i + 1, 4].Value2 = pages[pageNum].partTable[i, 4].Value2;
                                    pages[pageNum].partTable[i + 1, 26].Value2 = pages[pageNum].partTable[i, 26].Value2;
                                    pages[pageNum].partTable[i + 1, 28].Value2 = pages[pageNum].partTable[i, 28].Value2;
                                }

                            }

                            else
                            {
                                for (int i = endRowNumInPage; i >= rowNumInPage; i--)
                                {
                                    pages[pageNum].partTable[i + 1, 1].Value2 = pages[pageNum].partTable[i, 1].Value2;
                                    pages[pageNum].partTable[i + 1, 4].Value2 = pages[pageNum].partTable[i, 4].Value2;
                                    pages[pageNum].partTable[i + 1, 26].Value2 = pages[pageNum].partTable[i, 26].Value2;
                                    pages[pageNum].partTable[i + 1, 28].Value2 = pages[pageNum].partTable[i, 28].Value2;
                                }
                            }

                            pages[pageNum].partTable[rowNumInPage, 1].Value2 = "";     // удаление содержимого освободившейся строки
                            pages[pageNum].partTable[rowNumInPage, 4].Value2 = "";
                            pages[pageNum].partTable[rowNumInPage, 26].Value2 = "";
                            pages[pageNum].partTable[rowNumInPage, 28].Value2 = "";
                        }

                        else            // межстраничный сдвиг происходит
                        {
                            if (endRowNumInPage == pages[pages.Count - 1].partTable.Rows.Count)
                            {
                                if (endPageNum == pages.Count - 1)
                                {
                                    GenerateEmptyTable(currentSheet);
                                }

                                pages[endPageNum + 1].partTable[1, 1].Value2 = pages[endPageNum].partTable[endRowNumInPage, 1].Value2;
                                pages[endPageNum + 1].partTable[1, 4].Value2 = pages[endPageNum].partTable[endRowNumInPage, 4].Value2;
                                pages[endPageNum + 1].partTable[1, 26].Value2 = pages[endPageNum].partTable[endRowNumInPage, 26].Value2;
                                pages[endPageNum + 1].partTable[1, 28].Value2 = pages[endPageNum].partTable[endRowNumInPage, 28].Value2;

                                for (int i = endRowNumInPage - 1; i > 0; i--)       // перенос с нижней страницы
                                {
                                    pages[endPageNum].partTable[i + 1, 1].Value2 = pages[endPageNum].partTable[i, 1].Value2;
                                    pages[endPageNum].partTable[i + 1, 4].Value2 = pages[endPageNum].partTable[i, 4].Value2;
                                    pages[endPageNum].partTable[i + 1, 26].Value2 = pages[endPageNum].partTable[i, 26].Value2;
                                    pages[endPageNum].partTable[i + 1, 28].Value2 = pages[endPageNum].partTable[i, 28].Value2;
                                }
                            }

                            else
                            {
                                for (int i = endRowNumInPage; i > 0; i--)       // перенос с нижней страницы
                                {
                                    pages[endPageNum].partTable[i + 1, 1].Value2 = pages[endPageNum].partTable[i, 1].Value2;
                                    pages[endPageNum].partTable[i + 1, 4].Value2 = pages[endPageNum].partTable[i, 4].Value2;
                                    pages[endPageNum].partTable[i + 1, 26].Value2 = pages[endPageNum].partTable[i, 26].Value2;
                                    pages[endPageNum].partTable[i + 1, 28].Value2 = pages[endPageNum].partTable[i, 28].Value2;
                                }
                            }
                            
                            var temp = pages[endPageNum - 1].partTable;     // перенос последней строки предпоследней страницы на последнюю

                            pages[endPageNum].partTable[1, 1].Value2 = temp[temp.Rows.Count, 1].Value2;
                            pages[endPageNum].partTable[1, 4].Value2 = temp[temp.Rows.Count, 4].Value2;
                            pages[endPageNum].partTable[1, 26].Value2 = temp[temp.Rows.Count, 26].Value2;
                            pages[endPageNum].partTable[1, 28].Value2 = temp[temp.Rows.Count, 28].Value2;

                            for (int i = endPageNum - 1; i > pageNum; i--)  // перенос на промежуточных страницах
                            {
                                for (int j = pages[i].partTable.Rows.Count - 1; j > 0; j--)
                                {
                                    pages[i].partTable[j + 1, 1].Value2 = pages[i].partTable[j, 1].Value2;
                                    pages[i].partTable[j + 1, 4].Value2 = pages[i].partTable[j, 4].Value2;
                                    pages[i].partTable[j + 1, 26].Value2 = pages[i].partTable[j, 26].Value2;
                                    pages[i].partTable[j + 1, 28].Value2 = pages[i].partTable[j, 28].Value2;
                                }

                                temp = pages[i - 1].partTable;

                                pages[i].partTable[1, 1].Value2 = temp[temp.Rows.Count, 1].Value2;
                                pages[i].partTable[1, 4].Value2 = temp[temp.Rows.Count, 4].Value2;
                                pages[i].partTable[1, 26].Value2 = temp[temp.Rows.Count, 26].Value2;
                                pages[i].partTable[1, 28].Value2 = temp[temp.Rows.Count, 28].Value2;
                            }

                            for (int i = pages[pageNum].partTable.Rows.Count - 1; i >= rowNumInPage; i--)   // перенос на первой странице
                            {
                                pages[pageNum].partTable[i + 1, 1].Value2 = pages[pageNum].partTable[i, 1].Value2;
                                pages[pageNum].partTable[i + 1, 4].Value2 = pages[pageNum].partTable[i, 4].Value2;
                                pages[pageNum].partTable[i + 1, 26].Value2 = pages[pageNum].partTable[i, 26].Value2;
                                pages[pageNum].partTable[i + 1, 28].Value2 = pages[pageNum].partTable[i, 28].Value2;
                            }

                            pages[pageNum].partTable[rowNumInPage, 1].Value2 = "";     // удаление содержимого освободившейся строки
                            pages[pageNum].partTable[rowNumInPage, 4].Value2 = "";
                            pages[pageNum].partTable[rowNumInPage, 26].Value2 = "";
                            pages[pageNum].partTable[rowNumInPage, 28].Value2 = "";
                        }
                    }
                }
            }

        }
        void DropUp(Excel.Range activeCell)
        {
            if (pages.Count != 0)
            {
                int rowNum = activeCell.Row;        // номер строки выделенной ячейки
                int rowNumInPage = -1;              // номер строки выделенной ячейки в СК её страницы
                int pageNum = -1;                   // сюда запишем номер страницы с выделенной ячейкой

                int endPageNum = -1;                // номер последней страницы 
                int endRowNum = -1;                 // номер последней строки
                int endRowNumInPage = -1;           // номер последней строки в СК её страницы

                for (int i = 0; i < pages.Count; i++)                           // определим какой странице принадлежит выделенная ячейка
                {
                    if ((rowNum >= pages[i].partTable.Row) &&
                        (rowNum < pages[i].partTable.Row + pages[i].partTable.Rows.Count))
                    {
                        pageNum = i;
                        rowNumInPage = rowNum - (pages[i].partTable.Row - 1);   // находим смещение данного листа (списка) относительно 
                                                                                // A1 и вычитаем его из координат активной ячейки
                        break;
                    }
                }

                if (pageNum > -1)                                               // если активная ячейка принадлежит какой-нибудь странице
                {
                    bool _break = false;
                    for (int i = pages.Count - 1; i >= 0; i--)                  // найдем последнюю заполненную ячейку
                    {
                        for (int j = pages[i].partTable.Rows.Count; j > 0; j--)
                        {
                            if ((pages[i].partTable[j, 1].Value2 != null)||(pages[i].partTable[j, 4].Value2 != null)) // если ячейка "поз. обозначение" или "наименование" не пуста
                            {
                                endPageNum = i;
                                endRowNumInPage = j;                            // в системе координат страницы
                                endRowNum = j + (pages[i].partTable.Row - 1);   // в системе координат документа
                                _break = true;
                                break;
                            }
                        }
                        if (_break) break;
                    }
                }

                if (endPageNum > -1)
                {
                    if (endPageNum == pageNum)  // поднимем содержимое вверх (при условии что не происходит межстраничного сдвига)
                    {
                        for (int i = rowNumInPage; i < endRowNumInPage; i++)
                        {
                            pages[pageNum].partTable[i, 1].Value2 = pages[pageNum].partTable[i + 1, 1].Value2;
                            pages[pageNum].partTable[i, 4].Value2 = pages[pageNum].partTable[i + 1, 4].Value2;
                            pages[pageNum].partTable[i, 26].Value2 = pages[pageNum].partTable[i + 1, 26].Value2;
                            pages[pageNum].partTable[i, 28].Value2 = pages[pageNum].partTable[i + 1, 28].Value2;
                        }

                        pages[pageNum].partTable[endRowNumInPage, 1].Value2 = "";     // удаление содержимого последней строки
                        pages[pageNum].partTable[endRowNumInPage, 4].Value2 = "";
                        pages[pageNum].partTable[endRowNumInPage, 26].Value2 = "";
                        pages[pageNum].partTable[endRowNumInPage, 28].Value2 = "";
                    }

                    else
                    {
                        for (int i = rowNumInPage; i < pages[pageNum].partTable.Rows.Count; i++)       // перенос с первой страницы
                        {
                            pages[pageNum].partTable[i, 1].Value2 = pages[pageNum].partTable[i + 1, 1].Value2;
                            pages[pageNum].partTable[i, 4].Value2 = pages[pageNum].partTable[i + 1, 4].Value2;
                            pages[pageNum].partTable[i, 26].Value2 = pages[pageNum].partTable[i + 1, 26].Value2;
                            pages[pageNum].partTable[i, 28].Value2 = pages[pageNum].partTable[i + 1, 28].Value2;
                        }

                        var temp = pages[pageNum].partTable;     // перенос первой строки второй страницы на первую страницу

                        temp[temp.Rows.Count, 1].Value2 = pages[pageNum + 1].partTable[1, 1].Value2;
                        temp[temp.Rows.Count, 4].Value2 = pages[pageNum + 1].partTable[1, 4].Value2;
                        temp[temp.Rows.Count, 26].Value2 = pages[pageNum + 1].partTable[1, 26].Value2;
                        temp[temp.Rows.Count, 28].Value2 = pages[pageNum + 1].partTable[1, 28].Value2;

                        for (int i = pageNum + 1; i < endPageNum; i++)  // перенос на промежуточных страницах
                        {
                            for (int j = 1; j < pages[i].partTable.Rows.Count; j++)
                            {
                                pages[i].partTable[j, 1].Value2 = pages[i].partTable[j + 1, 1].Value2;
                                pages[i].partTable[j, 4].Value2 = pages[i].partTable[j + 1, 4].Value2;
                                pages[i].partTable[j, 26].Value2 = pages[i].partTable[j + 1, 26].Value2;
                                pages[i].partTable[j, 28].Value2 = pages[i].partTable[j + 1, 28].Value2;
                            }

                            temp = pages[i].partTable;

                            temp[temp.Rows.Count, 1].Value2 = pages[i + 1].partTable[1, 1];
                            temp[temp.Rows.Count, 4].Value2 = pages[i + 1].partTable[1, 4];
                            temp[temp.Rows.Count, 26].Value2 = pages[i + 1].partTable[1, 26];
                            temp[temp.Rows.Count, 28].Value2 = pages[i + 1].partTable[1, 28];
                        }

                        for (int i = 1; i < endRowNumInPage; i++)   // перенос на последней странице
                        {
                            pages[endPageNum].partTable[i, 1].Value2 = pages[endPageNum].partTable[i + 1, 1].Value2;
                            pages[endPageNum].partTable[i, 4].Value2 = pages[endPageNum].partTable[i + 1, 4].Value2;
                            pages[endPageNum].partTable[i, 26].Value2 = pages[endPageNum].partTable[i + 1, 26].Value2;
                            pages[endPageNum].partTable[i, 28].Value2 = pages[endPageNum].partTable[i + 1, 28].Value2;
                        }

                        pages[endPageNum].partTable[endRowNumInPage, 1].Value2 = "";     // удаление содержимого последней строки
                        pages[endPageNum].partTable[endRowNumInPage, 4].Value2 = "";
                        pages[endPageNum].partTable[endRowNumInPage, 26].Value2 = "";
                        pages[endPageNum].partTable[endRowNumInPage, 28].Value2 = "";
                    }
                }
            }
        }

        void RepairDocument(Excel.Worksheet activeSheet)
        {
            pages.Clear();
            
            for (int i = 0; i < 1000; i++)
            {
                if (activeSheet.Cells[1 + 53 * i, 4].Value2 == "Поз. обозна-чение")
                {
                    if (pages.Count == 0)
                    {
                        var startCell = activeSheet.Range["A1"]; // первая страница всегда будет от А1
                        
                        Excel.Range _allPage = activeSheet.Range[startCell,
                            activeSheet.Cells[startCell.Row + 52, startCell.Column + 36]]; // страница целиком

                        Excel.Range _partList = activeSheet.Range[activeSheet.Cells[startCell.Row, startCell.Column + 3],
                            activeSheet.Cells[startCell.Row + 44, startCell.Column + 36]]; // только таблица

                        Excel.Range _partTable = activeSheet.Range[activeSheet.Cells[startCell.Row + 3, startCell.Column + 3],
                            activeSheet.Cells[startCell.Row + 44, startCell.Column + 36]]; // таблица без шапки (только список)

                        Excel.Range _frameTable = activeSheet.Range[activeSheet.Cells[startCell.Row + 45, startCell.Column + 3],
                            activeSheet.Cells[startCell.Row + 52, startCell.Column + 36]]; // основная подпись

                        Excel.Range _addFrameTable = activeSheet.Range[activeSheet.Cells[startCell.Row + 24, startCell.Column + 1],
                            activeSheet.Cells[startCell.Row + 52, startCell.Column + 2]];  // дополнительная таблица слева

                        pages.Add(new Page()
                        {
                            allPage = _allPage,
                            partList = _partList,
                            partTable = _partTable,
                            frameTable = _frameTable,
                            addFrameTable = _addFrameTable
                        });
                    }
                    else
                    {
                        var startCell = activeSheet.Cells[1 + 53 * pages.Count, 1];

                        Excel.Range _allPage = activeSheet.Range[startCell,
                            activeSheet.Cells[startCell.Row + 52, startCell.Column + 36]]; // страница целиком

                        Excel.Range _partList = activeSheet.Range[activeSheet.Cells[startCell.Row, startCell.Column + 3],
                            activeSheet.Cells[startCell.Row + 49, startCell.Column + 36]]; // только таблица

                        Excel.Range _partTable = activeSheet.Range[activeSheet.Cells[startCell.Row + 3, startCell.Column + 3],
                            activeSheet.Cells[startCell.Row + 49, startCell.Column + 36]]; // таблица без шапки (только список)

                        Excel.Range _frameTable = activeSheet.Range[activeSheet.Cells[startCell.Row + 50, startCell.Column + 3],
                            activeSheet.Cells[startCell.Row + 52, startCell.Column + 36]]; // основная подпись

                        Excel.Range _addFrameTable = null;  // дополнительная таблица слева

                        pages.Add(new Page()
                        {
                            allPage = _allPage,
                            partList = _partList,
                            partTable = _partTable,
                            frameTable = _frameTable,
                            addFrameTable = _addFrameTable
                        });
                    }
                }
                else break;
            }
        }
    }
}
