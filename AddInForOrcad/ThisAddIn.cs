using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System;

namespace AddInForOrcad
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Выполнение кода происходит при загрузке надстройки VSTO.

            this.Application.SheetActivate += new Excel.AppEvents_SheetActivateEventHandler(Application_SheetActivate);
            this.Application.WindowActivate += new Excel.AppEvents_WindowActivateEventHandler(Application_WindowActivate);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Выполнения кода происходит непосредственно перед выгрузкой надстройки VSTO.
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            var ribbon = new SampleRibbon();
            ribbon.CreateTableClick += ribbon_CreateTableClicked;
            ribbon.CreateEmptyTableClick += ribbon_CreateEmptyTableClicked;
            ribbon.AddLineClick += ribbon_AddLineClicked;
            ribbon.DelLineClick += ribbon_DelLineClicked;
            return Globals.Factory.GetRibbonFactory().CreateRibbonManager(new IRibbonExtension[] { ribbon });
        }

        private void ribbon_CreateTableClicked()        // нажатие кнопки "создать таблицу"
        {
            Excel.Worksheet activeSheet = (Excel.Worksheet)Application.ActiveSheet;

            var partList = GeneratePartList(activeSheet, //создание списка компонентов
                Convert.ToInt32(Globals.Ribbons.SampleRibbon.tableWidth.Text),
                Convert.ToInt32(Globals.Ribbons.SampleRibbon.tableHeight.Text));

            GenerateTable(partList, activeSheet);   // создание таблицы ГОСТ на основе списка компонентов
        }

        private void ribbon_CreateEmptyTableClicked()   // нажатие кнопки "создать пустую таблицу"
        {
            Excel.Worksheet activeSheet = (Excel.Worksheet)Application.ActiveSheet;

            if (pages.Count == 0)                   // если страница первая
            {
                PrepareSheetForTable(activeSheet);      // подготовка страницы excel
                GenerateEmptyTable(activeSheet, true);  // создание первой пустой таблицы
            }
            else GenerateEmptyTable(activeSheet);   // создание последующих таблиц
        }

        private void ribbon_AddLineClicked()            // нажатие кнопки "добавить строку"
        {
            Excel.Worksheet activeSheet = (Excel.Worksheet)Application.ActiveSheet;

            DropDown(Application.ActiveCell, activeSheet);
        }

        private void ribbon_DelLineClicked()            // нажатие кнопки "удалить строку"
        {
            Excel.Worksheet activeSheet = (Excel.Worksheet)Application.ActiveSheet;

            DropUp(Application.ActiveCell);
        }

        void Application_SheetActivate(object Sh)
        {
            Excel.Worksheet activeSheet = (Excel.Worksheet)Application.ActiveSheet;

            RepairDocument(activeSheet);
        }

        void Application_WindowActivate(Excel.Workbook Wb, Excel.Window Wn)
        {
            Excel.Worksheet activeSheet = (Excel.Worksheet)Application.ActiveSheet;

            RepairDocument(activeSheet);
        }

        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        #region тестовые методы
        private void Sheet1_BeforeDoubleClick(Excel.Range AuthorRange, ref bool Proceed)
        {
            Excel.Worksheet activeSheet = (Excel.Worksheet)Application.ActiveSheet;

            activeSheet.Range["B2"].Value2 = "Buy VSTO Book";
            activeSheet.Range["B2"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGreen);
            activeSheet.Range["B2"].Font.Bold = true;
        }
        private void Sheet1_SelectionChange(Excel.Range AuthorCellPoint)
        {
            string AuthorCellSelection = AuthorCellPoint.get_Address
                (missing, missing, Excel.XlReferenceStyle.xlA1, missing, missing);

            Excel.Worksheet activeSheet = (Excel.Worksheet)Application.ActiveSheet;

            activeSheet.Range["B2"].Value2 = AuthorCellSelection;
        }
        public static bool EmailValidate(string strEmail)
        {
            string strRegex = @"^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}" +
            @"\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\" +
            @".)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$";
            Regex emailRegEx = new Regex(strRegex);
            if (emailRegEx.IsMatch(strEmail))
            {
                return (true);
            }
            else
            {
                return (false);
            }
        }

        #endregion
    }
}
