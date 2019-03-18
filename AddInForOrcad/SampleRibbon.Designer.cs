namespace AddInForOrcad
{
    partial class SampleRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public SampleRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.CreateTable = this.Factory.CreateRibbonButton();
            this.tableHeight = this.Factory.CreateRibbonEditBox();
            this.tableWidth = this.Factory.CreateRibbonEditBox();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.CreateEmptyTable = this.Factory.CreateRibbonButton();
            this.AddLine = this.Factory.CreateRibbonButton();
            this.DelLine = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "Для OrCAD";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.CreateTable);
            this.group1.Items.Add(this.tableHeight);
            this.group1.Items.Add(this.tableWidth);
            this.group1.Label = "Создание таблицы";
            this.group1.Name = "group1";
            // 
            // CreateTable
            // 
            this.CreateTable.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.CreateTable.Label = "Создание списка (ГОСТ)";
            this.CreateTable.Name = "CreateTable";
            this.CreateTable.OfficeImageId = "CacheListData";
            this.CreateTable.ShowImage = true;
            this.CreateTable.SuperTip = "Создание списка электронных компонентов на основе таблицы сгенерированной в OrCAD" +
    " Capture";
            this.CreateTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CreateTable_Click);
            // 
            // tableHeight
            // 
            this.tableHeight.Label = "Высота таблицы";
            this.tableHeight.Name = "tableHeight";
            this.tableHeight.SuperTip = "Количество строк таблицы (с учетом заголовка), сгенерированной программным компле" +
    "ксом OrCAD Capture";
            this.tableHeight.Text = "100";
            // 
            // tableWidth
            // 
            this.tableWidth.Label = "Ширина таблицы";
            this.tableWidth.Name = "tableWidth";
            this.tableWidth.SuperTip = "Количество столбцов таблицы, сгенерированной программным комплексом OrCAD Capture" +
    "";
            this.tableWidth.Text = "30";
            // 
            // group2
            // 
            this.group2.Items.Add(this.CreateEmptyTable);
            this.group2.Items.Add(this.AddLine);
            this.group2.Items.Add(this.DelLine);
            this.group2.Label = "Прочее";
            this.group2.Name = "group2";
            // 
            // CreateEmptyTable
            // 
            this.CreateEmptyTable.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.CreateEmptyTable.Label = "Создание пустой формы (ГОСТ)";
            this.CreateEmptyTable.Name = "CreateEmptyTable";
            this.CreateEmptyTable.OfficeImageId = "AccessFormDatasheet";
            this.CreateEmptyTable.ShowImage = true;
            this.CreateEmptyTable.SuperTip = "Создание пустой таблицы для списка электронных компонентов (ГОСТ)";
            this.CreateEmptyTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CreateEmptyTable_Click);
            // 
            // AddLine
            // 
            this.AddLine.Label = "Добавить строку";
            this.AddLine.Name = "AddLine";
            this.AddLine.OfficeImageId = "TableRowsInsertAboveExcel";
            this.AddLine.ShowImage = true;
            this.AddLine.SuperTip = "Добавление пустой строки в таблицу (относительно выделенной ячейки) со смещением " +
    "вниз всех элементов списка";
            this.AddLine.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddLine_Click);
            // 
            // DelLine
            // 
            this.DelLine.Label = "Удалить строку";
            this.DelLine.Name = "DelLine";
            this.DelLine.OfficeImageId = "TableRowsDeleteExcel";
            this.DelLine.ShowImage = true;
            this.DelLine.SuperTip = "Удаление строки из таблицы (относительно выделенной ячейки) со смещением вверх вс" +
    "ех элементов списка.";
            this.DelLine.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DelLine_Click);
            // 
            // SampleRibbon
            // 
            this.Name = "SampleRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.SampleRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        public Microsoft.Office.Tools.Ribbon.RibbonEditBox tableHeight;
        public Microsoft.Office.Tools.Ribbon.RibbonEditBox tableWidth;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CreateEmptyTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CreateTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddLine;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DelLine;
    }

    partial class ThisRibbonCollection
    {
        internal SampleRibbon SampleRibbon
        {
            get { return this.GetRibbon<SampleRibbon>(); }
        }
    }
}
