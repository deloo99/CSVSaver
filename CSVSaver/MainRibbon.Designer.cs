namespace CSVSaver
{
    partial class MainRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MainRibbon()
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
            this.KTO = this.Factory.CreateRibbonGroup();
            this.ImportToCSV = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.KTO.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.KTO);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // KTO
            // 
            this.KTO.Items.Add(this.ImportToCSV);
            this.KTO.Label = "КТО";
            this.KTO.Name = "KTO";
            // 
            // ImportToCSV
            // 
            this.ImportToCSV.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ImportToCSV.Description = "Испортировать содержимео в CSV файл";
            this.ImportToCSV.Image = global::CSVSaver.Properties.Resources.icons8_импорт_из_csv_50;
            this.ImportToCSV.Label = "Импорт в CSV файл";
            this.ImportToCSV.Name = "ImportToCSV";
            this.ImportToCSV.ShowImage = true;
            this.ImportToCSV.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ImportToCSV_Click);
            // 
            // MainRibbon
            // 
            this.Name = "MainRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.KTO.ResumeLayout(false);
            this.KTO.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup KTO;
        private Microsoft.Office.Tools.Ribbon.RibbonButton ImportToCSV;
    }

    partial class ThisRibbonCollection
    {
        internal MainRibbon Ribbon1
        {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }
}
