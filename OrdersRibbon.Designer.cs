
namespace OnlineOrders
{
    partial class OrdersRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public OrdersRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnOnlineOrders = this.Factory.CreateRibbonButton();
            this.ddbPartSelector = this.Factory.CreateRibbonDropDown();
            this.btnGenerate = this.Factory.CreateRibbonButton();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "CustomAddIn";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnOnlineOrders);
            this.group1.Items.Add(this.ddbPartSelector);
            this.group1.Items.Add(this.btnGenerate);
            this.group1.Label = "Store Orders";
            this.group1.Name = "group1";
            // 
            // btnOnlineOrders
            // 
            this.btnOnlineOrders.Label = "Load .csv files";
            this.btnOnlineOrders.Name = "btnOnlineOrders";
            this.btnOnlineOrders.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnOnlineOrders_Click);
            // 
            // ddbPartSelector
            // 
            ribbonDropDownItemImpl1.Label = "1";
            ribbonDropDownItemImpl2.Label = "2";
            ribbonDropDownItemImpl3.Label = "3";
            ribbonDropDownItemImpl4.Label = "4";
            ribbonDropDownItemImpl5.Label = "5";
            this.ddbPartSelector.Items.Add(ribbonDropDownItemImpl1);
            this.ddbPartSelector.Items.Add(ribbonDropDownItemImpl2);
            this.ddbPartSelector.Items.Add(ribbonDropDownItemImpl3);
            this.ddbPartSelector.Items.Add(ribbonDropDownItemImpl4);
            this.ddbPartSelector.Items.Add(ribbonDropDownItemImpl5);
            this.ddbPartSelector.Label = " Part";
            this.ddbPartSelector.Name = "ddbPartSelector";
            this.ddbPartSelector.SizeString = "1";
            // 
            // btnGenerate
            // 
            this.btnGenerate.Label = "Generate";
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnGenerate_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // OrdersRibbon
            // 
            this.Name = "OrdersRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.OrdersRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOnlineOrders;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGenerate;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddbPartSelector;
    }

    partial class ThisRibbonCollection
    {
        internal OrdersRibbon Ribbon1
        {
            get { return this.GetRibbon<OrdersRibbon>(); }
        }
    }
}
