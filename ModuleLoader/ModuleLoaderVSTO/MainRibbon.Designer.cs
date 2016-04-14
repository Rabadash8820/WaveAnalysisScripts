namespace ModuleLoader {
    partial class MainRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MainRibbon()
            : base(Globals.Factory.GetRibbonFactory()) {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.ModLoaderTab = this.Factory.CreateRibbonTab();
            this.MainGrp = this.Factory.CreateRibbonGroup();
            this.BrowseBtn = this.Factory.CreateRibbonButton();
            this.ModulesDrop = this.Factory.CreateRibbonDropDown();
            this.ClassesDrop = this.Factory.CreateRibbonDropDown();
            this.FormsDrop = this.Factory.CreateRibbonDropDown();
            this.ModLoaderTab.SuspendLayout();
            this.MainGrp.SuspendLayout();
            this.SuspendLayout();
            // 
            // ModLoaderTab
            // 
            this.ModLoaderTab.Groups.Add(this.MainGrp);
            this.ModLoaderTab.Label = "Module Loader";
            this.ModLoaderTab.Name = "ModLoaderTab";
            // 
            // MainGrp
            // 
            this.MainGrp.Items.Add(this.ModulesDrop);
            this.MainGrp.Items.Add(this.ClassesDrop);
            this.MainGrp.Items.Add(this.FormsDrop);
            this.MainGrp.Items.Add(this.BrowseBtn);
            this.MainGrp.Label = "Load";
            this.MainGrp.Name = "MainGrp";
            // 
            // BrowseBtn
            // 
            this.BrowseBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BrowseBtn.Label = "Browse";
            this.BrowseBtn.Name = "BrowseBtn";
            this.BrowseBtn.OfficeImageId = "BrowseBackgroundImage";
            this.BrowseBtn.ScreenTip = "From File";
            this.BrowseBtn.ShowImage = true;
            this.BrowseBtn.SuperTip = "Import modules from your computer or from other computers you\'re connected to.";
            this.BrowseBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BrowseBtn_Click);
            // 
            // ModulesDrop
            // 
            this.ModulesDrop.Label = "Modules";
            this.ModulesDrop.Name = "ModulesDrop";
            this.ModulesDrop.OfficeImageId = "ModuleInsert";
            this.ModulesDrop.ShowImage = true;
            this.ModulesDrop.ShowItemImage = false;
            this.ModulesDrop.SizeString = "The really long name of a module";
            this.ModulesDrop.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ModulesDrop_SelectionChanged);
            // 
            // ClassesDrop
            // 
            this.ClassesDrop.Label = "Classes";
            this.ClassesDrop.Name = "ClassesDrop";
            this.ClassesDrop.OfficeImageId = "ComAddInsDialog";
            this.ClassesDrop.ShowImage = true;
            this.ClassesDrop.ShowItemImage = false;
            this.ClassesDrop.SizeString = "The really long name of a module";
            this.ClassesDrop.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ClassesDrop_SelectionChanged);
            // 
            // FormsDrop
            // 
            this.FormsDrop.Label = "Forms";
            this.FormsDrop.Name = "FormsDrop";
            this.FormsDrop.OfficeImageId = "FormPublish";
            this.FormsDrop.ShowImage = true;
            this.FormsDrop.ShowItemImage = false;
            this.FormsDrop.SizeString = "The really long name of a module";
            this.FormsDrop.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.FormsDrop_SelectionChanged);
            // 
            // MainRibbon
            // 
            this.Name = "MainRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.ModLoaderTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MainRibbon_Load);
            this.ModLoaderTab.ResumeLayout(false);
            this.ModLoaderTab.PerformLayout();
            this.MainGrp.ResumeLayout(false);
            this.MainGrp.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab ModLoaderTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup MainGrp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BrowseBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ModulesDrop;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ClassesDrop;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown FormsDrop;
    }

    partial class ThisRibbonCollection {
        internal MainRibbon MainRibbon {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }
}
