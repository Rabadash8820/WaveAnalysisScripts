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
            this.ImportGrp = this.Factory.CreateRibbonGroup();
            this.ImportModulesDrop = this.Factory.CreateRibbonDropDown();
            this.ImportClassesDrop = this.Factory.CreateRibbonDropDown();
            this.ImportFormsDrop = this.Factory.CreateRibbonDropDown();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.ExportGrp = this.Factory.CreateRibbonGroup();
            this.ExportModulesDrop = this.Factory.CreateRibbonDropDown();
            this.ExportClassesDrop = this.Factory.CreateRibbonDropDown();
            this.ExportFormsDrop = this.Factory.CreateRibbonDropDown();
            this.RefreshGrp = this.Factory.CreateRibbonGroup();
            this.OverwriteGrp = this.Factory.CreateRibbonGroup();
            this.RefreshBtn = this.Factory.CreateRibbonButton();
            this.ImportAllBtn = this.Factory.CreateRibbonButton();
            this.BrowseBtn = this.Factory.CreateRibbonButton();
            this.ExportAllBtn = this.Factory.CreateRibbonButton();
            this.AlwaysReplaceRadio = this.Factory.CreateRibbonToggleButton();
            this.NeverReplaceRadio = this.Factory.CreateRibbonToggleButton();
            this.ModLoaderTab.SuspendLayout();
            this.ImportGrp.SuspendLayout();
            this.ExportGrp.SuspendLayout();
            this.RefreshGrp.SuspendLayout();
            this.OverwriteGrp.SuspendLayout();
            this.SuspendLayout();
            // 
            // ModLoaderTab
            // 
            this.ModLoaderTab.Groups.Add(this.RefreshGrp);
            this.ModLoaderTab.Groups.Add(this.ImportGrp);
            this.ModLoaderTab.Groups.Add(this.ExportGrp);
            this.ModLoaderTab.Groups.Add(this.OverwriteGrp);
            this.ModLoaderTab.Label = "Module Loader";
            this.ModLoaderTab.Name = "ModLoaderTab";
            // 
            // ImportGrp
            // 
            this.ImportGrp.Items.Add(this.ImportAllBtn);
            this.ImportGrp.Items.Add(this.ImportModulesDrop);
            this.ImportGrp.Items.Add(this.ImportClassesDrop);
            this.ImportGrp.Items.Add(this.ImportFormsDrop);
            this.ImportGrp.Items.Add(this.separator1);
            this.ImportGrp.Items.Add(this.BrowseBtn);
            this.ImportGrp.Label = "Import";
            this.ImportGrp.Name = "ImportGrp";
            // 
            // ImportModulesDrop
            // 
            this.ImportModulesDrop.Label = "Modules";
            this.ImportModulesDrop.Name = "ImportModulesDrop";
            this.ImportModulesDrop.OfficeImageId = "ModuleInsert";
            this.ImportModulesDrop.ScreenTip = "Import Module";
            this.ImportModulesDrop.ShowImage = true;
            this.ImportModulesDrop.ShowItemImage = false;
            this.ImportModulesDrop.SizeString = "The really long name of a module";
            this.ImportModulesDrop.SuperTip = "Choose one of the Basic Modules from the same directory as this workbook to impor" +
    "t.";
            this.ImportModulesDrop.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ImportModulesDrop_SelectionChanged);
            // 
            // ImportClassesDrop
            // 
            this.ImportClassesDrop.Label = "Classes";
            this.ImportClassesDrop.Name = "ImportClassesDrop";
            this.ImportClassesDrop.OfficeImageId = "ComAddInsDialog";
            this.ImportClassesDrop.ScreenTip = "Import Class";
            this.ImportClassesDrop.ShowImage = true;
            this.ImportClassesDrop.ShowItemImage = false;
            this.ImportClassesDrop.SizeString = "The really long name of a module";
            this.ImportClassesDrop.SuperTip = "Choose one of the Class Modules from the same directory as this workbook to impor" +
    "t.";
            this.ImportClassesDrop.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ImportClassesDrop_SelectionChanged);
            // 
            // ImportFormsDrop
            // 
            this.ImportFormsDrop.Label = "Forms";
            this.ImportFormsDrop.Name = "ImportFormsDrop";
            this.ImportFormsDrop.OfficeImageId = "FormPublish";
            this.ImportFormsDrop.ScreenTip = "Import Form";
            this.ImportFormsDrop.ShowImage = true;
            this.ImportFormsDrop.ShowItemImage = false;
            this.ImportFormsDrop.SizeString = "The really long name of a module";
            this.ImportFormsDrop.SuperTip = "Choose one of the Form Modules from the same directory as this workbook to import" +
    ".";
            this.ImportFormsDrop.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ImportFormsDrop_SelectionChanged);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // ExportGrp
            // 
            this.ExportGrp.Items.Add(this.ExportAllBtn);
            this.ExportGrp.Items.Add(this.ExportModulesDrop);
            this.ExportGrp.Items.Add(this.ExportClassesDrop);
            this.ExportGrp.Items.Add(this.ExportFormsDrop);
            this.ExportGrp.Label = "Export";
            this.ExportGrp.Name = "ExportGrp";
            // 
            // ExportModulesDrop
            // 
            this.ExportModulesDrop.Label = "Modules";
            this.ExportModulesDrop.Name = "ExportModulesDrop";
            this.ExportModulesDrop.OfficeImageId = "ModuleInsert";
            this.ExportModulesDrop.ScreenTip = "Export Module";
            this.ExportModulesDrop.ShowImage = true;
            this.ExportModulesDrop.ShowItemImage = false;
            this.ExportModulesDrop.SizeString = "The really long name of a module";
            this.ExportModulesDrop.SuperTip = "Choose one of the Basic Modules from the same directory as this workbook to impor" +
    "t.";
            this.ExportModulesDrop.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExportModulesDrop_SelectionChanged);
            // 
            // ExportClassesDrop
            // 
            this.ExportClassesDrop.Label = "Classes";
            this.ExportClassesDrop.Name = "ExportClassesDrop";
            this.ExportClassesDrop.OfficeImageId = "ComAddInsDialog";
            this.ExportClassesDrop.ScreenTip = "Export Class";
            this.ExportClassesDrop.ShowImage = true;
            this.ExportClassesDrop.ShowItemImage = false;
            this.ExportClassesDrop.SizeString = "The really long name of a module";
            this.ExportClassesDrop.SuperTip = "Choose one of the Class Modules from the same directory as this workbook to impor" +
    "t.";
            this.ExportClassesDrop.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExportClassesDrop_SelectionChanged);
            // 
            // ExportFormsDrop
            // 
            this.ExportFormsDrop.Label = "Forms";
            this.ExportFormsDrop.Name = "ExportFormsDrop";
            this.ExportFormsDrop.OfficeImageId = "FormPublish";
            this.ExportFormsDrop.ScreenTip = "Export Form";
            this.ExportFormsDrop.ShowImage = true;
            this.ExportFormsDrop.ShowItemImage = false;
            this.ExportFormsDrop.SizeString = "The really long name of a module";
            this.ExportFormsDrop.SuperTip = "Choose one of the Form Modules from the same directory as this workbook to import" +
    ".";
            this.ExportFormsDrop.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExportFormsDrop_SelectionChanged);
            // 
            // RefreshGrp
            // 
            this.RefreshGrp.Items.Add(this.RefreshBtn);
            this.RefreshGrp.Name = "RefreshGrp";
            // 
            // OverwriteGrp
            // 
            this.OverwriteGrp.Items.Add(this.AlwaysReplaceRadio);
            this.OverwriteGrp.Items.Add(this.NeverReplaceRadio);
            this.OverwriteGrp.Label = "Overwrite";
            this.OverwriteGrp.Name = "OverwriteGrp";
            // 
            // RefreshBtn
            // 
            this.RefreshBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.RefreshBtn.Label = "Refresh All";
            this.RefreshBtn.Name = "RefreshBtn";
            this.RefreshBtn.OfficeImageId = "RefreshAll";
            this.RefreshBtn.ScreenTip = "Refresh All Drop Downs";
            this.RefreshBtn.ShowImage = true;
            this.RefreshBtn.SuperTip = "Refresh the import drop downs with VB files in this workbook\'s directory.  Refres" +
    "h the export drop downs with this workbook\'s VB modules.";
            this.RefreshBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RefreshBtn_Click);
            // 
            // ImportAllBtn
            // 
            this.ImportAllBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ImportAllBtn.Label = "Import All";
            this.ImportAllBtn.Name = "ImportAllBtn";
            this.ImportAllBtn.OfficeImageId = "Import";
            this.ImportAllBtn.ScreenTip = "Import All";
            this.ImportAllBtn.ShowImage = true;
            this.ImportAllBtn.SuperTip = "Import all VB files from the same directory as this workbook.";
            this.ImportAllBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ImportAllBtn_Click);
            // 
            // BrowseBtn
            // 
            this.BrowseBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BrowseBtn.Label = "From Other Source";
            this.BrowseBtn.Name = "BrowseBtn";
            this.BrowseBtn.OfficeImageId = "BrowseBackgroundImage";
            this.BrowseBtn.ScreenTip = "Import VB File From Other Source";
            this.BrowseBtn.ShowImage = true;
            this.BrowseBtn.SuperTip = "Import modules from your computer or from other computers you\'re connected to.";
            this.BrowseBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BrowseBtn_Click);
            // 
            // ExportAllBtn
            // 
            this.ExportAllBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ExportAllBtn.Label = "Export All";
            this.ExportAllBtn.Name = "ExportAllBtn";
            this.ExportAllBtn.OfficeImageId = "Export";
            this.ExportAllBtn.ScreenTip = "Export All";
            this.ExportAllBtn.ShowImage = true;
            this.ExportAllBtn.SuperTip = "Export all VB files in this workbook to its containing directory";
            this.ExportAllBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExportAllBtn_Click);
            // 
            // AlwaysReplaceRadio
            // 
            this.AlwaysReplaceRadio.Checked = true;
            this.AlwaysReplaceRadio.Label = "Always";
            this.AlwaysReplaceRadio.Name = "AlwaysReplaceRadio";
            this.AlwaysReplaceRadio.OfficeImageId = "ChangeToAcceptInvitation";
            this.AlwaysReplaceRadio.ShowImage = true;
            this.AlwaysReplaceRadio.SuperTip = "Importing modules will always overwrite modules with the same name.  Exporting mo" +
    "dules will always overwrite VB files with the same name.";
            this.AlwaysReplaceRadio.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AlwaysReplaceRadio_Click);
            // 
            // NeverReplaceRadio
            // 
            this.NeverReplaceRadio.Label = "Never";
            this.NeverReplaceRadio.Name = "NeverReplaceRadio";
            this.NeverReplaceRadio.OfficeImageId = "ChangeToDeclineInvitation";
            this.NeverReplaceRadio.ScreenTip = "Never Overwrite Files";
            this.NeverReplaceRadio.ShowImage = true;
            this.NeverReplaceRadio.SuperTip = "Importing modules will never overwrite existing modules with the same name.  Expo" +
    "rting modules will never overwrite VB files with the same name.";
            this.NeverReplaceRadio.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.NeverReplaceRadio_Click);
            // 
            // MainRibbon
            // 
            this.Name = "MainRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.ModLoaderTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MainRibbon_Load);
            this.ModLoaderTab.ResumeLayout(false);
            this.ModLoaderTab.PerformLayout();
            this.ImportGrp.ResumeLayout(false);
            this.ImportGrp.PerformLayout();
            this.ExportGrp.ResumeLayout(false);
            this.ExportGrp.PerformLayout();
            this.RefreshGrp.ResumeLayout(false);
            this.RefreshGrp.PerformLayout();
            this.OverwriteGrp.ResumeLayout(false);
            this.OverwriteGrp.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab ModLoaderTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ImportGrp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BrowseBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ImportModulesDrop;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ImportClassesDrop;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ImportFormsDrop;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ExportGrp;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ExportModulesDrop;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ExportClassesDrop;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ExportFormsDrop;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ImportAllBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ExportAllBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup RefreshGrp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RefreshBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton AlwaysReplaceRadio;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton NeverReplaceRadio;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup OverwriteGrp;
    }

    partial class ThisRibbonCollection {
        internal MainRibbon MainRibbon {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }
}
