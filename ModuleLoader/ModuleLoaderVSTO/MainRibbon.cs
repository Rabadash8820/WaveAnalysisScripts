using System;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace ModuleLoader {

    public partial class MainRibbon {
        // HIDDEN FIELDS
        private Excel.Application _app;

        // EVENT HANDLERS
        private void MainRibbon_Load(object sender, RibbonUIEventArgs e) {
            _app = Globals.ModuleLoader.Application;
            _app.WorkbookActivate += Application_WorkbookActivate;
            _app.WorkbookBeforeClose += Application_WorkbookBeforeClose;
        }

        private void Application_WorkbookActivate(Workbook Wb) {
            refreshImportItems(Wb.Path);
            setVisible(true);
        }
        private void Application_WorkbookBeforeClose(Workbook Wb, ref bool Cancel) {
            if (_app.Workbooks.Count == 1)
                setVisible(false);
        }

        private void ImportAllBtn_Click(object sender, RibbonControlEventArgs e) {
            // Refresh the drop down lists of importable VB files
            refreshImportItems(_app.ActiveWorkbook.Path);

            // Try to import all of those files
            IEnumerable<RibbonDropDownItem> allItems = ImportModulesDrop.Items.Union(
                                                       ImportClassesDrop.Items).Union(
                                                       ImportFormsDrop.Items);
            foreach (RibbonDropDownItem item in allItems)
                importItem(item);
        }
        private void ImportModulesDrop_SelectionChanged(object sender, RibbonControlEventArgs e) {
            importItem(ImportModulesDrop.SelectedItem);
        }
        private void ImportClassesDrop_SelectionChanged(object sender, RibbonControlEventArgs e) {
            importItem(ImportClassesDrop.SelectedItem);
        }
        private void ImportFormsDrop_SelectionChanged(object sender, RibbonControlEventArgs e) {
            importItem(ImportFormsDrop.SelectedItem);
        }
        private void BrowseBtn_Click(object sender, RibbonControlEventArgs e) {
            // Show an Open File dialog
            OpenFileDialog dialog = new OpenFileDialog() {
                Multiselect = false,
                Filter = string.Join("|",
                    "VB Files (*.frm, *.bas, *.cls)|*.frm;*.bas;*.cls",
                    "Form Files (*.frm)|*.frm",
                    "Basic Files (*.bas)|*.bas",
                    "Class Files (*.cls)|*.cls"),
                Title = "Import File"
            };
            DialogResult result = dialog.ShowDialog();

            // If the user actually selected a file, then import it
            if (result != DialogResult.Cancel)
                importModule(dialog.FileName);
        }

        private void ExportAllBtn_Click(object sender, RibbonControlEventArgs e) {

        }
        private void ExportModulesDrop_SelectionChanged(object sender, RibbonControlEventArgs e) {

        }
        private void ExportClassesDrop_SelectionChanged(object sender, RibbonControlEventArgs e) {

        }
        private void ExportFormsDrop_SelectionChanged(object sender, RibbonControlEventArgs e) {

        }

        private void RefreshBtn_Click(object sender, RibbonControlEventArgs e) {
            refreshImportItems(_app.ActiveWorkbook.Path);
        }
        private void AlwaysReplaceRadio_Click(object sender, RibbonControlEventArgs e) {
            toggleReplaceBtns(AlwaysReplaceRadio, NeverReplaceRadio);
        }
        private void NeverReplaceRadio_Click(object sender, RibbonControlEventArgs e) {
            toggleReplaceBtns(NeverReplaceRadio, AlwaysReplaceRadio);
        }

        // HELPER FUNCTIONS
        private void setVisible(bool visibility) {
            ModLoaderTab.Visible = visibility;
        }
        private void refreshImportItems(string path = "") {
            // Enumerate all the VB files in the same directory as this workbook and add them to the import drop downs
            DirectoryInfo folder = (path == "" ? null : new DirectoryInfo(path));
            refreshDropDownItems(ImportModulesDrop, folder, "*.bas", "ModuleInsert");
            refreshDropDownItems(ImportClassesDrop, folder, "*.cls", "ComAddInsDialog");
            refreshDropDownItems(ImportFormsDrop, folder, "*.frm", "FormPublish");
        }
        private void refreshDropDownItems(RibbonDropDown dropDown, DirectoryInfo folder, string searchPattern, string officeImgId) {
            // No matter what, clear the drop down's items and add a placeholder item
            IList<RibbonDropDownItem> items = dropDown.Items;
            items.Clear();
            RibbonDropDownItem item = Factory.CreateRibbonDropDownItem();
            item.Label = "Select a File...";
            items.Add(item);

            // If no folder was provided, or it contains no files matching the given search pattern, then just return
            if (folder == null)
                return;
            IEnumerable<FileInfo> files = folder.EnumerateFiles(searchPattern);
            if (files.Count() == 0)
                return;
            
            // Otherwise, add the VB file names to the given drop down
            foreach (FileInfo m in files) {
                item = Factory.CreateRibbonDropDownItem();
                item.Label = Path.GetFileName(m.Name);
                item.OfficeImageId = officeImgId;
                item.Tag = m;
                items.Add(item);
            }
        }
        private void importItem(RibbonDropDownItem item) {
            RibbonDropDown dropDown = item.Parent as RibbonDropDown;
            FileInfo moduleFile = item.Tag as FileInfo;
            if (moduleFile != null) {
                bool success = importModule(moduleFile.FullName);
                if (!success)
                    dropDown.Items.Remove(item);
            }
            dropDown.SelectedItemIndex = 0;
        }
        private bool importModule(string path) {
            // Try to get the VB name of this module
            // Inform the user if the file could not be found or could not be loaded
            string name = "";
            try {
                name = getModuleName(path);
            }
            catch (FileNotFoundException) {
                MessageBox.Show(
                    $"{path} could not be found.  Is it possible that the file was moved or renamed?\n\nYou can click '{RefreshBtn.Label}' to refresh the list of available VB files.",
                    "ModuleLoader", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return false;
            }
            catch (FileLoadException) {
                MessageBox.Show(
                    $"{path} could not be loaded.  If you edited this file, make sure that there is a line similar to 'Attribute VB_Name = \"SomeName\"' above the actual code.",
                    "ModuleLoader", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return true;
            }

            // Check if a module with this name already exists in this workbook
            VBComponents currModules = _app.ActiveWorkbook.VBProject.VBComponents;
            VBComponent vbc = null;
            try {
                vbc = currModules.Item(name);
            }
            catch (IndexOutOfRangeException) { }

            // If so, ask the user if they want to replace it, unless they have set one of the always/never overwrite options
            bool import = false;
            if (vbc == null)
                import = true;
            else {
                if (AlwaysReplaceRadio.Checked) {
                    currModules.Remove(vbc);
                    import = true;
                }
                else if (NeverReplaceRadio.Checked)
                    return true;
                else {
                    DialogResult result = MessageBox.Show(
                        $"A module with the name {name} already exists in this workbook.\nDo you want to replace it?",
                        "Module Loader", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                    if (result == DialogResult.Yes) {
                        currModules.Remove(vbc);
                        import = true;
                    }
                    else
                        return true;
                }
            }

            // If the import is set to continue, then do it!
            // Report any errors to the user
            bool success = false;
            if (import) {
                currModules.Import(path);
                success = true;
            }
            return success;
        }
        private string getModuleName(string path) {
            string name = "";

            // Read the file, looking for the "Attribute VB_Name" line
            string line;
            bool nameLineFound = false;
            using (StreamReader file = new StreamReader(path)) {
                while ((line = file.ReadLine()) != null) {
                    nameLineFound = (line.Contains("Attribute VB_Name = "));
                    if (!nameLineFound)
                        continue;

                    string[] parts = line.Split('\"');
                    if (parts.Count() == 3) {
                        name = parts[1];
                        break;
                    }
                    else
                        throw new FileLoadException($"{path} had a line with 'Attribute VB_Name  = ' but no name.  What gives?", path);
                }
            }

            // If everything went well, return the Module's name
            // If the "Attribute VB_Name" line was not found, then this file is invalid so throw an exception
            if (nameLineFound)
                return name;
            else
                throw new FileLoadException($"{path} did not have a line with 'Attribute VB_Name = '.?", path);
        }
        private void toggleReplaceBtns(RibbonToggleButton trigger, RibbonToggleButton listener) {
            if (trigger.Checked && listener.Checked)
                listener.Checked = false;
        }
    }

}
