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

        private void Application_WorkbookBeforeClose(Workbook Wb, ref bool Cancel) {
            if (_app.Workbooks.Count == 1)
                refreshImportItems();
        }

        private void Application_WorkbookActivate(Workbook Wb) {
            refreshImportItems(Wb.Path);
        }
        private void ImportAllBtn_Click(object sender, RibbonControlEventArgs e) {

        }
        private void ImportModulesDrop_SelectionChanged(object sender, RibbonControlEventArgs e) {
            RibbonDropDownItem item = ImportModulesDrop.SelectedItem;
            FileInfo moduleFile = item.Tag as FileInfo;
            if (moduleFile != null) {
                bool success = doImport(moduleFile);
                if (!success)
                    ImportModulesDrop.Items.Remove(item);
            }
            ImportModulesDrop.SelectedItemIndex = 0;
        }
        private void ImportClassesDrop_SelectionChanged(object sender, RibbonControlEventArgs e) {
            RibbonDropDownItem item = ImportClassesDrop.SelectedItem;
            FileInfo classFile = item.Tag as FileInfo;
            if (classFile != null) {
                bool success = doImport(classFile);
                if (!success)
                    ImportClassesDrop.Items.Remove(item);
            }
            ImportClassesDrop.SelectedItemIndex = 0;
        }
        private void ImportFormsDrop_SelectionChanged(object sender, RibbonControlEventArgs e) {
            RibbonDropDownItem item = ImportFormsDrop.SelectedItem;
            FileInfo formFile = item.Tag as FileInfo;
            if (formFile != null) {
                bool success = doImport(formFile);
                if (!success)
                    ImportFormsDrop.Items.Remove(item);
            }
            ImportFormsDrop.SelectedItemIndex = 0;
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
            if (result != DialogResult.Cancel) {
                FileInfo f = new FileInfo(dialog.FileName);
                doImport(f);
            }
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
            string path = Globals.ModuleLoader.Application.ActiveWorkbook.Path;
            if (path != "")
                refreshImportItems(path);
        }

        // HELPER FUNCTIONS
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
        private bool doImport(FileInfo module) {
            // Try to get the VB name of this module
            // Inform the user if the file could not be found or could not be loaded
            string name = "";
            try {
                name = getModuleName(module.FullName);
            }
            catch (FileNotFoundException) {
                MessageBox.Show(
                    $"{module.FullName} could not be found.  Is it possible that the file was moved or renamed?\n\nYou can click '{RefreshBtn.Label}' to refresh the list of available VB files.",
                    "ModuleLoader", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return false;
            }
            catch (FileLoadException) {
                MessageBox.Show(
                    $"{module.FullName} could not be loaded.  If you edited this file, make sure that there is a line similar to 'Attribute VB_Name \"SomeName\"' above the actual code.",
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

            // If so, ask the user if they want to replace it
            bool import = false;
            if (vbc == null)
                import = true;
            else {
                DialogResult result = MessageBox.Show(
                    $"A module with the name {name} already exists in this workbook.\nDo you want to replace it?",
                    "Module Loader", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                if (result == DialogResult.Yes) {
                    currModules.Remove(vbc);
                    import = true;
                }
                else if (result== DialogResult.No)
                    return true;
            }

            // If the import is set to continue, then do it!
            // Report any errors to the user
            bool success = false;
            if (import) {
                currModules.Import(module.FullName);
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
                    nameLineFound = (line.Contains("Attribute VB_Name"));
                    if (!nameLineFound)
                        continue;

                    string[] parts = line.Split('\"');
                    if (parts.Count() == 3) {
                        name = parts[1];
                        break;
                    }
                    else
                        throw new FileLoadException($"{path} had a line with 'Attribute VB_Name' but no name.  What gives?", path);
                }
            }

            // If everything went well, return the Module's name
            // If the "Attribute VB_Name" line was not found, then this file is invalid so throw an exception
            if (nameLineFound)
                return name;
            else
                throw new FileLoadException($"{path} did not have a line with 'Attribute VB_Name'.?", path);
        }

    }

}
