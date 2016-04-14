using System;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace ModuleLoader {

    public partial class MainRibbon {
        // EVENT HANDLERS
        private void MainRibbon_Load(object sender, RibbonUIEventArgs e) {
            Globals.ModuleLoader.Application.WorkbookActivate += Application_WorkbookActivate;
        }
        private void Application_WorkbookActivate(Workbook Wb) {
            // Enumerate all the VB files in the same directory as this workbook
            IEnumerable<FileInfo> modules = null;
            IEnumerable<FileInfo> classes = null;
            IEnumerable<FileInfo> forms = null;
            if (Wb.Path != "") {
                DirectoryInfo folder = new DirectoryInfo(Wb.Path);
                modules = folder.EnumerateFiles("*.bas");
                classes = folder.EnumerateFiles("*.cls");
                forms = folder.EnumerateFiles("*.frm");
            }

            // Add their names to the appropriate menus
            replaceDropDownItems(ModulesDrop, modules, "ModuleInsert");
            replaceDropDownItems(ClassesDrop, classes, "ComAddInsDialog");
            replaceDropDownItems(FormsDrop, forms, "FormPublish");
        }
        private void ModulesDrop_SelectionChanged(object sender, RibbonControlEventArgs e) {
            FileInfo moduleFile = ModulesDrop.SelectedItem.Tag as FileInfo;
            if (moduleFile != null)
                doImport(moduleFile);
            ModulesDrop.SelectedItemIndex = 0;
        }
        private void ClassesDrop_SelectionChanged(object sender, RibbonControlEventArgs e) {
            FileInfo classFile = ClassesDrop.SelectedItem.Tag as FileInfo;
            if (classFile != null)
                doImport(classFile);
            ClassesDrop.SelectedItemIndex = 0;
        }
        private void FormsDrop_SelectionChanged(object sender, RibbonControlEventArgs e) {
            FileInfo formFile = FormsDrop.SelectedItem.Tag as FileInfo;
            if (formFile != null)
                doImport(formFile);
            FormsDrop.SelectedItemIndex = 0;
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

        // HELPER FUNCTIONS
        private void replaceDropDownItems(RibbonDropDown dropDown, IEnumerable<FileInfo> files, string officeImgId) {
            // No matter what, add a placeholder item to the drop down
            IList<RibbonDropDownItem> items = dropDown.Items;
            items.Clear();
            RibbonDropDownItem item = Factory.CreateRibbonDropDownItem();
            item.Label = "Select a File...";
            items.Add(item);

            // If the given list contains some files, then add their names to the drop down
            if (files?.Count() > 0) {
                foreach (FileInfo m in files) {
                    item = Factory.CreateRibbonDropDownItem();
                    item.Label = Path.GetFileName(m.Name);
                    item.OfficeImageId = officeImgId;
                    item.Tag = m;
                    items.Add(item);
                }
            }
        }
        private void doImport(FileInfo module) {
            // Check if the selected module has already been imported to this workbook
            string name = Path.GetFileNameWithoutExtension(module.Name);
            VBComponents currModules = Globals.ModuleLoader.Application.ActiveWorkbook.VBProject.VBComponents;
            VBComponent vbc = null;
            try {
                vbc = currModules.Item(name);
            }
            catch (IndexOutOfRangeException e) { }

            // If so, ask the user if they want to re-import
            bool import = false;
            if (vbc == null)
                import = true;
            else {
                DialogResult result = MessageBox.Show("A module with this name already exists in this workbook.\nDo you want to overwrite it?",
                                                      "Module Loader", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                if (result == DialogResult.Yes) {
                    currModules.Remove(vbc);
                    import = true;
                }
            }

            // If the import is set to continue, then do it!
            if (import)
                currModules.Import(module.FullName);
        }
    }

}
