using System;
using System.Runtime.InteropServices;
using Inventor;
using System.Windows.Forms;

namespace InventorFileManager
{
    [GuidAttribute("12345678-1234-1234-1234-123456789012")]
    public class StandardAddInServer : Inventor.ApplicationAddInServer
    {
        private Inventor.Application m_inventorApplication;
        private ButtonDefinition m_exportButton;
        private ButtonDefinition m_renameButton;
        private CommandManager m_commandManager;

        public void Activate(Inventor.ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            try
            {
                m_inventorApplication = addInSiteObject.Application;
                m_commandManager = m_inventorApplication.CommandManager;

                // Create the export button
                m_exportButton = m_commandManager.ControlDefinitions.AddButtonDefinition(
                    "File Name Export",
                    "FileExport",
                    CommandTypesEnum.kShapeEditCmdType,
                    "{12345678-1234-1234-1234-123456789012}",
                    "Export file names to Excel",
                    "Export .iam, .ipt, .idw file names from selected folder to Excel",
                    null,
                    null,
                    ButtonDisplayEnum.kDisplayTextInLearningMode);

                // Create the rename button
                m_renameButton = m_commandManager.ControlDefinitions.AddButtonDefinition(
                    "File Rename",
                    "FileRename",
                    CommandTypesEnum.kShapeEditCmdType,
                    "{12345678-1234-1234-1234-123456789013}",
                    "Rename files from Excel",
                    "Rename files based on Excel mapping",
                    null,
                    null,
                    ButtonDisplayEnum.kDisplayTextInLearningMode);

                // Add event handlers
                m_exportButton.OnExecute += ExportButton_OnExecute;
                m_renameButton.OnExecute += RenameButton_OnExecute;

                // Add buttons to ribbon
                AddButtonsToRibbon();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error during activation: " + ex.Message);
            }
        }

        public void Deactivate()
        {
            try
            {
                if (m_exportButton != null)
                {
                    m_exportButton.OnExecute -= ExportButton_OnExecute;
                }
                if (m_renameButton != null)
                {
                    m_renameButton.OnExecute -= RenameButton_OnExecute;
                }

                m_inventorApplication = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error during deactivation: " + ex.Message);
            }
        }

        public void ExecuteCommand(int commandID)
        {
            // Not used in this implementation
        }

        public object Automation
        {
            get { return null; }
        }

        private void AddButtonsToRibbon()
        {
            try
            {
                // Get the ribbon
                Ribbon ribbon = m_inventorApplication.UserInterfaceManager.Ribbons["Part"];
                RibbonTab addInsTab = ribbon.RibbonTabs["id_TabAddIns"];
                
                // Create a new panel for our buttons
                RibbonPanel fileManagerPanel = addInsTab.RibbonPanels.Add(
                    "File Manager",
                    "FileManagerPanel",
                    "{12345678-1234-1234-1234-123456789014}");

                // Add buttons to the panel
                fileManagerPanel.CommandControls.AddButton(m_exportButton);
                fileManagerPanel.CommandControls.AddButton(m_renameButton);

                // Also add to Assembly and Drawing ribbons
                AddToRibbon("Assembly");
                AddToRibbon("Drawing");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error adding buttons to ribbon: " + ex.Message);
            }
        }

        private void AddToRibbon(string ribbonName)
        {
            try
            {
                Ribbon ribbon = m_inventorApplication.UserInterfaceManager.Ribbons[ribbonName];
                RibbonTab addInsTab = ribbon.RibbonTabs["id_TabAddIns"];
                
                RibbonPanel fileManagerPanel = addInsTab.RibbonPanels.Add(
                    "File Manager",
                    "FileManagerPanel" + ribbonName,
                    "{12345678-1234-1234-1234-12345678901" + ribbonName.Length + "}");

                fileManagerPanel.CommandControls.AddButton(m_exportButton);
                fileManagerPanel.CommandControls.AddButton(m_renameButton);
            }
            catch
            {
                // Ribbon might not exist, ignore
            }
        }

        private void ExportButton_OnExecute(NameValueMap context)
        {
            try
            {
                FileExportForm exportForm = new FileExportForm(m_inventorApplication);
                exportForm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error opening export form: " + ex.Message);
            }
        }

        private void RenameButton_OnExecute(NameValueMap context)
        {
            try
            {
                FileRenameForm renameForm = new FileRenameForm(m_inventorApplication);
                renameForm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error opening rename form: " + ex.Message);
            }
        }
    }
}
