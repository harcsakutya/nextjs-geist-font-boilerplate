using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;
using Inventor;

namespace InventorFileManager
{
    public partial class FileRenameForm : Form
    {
        private Inventor.Application m_inventorApp;
        private TextBox txtSourceFolder;
        private TextBox txtRenameFile;
        private Button btnBrowseSource;
        private Button btnBrowseRename;
        private Button btnRename;
        private Button btnCancel;
        private Button btnPreview;
        private ProgressBar progressBar;
        private Label lblStatus;
        private DataGridView dgvPreview;
        private CheckBox chkBackup;

        public FileRenameForm(Inventor.Application inventorApp)
        {
            m_inventorApp = inventorApp;
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Rename Files from Excel";
            this.Size = new System.Drawing.Size(800, 600);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimumSize = new System.Drawing.Size(800, 600);

            // Source folder selection
            Label lblSource = new Label();
            lblSource.Text = "Source Folder:";
            lblSource.Location = new System.Drawing.Point(20, 20);
            lblSource.Size = new System.Drawing.Size(100, 23);
            this.Controls.Add(lblSource);

            txtSourceFolder = new TextBox();
            txtSourceFolder.Location = new System.Drawing.Point(130, 20);
            txtSourceFolder.Size = new System.Drawing.Size(450, 23);
            this.Controls.Add(txtSourceFolder);

            btnBrowseSource = new Button();
            btnBrowseSource.Text = "Browse Folder";
            btnBrowseSource.Location = new System.Drawing.Point(590, 20);
            btnBrowseSource.Size = new System.Drawing.Size(100, 23);
            btnBrowseSource.Click += BtnBrowseSource_Click;
            this.Controls.Add(btnBrowseSource);

            // Rename Excel file selection
            Label lblRename = new Label();
            lblRename.Text = "Rename Excel File:";
            lblRename.Location = new System.Drawing.Point(20, 60);
            lblRename.Size = new System.Drawing.Size(100, 23);
            this.Controls.Add(lblRename);

            txtRenameFile = new TextBox();
            txtRenameFile.Location = new System.Drawing.Point(130, 60);
            txtRenameFile.Size = new System.Drawing.Size(450, 23);
            this.Controls.Add(txtRenameFile);

            btnBrowseRename = new Button();
            btnBrowseRename.Text = "Select XLSX";
            btnBrowseRename.Location = new System.Drawing.Point(590, 60);
            btnBrowseRename.Size = new System.Drawing.Size(100, 23);
            btnBrowseRename.Click += BtnBrowseRename_Click;
            this.Controls.Add(btnBrowseRename);

            // Backup checkbox
            chkBackup = new CheckBox();
            chkBackup.Text = "Create backup before renaming";
            chkBackup.Location = new System.Drawing.Point(130, 90);
            chkBackup.Size = new System.Drawing.Size(200, 23);
            chkBackup.Checked = true;
            this.Controls.Add(chkBackup);

            // Preview button
            btnPreview = new Button();
            btnPreview.Text = "Preview Changes";
            btnPreview.Location = new System.Drawing.Point(590, 90);
            btnPreview.Size = new System.Drawing.Size(100, 30);
            btnPreview.Click += BtnPreview_Click;
            this.Controls.Add(btnPreview);

            // Preview grid
            dgvPreview = new DataGridView();
            dgvPreview.Location = new System.Drawing.Point(20, 130);
            dgvPreview.Size = new System.Drawing.Size(750, 350);
            dgvPreview.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            dgvPreview.AllowUserToAddRows = false;
            dgvPreview.AllowUserToDeleteRows = false;
            dgvPreview.ReadOnly = true;
            dgvPreview.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvPreview.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            this.Controls.Add(dgvPreview);

            // Progress bar
            progressBar = new ProgressBar();
            progressBar.Location = new System.Drawing.Point(20, 490);
            progressBar.Size = new System.Drawing.Size(750, 23);
            progressBar.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            progressBar.Visible = false;
            this.Controls.Add(progressBar);

            // Status label
            lblStatus = new Label();
            lblStatus.Location = new System.Drawing.Point(20, 520);
            lblStatus.Size = new System.Drawing.Size(750, 23);
            lblStatus.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            lblStatus.Text = "Ready to rename files";
            this.Controls.Add(lblStatus);

            // Buttons
            btnRename = new Button();
            btnRename.Text = "Rename Files";
            btnRename.Location = new System.Drawing.Point(580, 550);
            btnRename.Size = new System.Drawing.Size(100, 30);
            btnRename.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnRename.Click += BtnRename_Click;
            btnRename.Enabled = false;
            this.Controls.Add(btnRename);

            btnCancel = new Button();
            btnCancel.Text = "Cancel";
            btnCancel.Location = new System.Drawing.Point(690, 550);
            btnCancel.Size = new System.Drawing.Size(80, 30);
            btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnCancel.DialogResult = DialogResult.Cancel;
            this.Controls.Add(btnCancel);

            this.CancelButton = btnCancel;
        }

        private void BtnBrowseSource_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select folder containing Inventor files to rename";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtSourceFolder.Text = dialog.SelectedPath;
                }
            }
        }

        private void BtnBrowseRename_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                dialog.Title = "Select Rename Excel File";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtRenameFile.Text = dialog.FileName;
                }
            }
        }

        private void BtnPreview_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtSourceFolder.Text))
            {
                MessageBox.Show("Please select a source folder.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (string.IsNullOrEmpty(txtRenameFile.Text))
            {
                MessageBox.Show("Please select a rename Excel file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!Directory.Exists(txtSourceFolder.Text))
            {
                MessageBox.Show("Source folder does not exist.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!File.Exists(txtRenameFile.Text))
            {
                MessageBox.Show("Rename Excel file does not exist.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                LoadPreview();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading preview: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadPreview()
        {
            lblStatus.Text = "Loading preview...";
            dgvPreview.DataSource = null;
            dgvPreview.Columns.Clear();

            // Load rename mappings from Excel
            Dictionary<string, string> renameMappings = LoadRenameMappings();
            
            if (renameMappings.Count == 0)
            {
                MessageBox.Show("No valid rename mappings found in Excel file.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Get all Inventor files in the source folder
            List<FileInfo> inventorFiles = GetInventorFiles(txtSourceFolder.Text);

            // Create preview data
            var previewData = new List<RenamePreviewItem>();

            foreach (var mapping in renameMappings)
            {
                string oldFileName = mapping.Key;
                string newFileName = mapping.Value;

                // Find matching files
                var matchingFiles = inventorFiles.Where(f => 
                    Path.GetFileNameWithoutExtension(f.Name).Equals(oldFileName, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                foreach (var file in matchingFiles)
                {
                    string newFullPath = Path.Combine(file.DirectoryName, newFileName + file.Extension);
                    bool canRename = !File.Exists(newFullPath) || newFullPath.Equals(file.FullName, StringComparison.OrdinalIgnoreCase);

                    previewData.Add(new RenamePreviewItem
                    {
                        OldFileName = file.Name,
                        NewFileName = newFileName + file.Extension,
                        FullPath = file.FullName,
                        Status = canRename ? "Ready" : "Conflict - Target exists",
                        CanRename = canRename
                    });
                }
            }

            // Setup DataGridView
            dgvPreview.Columns.Add("OldFileName", "Current Name");
            dgvPreview.Columns.Add("NewFileName", "New Name");
            dgvPreview.Columns.Add("FullPath", "Full Path");
            dgvPreview.Columns.Add("Status", "Status");

            dgvPreview.Columns["FullPath"].Width = 300;

            foreach (var item in previewData)
            {
                int rowIndex = dgvPreview.Rows.Add(item.OldFileName, item.NewFileName, item.FullPath, item.Status);
                
                if (!item.CanRename)
                {
                    dgvPreview.Rows[rowIndex].DefaultCellStyle.BackColor = System.Drawing.Color.LightCoral;
                }
                else
                {
                    dgvPreview.Rows[rowIndex].DefaultCellStyle.BackColor = System.Drawing.Color.LightGreen;
                }
            }

            btnRename.Enabled = previewData.Any(p => p.CanRename);
            lblStatus.Text = $"Preview loaded: {previewData.Count} files to process, {previewData.Count(p => p.CanRename)} can be renamed";
        }

        private Dictionary<string, string> LoadRenameMappings()
        {
            Dictionary<string, string> mappings = new Dictionary<string, string>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(txtRenameFile.Text)))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    throw new Exception("No worksheet found in Excel file");
                }

                int rowCount = worksheet.Dimension?.Rows ?? 0;
                
                // Look for "Old File" and "New File" columns
                int oldFileCol = -1, newFileCol = -1;
                
                for (int col = 1; col <= worksheet.Dimension?.Columns; col++)
                {
                    string header = worksheet.Cells[1, col].Value?.ToString()?.Trim();
                    if (string.IsNullOrEmpty(header)) continue;

                    if (header.Equals("Old File", StringComparison.OrdinalIgnoreCase) || 
                        header.Equals("OldFile", StringComparison.OrdinalIgnoreCase))
                    {
                        oldFileCol = col;
                    }
                    else if (header.Equals("New File", StringComparison.OrdinalIgnoreCase) || 
                             header.Equals("NewFile", StringComparison.OrdinalIgnoreCase))
                    {
                        newFileCol = col;
                    }
                }

                if (oldFileCol == -1 || newFileCol == -1)
                {
                    throw new Exception("Excel file must contain 'Old File' and 'New File' columns");
                }

                for (int row = 2; row <= rowCount; row++)
                {
                    string oldFile = worksheet.Cells[row, oldFileCol].Value?.ToString()?.Trim();
                    string newFile = worksheet.Cells[row, newFileCol].Value?.ToString()?.Trim();

                    if (!string.IsNullOrEmpty(oldFile) && !string.IsNullOrEmpty(newFile))
                    {
                        // Remove file extension if present
                        oldFile = Path.GetFileNameWithoutExtension(oldFile);
                        newFile = Path.GetFileNameWithoutExtension(newFile);
                        
                        if (!mappings.ContainsKey(oldFile))
                        {
                            mappings[oldFile] = newFile;
                        }
                    }
                }
            }

            return mappings;
        }

        private List<FileInfo> GetInventorFiles(string rootPath)
        {
            List<FileInfo> files = new List<FileInfo>();
            string[] extensions = { "*.iam", "*.ipt", "*.idw" };

            foreach (string extension in extensions)
            {
                try
                {
                    var foundFiles = Directory.GetFiles(rootPath, extension, SearchOption.AllDirectories)
                                           .Select(f => new FileInfo(f))
                                           .ToList();
                    files.AddRange(foundFiles);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error searching for {extension}: {ex.Message}");
                }
            }

            return files.OrderBy(f => f.FullName).ToList();
        }

        private async void BtnRename_Click(object sender, EventArgs e)
        {
            if (dgvPreview.Rows.Count == 0)
            {
                MessageBox.Show("Please preview changes first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var result = MessageBox.Show(
                "Are you sure you want to rename the files? This action cannot be undone automatically.",
                "Confirm Rename",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result != DialogResult.Yes)
                return;

            try
            {
                btnRename.Enabled = false;
                btnPreview.Enabled = false;
                progressBar.Visible = true;
                progressBar.Maximum = dgvPreview.Rows.Count;
                progressBar.Value = 0;

                await System.Threading.Tasks.Task.Run(() => PerformRename());

                lblStatus.Text = "Rename operation completed successfully!";
                MessageBox.Show("Files renamed successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.DialogResult = DialogResult.OK;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during rename: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "Rename operation failed";
            }
            finally
            {
                btnRename.Enabled = true;
                btnPreview.Enabled = true;
                progressBar.Visible = false;
            }
        }

        private void PerformRename()
        {
            int processedCount = 0;
            int successCount = 0;

            foreach (DataGridViewRow row in dgvPreview.Rows)
            {
                if (row.DefaultCellStyle.BackColor == System.Drawing.Color.LightGreen)
                {
                    string oldFileName = row.Cells["OldFileName"].Value.ToString();
                    string newFileName = row.Cells["NewFileName"].Value.ToString();
                    string fullPath = row.Cells["FullPath"].Value.ToString();

                    try
                    {
                        this.Invoke(new Action(() =>
                        {
                            lblStatus.Text = $"Renaming: {oldFileName} -> {newFileName}";
                        }));

                        string directory = Path.GetDirectoryName(fullPath);
                        string newFullPath = Path.Combine(directory, newFileName);

                        // Create backup if requested
                        if (chkBackup.Checked)
                        {
                            string backupPath = fullPath + ".backup";
                            File.Copy(fullPath, backupPath, true);
                        }

                        // Rename the file
                        File.Move(fullPath, newFullPath);
                        
                        // Update internal references if it's an assembly file
                        if (Path.GetExtension(newFullPath).Equals(".iam", StringComparison.OrdinalIgnoreCase))
                        {
                            UpdateAssemblyReferences(newFullPath);
                        }

                        successCount++;
                    }
                    catch (Exception ex)
                    {
                        this.Invoke(new Action(() =>
                        {
                            row.Cells["Status"].Value = $"Error: {ex.Message}";
                            row.DefaultCellStyle.BackColor = System.Drawing.Color.LightCoral;
                        }));
                    }
                }

                processedCount++;
                this.Invoke(new Action(() =>
                {
                    progressBar.Value = processedCount;
                }));
            }

            this.Invoke(new Action(() =>
            {
                lblStatus.Text = $"Completed: {successCount} files renamed successfully out of {processedCount} processed";
            }));
        }

        private void UpdateAssemblyReferences(string assemblyPath)
        {
            try
            {
                // This would require opening the assembly in Inventor and updating references
                // For now, we'll just log that this needs to be done
                System.Diagnostics.Debug.WriteLine($"Assembly references need to be updated for: {assemblyPath}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error updating assembly references: {ex.Message}");
            }
        }

        private class RenamePreviewItem
        {
            public string OldFileName { get; set; }
            public string NewFileName { get; set; }
            public string FullPath { get; set; }
            public string Status { get; set; }
            public bool CanRename { get; set; }
        }
    }
}
