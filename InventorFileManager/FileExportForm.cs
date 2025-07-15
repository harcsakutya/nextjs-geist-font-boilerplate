using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;
using Inventor;

namespace InventorFileManager
{
    public partial class FileExportForm : Form
    {
        private Inventor.Application m_inventorApp;
        private TextBox txtSourceFolder;
        private TextBox txtOutputFile;
        private Button btnBrowseSource;
        private Button btnBrowseOutput;
        private Button btnExport;
        private Button btnCancel;
        private CheckBox chkIncludeSubfolders;
        private ProgressBar progressBar;
        private Label lblStatus;

        public FileExportForm(Inventor.Application inventorApp)
        {
            m_inventorApp = inventorApp;
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Export File Names";
            this.Size = new System.Drawing.Size(600, 300);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Source folder selection
            Label lblSource = new Label();
            lblSource.Text = "Source Folder:";
            lblSource.Location = new System.Drawing.Point(20, 20);
            lblSource.Size = new System.Drawing.Size(100, 23);
            this.Controls.Add(lblSource);

            txtSourceFolder = new TextBox();
            txtSourceFolder.Location = new System.Drawing.Point(130, 20);
            txtSourceFolder.Size = new System.Drawing.Size(350, 23);
            this.Controls.Add(txtSourceFolder);

            btnBrowseSource = new Button();
            btnBrowseSource.Text = "Browse";
            btnBrowseSource.Location = new System.Drawing.Point(490, 20);
            btnBrowseSource.Size = new System.Drawing.Size(75, 23);
            btnBrowseSource.Click += BtnBrowseSource_Click;
            this.Controls.Add(btnBrowseSource);

            // Include subfolders checkbox
            chkIncludeSubfolders = new CheckBox();
            chkIncludeSubfolders.Text = "Include Subfolders";
            chkIncludeSubfolders.Location = new System.Drawing.Point(130, 50);
            chkIncludeSubfolders.Size = new System.Drawing.Size(150, 23);
            chkIncludeSubfolders.Checked = true;
            this.Controls.Add(chkIncludeSubfolders);

            // Output file selection
            Label lblOutput = new Label();
            lblOutput.Text = "Output Excel File:";
            lblOutput.Location = new System.Drawing.Point(20, 80);
            lblOutput.Size = new System.Drawing.Size(100, 23);
            this.Controls.Add(lblOutput);

            txtOutputFile = new TextBox();
            txtOutputFile.Location = new System.Drawing.Point(130, 80);
            txtOutputFile.Size = new System.Drawing.Size(350, 23);
            this.Controls.Add(txtOutputFile);

            btnBrowseOutput = new Button();
            btnBrowseOutput.Text = "Browse";
            btnBrowseOutput.Location = new System.Drawing.Point(490, 80);
            btnBrowseOutput.Size = new System.Drawing.Size(75, 23);
            btnBrowseOutput.Click += BtnBrowseOutput_Click;
            this.Controls.Add(btnBrowseOutput);

            // Progress bar
            progressBar = new ProgressBar();
            progressBar.Location = new System.Drawing.Point(20, 120);
            progressBar.Size = new System.Drawing.Size(545, 23);
            progressBar.Visible = false;
            this.Controls.Add(progressBar);

            // Status label
            lblStatus = new Label();
            lblStatus.Location = new System.Drawing.Point(20, 150);
            lblStatus.Size = new System.Drawing.Size(545, 23);
            lblStatus.Text = "Ready to export";
            this.Controls.Add(lblStatus);

            // Buttons
            btnExport = new Button();
            btnExport.Text = "Export";
            btnExport.Location = new System.Drawing.Point(400, 200);
            btnExport.Size = new System.Drawing.Size(75, 30);
            btnExport.Click += BtnExport_Click;
            this.Controls.Add(btnExport);

            btnCancel = new Button();
            btnCancel.Text = "Cancel";
            btnCancel.Location = new System.Drawing.Point(490, 200);
            btnCancel.Size = new System.Drawing.Size(75, 30);
            btnCancel.DialogResult = DialogResult.Cancel;
            this.Controls.Add(btnCancel);

            this.CancelButton = btnCancel;
        }

        private void BtnBrowseSource_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select source folder containing Inventor files";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtSourceFolder.Text = dialog.SelectedPath;
                }
            }
        }

        private void BtnBrowseOutput_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                dialog.DefaultExt = "xlsx";
                dialog.FileName = "exported.xlsx";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtOutputFile.Text = dialog.FileName;
                }
            }
        }

        private async void BtnExport_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtSourceFolder.Text))
            {
                MessageBox.Show("Please select a source folder.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (string.IsNullOrEmpty(txtOutputFile.Text))
            {
                MessageBox.Show("Please specify an output file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!Directory.Exists(txtSourceFolder.Text))
            {
                MessageBox.Show("Source folder does not exist.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                btnExport.Enabled = false;
                progressBar.Visible = true;
                lblStatus.Text = "Scanning files...";

                await System.Threading.Tasks.Task.Run(() => ExportFileNames());

                lblStatus.Text = "Export completed successfully!";
                MessageBox.Show("File names exported successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.DialogResult = DialogResult.OK;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during export: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "Export failed";
            }
            finally
            {
                btnExport.Enabled = true;
                progressBar.Visible = false;
            }
        }

        private void ExportFileNames()
        {
            List<FileInfo> inventorFiles = GetInventorFiles(txtSourceFolder.Text, chkIncludeSubfolders.Checked);
            
            this.Invoke(new Action(() =>
            {
                progressBar.Maximum = inventorFiles.Count;
                progressBar.Value = 0;
            }));

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("File Names");
                
                // Headers
                worksheet.Cells[1, 1].Value = "File Name";
                worksheet.Cells[1, 2].Value = "Full Path";
                worksheet.Cells[1, 3].Value = "File Type";
                worksheet.Cells[1, 4].Value = "Size (KB)";
                worksheet.Cells[1, 5].Value = "Modified Date";
                worksheet.Cells[1, 6].Value = "Relative Path";

                // Style headers
                using (var range = worksheet.Cells[1, 1, 1, 6])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }

                int row = 2;
                foreach (var file in inventorFiles)
                {
                    worksheet.Cells[row, 1].Value = file.Name;
                    worksheet.Cells[row, 2].Value = file.FullName;
                    worksheet.Cells[row, 3].Value = file.Extension.ToUpper();
                    worksheet.Cells[row, 4].Value = Math.Round(file.Length / 1024.0, 2);
                    worksheet.Cells[row, 5].Value = file.LastWriteTime;
                    worksheet.Cells[row, 6].Value = GetRelativePath(txtSourceFolder.Text, file.FullName);

                    this.Invoke(new Action(() =>
                    {
                        progressBar.Value = row - 1;
                        lblStatus.Text = $"Processing: {file.Name}";
                    }));

                    row++;
                }

                // Auto-fit columns
                worksheet.Cells.AutoFitColumns();

                // Save the file
                package.SaveAs(new FileInfo(txtOutputFile.Text));
            }
        }

        private List<FileInfo> GetInventorFiles(string rootPath, bool includeSubfolders)
        {
            List<FileInfo> files = new List<FileInfo>();
            string[] extensions = { "*.iam", "*.ipt", "*.idw" };

            SearchOption searchOption = includeSubfolders ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;

            foreach (string extension in extensions)
            {
                try
                {
                    var foundFiles = Directory.GetFiles(rootPath, extension, searchOption)
                                           .Select(f => new FileInfo(f))
                                           .ToList();
                    files.AddRange(foundFiles);
                }
                catch (Exception ex)
                {
                    // Log error but continue with other extensions
                    System.Diagnostics.Debug.WriteLine($"Error searching for {extension}: {ex.Message}");
                }
            }

            return files.OrderBy(f => f.FullName).ToList();
        }

        private string GetRelativePath(string rootPath, string fullPath)
        {
            Uri rootUri = new Uri(rootPath + Path.DirectorySeparatorChar);
            Uri fullUri = new Uri(fullPath);
            return Uri.UnescapeDataString(rootUri.MakeRelativeUri(fullUri).ToString().Replace('/', Path.DirectorySeparatorChar));
        }
    }
}
