using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace Office_Pragmatics_Word
{
    public static class DocComparison
    {
        #region Entry Point

        public static void RunDocComparison()
        {
            try
            {
                using (var dialog = new ComparisonDialog())
                {
                    dialog.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error launching document comparison: {ex.Message}",
                    "Document Comparison Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Main Comparison Dialog

        private class ComparisonDialog : Form
        {
            // File selection controls
            private TextBox txtOriginalDoc;
            private TextBox txtModifiedDoc;
            private Button btnBrowseOriginal;
            private Button btnBrowseModified;
            private Button btnActiveOriginal;
            private Button btnActiveModified;

            // Output options - always New Document

            // Progress
            private ProgressBar progressBar;
            private Label lblStatus;

            // Action buttons
            private Button btnCompare;
            private Button btnCancel;

            public ComparisonDialog()
            {
                InitializeComponents();
                SetupDragDrop();
                this.FormClosed += (s, e) => CleanupTempFiles();
            }

            private void CleanupTempFiles()
            {
                try
                {
                    string tempDir = Path.Combine(Path.GetTempPath(), "DocComparison");
                    if (Directory.Exists(tempDir))
                    {
                        // Clean up files older than 1 hour
                        foreach (var file in Directory.GetFiles(tempDir))
                        {
                            try
                            {
                                var fileInfo = new FileInfo(file);
                                if (fileInfo.CreationTime < DateTime.Now.AddHours(-1))
                                {
                                    File.Delete(file);
                                }
                            }
                            catch { }
                        }
                    }
                }
                catch { }
            }

            private void InitializeComponents()
            {
                // Form settings
                this.Text = "Document Comparison";
                this.Size = new Size(800, 340);
                this.StartPosition = FormStartPosition.CenterParent;
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.Font = new Font("Segoe UI", 9F);
                this.BackColor = Color.White;

                int yPos = 20;
                int leftMargin = 20;

                // === File Selection Panel ===
                var filePanel = CreateGroupBox("Document Selection", leftMargin, yPos, 740, 130);
                this.Controls.Add(filePanel);

                // Original Document
                var lblOriginal = new Label
                {
                    Text = "Original:",
                    Location = new Point(15, 30),
                    Size = new Size(105, 23),
                    TextAlign = ContentAlignment.MiddleLeft
                };
                filePanel.Controls.Add(lblOriginal);

                txtOriginalDoc = new TextBox
                {
                    Location = new Point(125, 28),
                    Size = new Size(355, 23),
                    AllowDrop = true
                };
                filePanel.Controls.Add(txtOriginalDoc);

                btnBrowseOriginal = new Button
                {
                    Text = "Browse",
                    Location = new Point(490, 25),
                    Size = new Size(80, 30),
                    FlatStyle = FlatStyle.Flat
                };
                btnBrowseOriginal.Click += (s, e) => BrowseForDocument(txtOriginalDoc);
                filePanel.Controls.Add(btnBrowseOriginal);

                btnActiveOriginal = new Button
                {
                    Text = "Active",
                    Location = new Point(580, 25),
                    Size = new Size(80, 30),
                    FlatStyle = FlatStyle.Flat
                };
                btnActiveOriginal.Click += (s, e) => SelectActiveFile(txtOriginalDoc);
                filePanel.Controls.Add(btnActiveOriginal);

                // Modified Document
                var lblModified = new Label
                {
                    Text = "Modified:",
                    Location = new Point(15, 65),
                    Size = new Size(105, 23),
                    TextAlign = ContentAlignment.MiddleLeft
                };
                filePanel.Controls.Add(lblModified);

                txtModifiedDoc = new TextBox
                {
                    Location = new Point(125, 63),
                    Size = new Size(355, 23),
                    AllowDrop = true
                };
                filePanel.Controls.Add(txtModifiedDoc);

                btnBrowseModified = new Button
                {
                    Text = "Browse",
                    Location = new Point(490, 60),
                    Size = new Size(80, 30),
                    FlatStyle = FlatStyle.Flat
                };
                btnBrowseModified.Click += (s, e) => BrowseForDocument(txtModifiedDoc);
                filePanel.Controls.Add(btnBrowseModified);

                btnActiveModified = new Button
                {
                    Text = "Active",
                    Location = new Point(580, 60),
                    Size = new Size(80, 30),
                    FlatStyle = FlatStyle.Flat
                };
                btnActiveModified.Click += (s, e) => SelectActiveFile(txtModifiedDoc);
                filePanel.Controls.Add(btnActiveModified);

                // Drag & drop hint
                var lblDragHint = new Label
                {
                    Text = "Tip: You can drag and drop DOCX, DOC, and PDF files into the fields above",
                    Location = new Point(15, 98),
                    Size = new Size(500, 20),
                    ForeColor = Color.Gray,
                    Font = new Font("Segoe UI", 8F, FontStyle.Italic)
                };
                filePanel.Controls.Add(lblDragHint);

                yPos += 145;

                // === Progress Section ===
                progressBar = new ProgressBar
                {
                    Location = new Point(leftMargin, yPos),
                    Size = new Size(740, 23),
                    Style = ProgressBarStyle.Continuous,
                    Visible = false
                };
                this.Controls.Add(progressBar);

                lblStatus = new Label
                {
                    Location = new Point(leftMargin, yPos + 28),
                    Size = new Size(740, 30),
                    Text = "",
                    ForeColor = Color.Gray
                };
                this.Controls.Add(lblStatus);

                yPos += 70;

                // === Action Buttons ===
                btnCompare = new Button
                {
                    Text = "Compare",
                    Location = new Point(555, yPos + 5),
                    Size = new Size(95, 32),
                    BackColor = Color.FromArgb(0, 120, 212),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 9F, FontStyle.Bold)
                };
                btnCompare.FlatAppearance.BorderSize = 0;
                btnCompare.Click += BtnCompare_Click;
                this.Controls.Add(btnCompare);

                btnCancel = new Button
                {
                    Text = "Cancel",
                    Location = new Point(660, yPos + 5),
                    Size = new Size(95, 32),
                    FlatStyle = FlatStyle.Flat
                };
                btnCancel.Click += (s, e) => this.Close();
                this.Controls.Add(btnCancel);

                this.AcceptButton = btnCompare;
                this.CancelButton = btnCancel;
            }

            private GroupBox CreateGroupBox(string title, int x, int y, int width, int height)
            {
                return new GroupBox
                {
                    Text = title,
                    Location = new Point(x, y),
                    Size = new Size(width, height),
                    Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                    ForeColor = Color.FromArgb(0, 120, 212)
                };
            }

            private void SetupDragDrop()
            {
                txtOriginalDoc.DragEnter += TextBox_DragEnter;
                txtOriginalDoc.DragDrop += TextBox_DragDrop;
                txtModifiedDoc.DragEnter += TextBox_DragEnter;
                txtModifiedDoc.DragDrop += TextBox_DragDrop;
            }

            private void TextBox_DragEnter(object sender, DragEventArgs e)
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                    if (files.Length == 1 && IsValidWordDocument(files[0]))
                    {
                        e.Effect = DragDropEffects.Copy;
                        return;
                    }
                }
                e.Effect = DragDropEffects.None;
            }

            private void TextBox_DragDrop(object sender, DragEventArgs e)
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0 && sender is TextBox textBox)
                {
                    textBox.Text = files[0];
                }
            }

            private bool IsValidWordDocument(string path)
            {
                var ext = Path.GetExtension(path)?.ToLowerInvariant();
                return ext == ".docx" || ext == ".doc" || ext == ".pdf";
            }

            private void BrowseForDocument(TextBox targetTextBox)
            {
                using (var ofd = new OpenFileDialog())
                {
                    ofd.Filter = "Word Documents & PDFs (*.docx;*.doc;*.pdf)|*.docx;*.doc;*.pdf|Word Documents (*.docx;*.doc)|*.docx;*.doc|PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*";
                    ofd.Title = "Select Document";

                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        targetTextBox.Text = ofd.FileName;
                    }
                }
            }

            private void SelectActiveFile(TextBox targetTextBox)
            {
                try
                {
                    var wordApp = Globals.ThisAddIn.Application;
                    var openDocs = new List<OpenDocumentInfo>();

                    // Get all open documents
                    foreach (Word.Document doc in wordApp.Documents)
                    {
                        var docInfo = new OpenDocumentInfo
                        {
                            FileName = doc.Name,
                            Document = doc
                        };

                        // Check if document has been saved
                        if (!string.IsNullOrEmpty(doc.Path))
                        {
                            docInfo.FullPath = doc.FullName;
                            docInfo.IsSaved = true;
                        }
                        else
                        {
                            docInfo.FileName = doc.Name + " (unsaved)";
                            docInfo.IsSaved = false;
                        }

                        openDocs.Add(docInfo);
                    }

                    if (openDocs.Count == 0)
                    {
                        MessageBox.Show("No documents are currently open in Word.",
                            "No Open Documents", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    OpenDocumentInfo selectedDoc = null;

                    // If only one document, auto-select it
                    if (openDocs.Count == 1)
                    {
                        selectedDoc = openDocs[0];
                    }
                    else
                    {
                        // Show selection dialog
                        string activeDocName = wordApp.ActiveDocument?.Name;
                        using (var selectDialog = new ActiveFileSelectDialog(openDocs, activeDocName))
                        {
                            if (selectDialog.ShowDialog() == DialogResult.OK && selectDialog.SelectedDocument != null)
                            {
                                selectedDoc = selectDialog.SelectedDocument;
                            }
                        }
                    }

                    if (selectedDoc != null)
                    {
                        if (!selectedDoc.IsSaved)
                        {
                            MessageBox.Show("Please save the document first before using it for comparison.",
                                "Unsaved Document", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        // For open documents, save a temp copy to avoid file lock issues
                        string tempPath = SaveDocumentToTemp(selectedDoc.Document);
                        if (!string.IsNullOrEmpty(tempPath))
                        {
                            targetTextBox.Text = tempPath;
                            targetTextBox.Tag = selectedDoc.FullPath; // Store original path for reference
                        }
                        else
                        {
                            MessageBox.Show("Failed to create a temporary copy of the document.",
                                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error accessing open documents: {ex.Message}",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            private string SaveDocumentToTemp(Word.Document doc)
            {
                string tempDir = Path.Combine(Path.GetTempPath(), "DocComparison");
                if (!Directory.Exists(tempDir))
                {
                    Directory.CreateDirectory(tempDir);
                }

                string tempPath = Path.Combine(tempDir, $"{Path.GetFileNameWithoutExtension(doc.Name)}_{Guid.NewGuid():N}.docx");

                try
                {
                    // First, try to copy the file directly if it's a local file
                    if (!string.IsNullOrEmpty(doc.Path) && File.Exists(doc.FullName))
                    {
                        // Save any pending changes first
                        if (!doc.Saved)
                        {
                            doc.Save();
                        }

                        // Copy the file
                        File.Copy(doc.FullName, tempPath, true);
                        return tempPath;
                    }
                }
                catch
                {
                    // File copy failed, try alternative method
                }

                try
                {
                    // Alternative: Create new document and copy content
                    var wordApp = doc.Application;
                    var newDoc = wordApp.Documents.Add(Visible: false);

                    // Copy all content
                    doc.Content.Copy();
                    newDoc.Content.Paste();

                    // Save and close the new document
                    newDoc.SaveAs2(tempPath, Word.WdSaveFormat.wdFormatXMLDocument, AddToRecentFiles: false);
                    newDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);

                    return tempPath;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error saving temp document: {ex.Message}");
                    return null;
                }
            }

            private async void BtnCompare_Click(object sender, EventArgs e)
            {
                // Validate inputs
                if (!ValidateInputs()) return;

                // Disable UI during comparison
                SetUIEnabled(false);
                progressBar.Visible = true;
                progressBar.Value = 0;

                try
                {
                    // Get display names - use Tag if set (original name from active file), otherwise use filename
                    string originalDisplayName = txtOriginalDoc.Tag as string;
                    if (string.IsNullOrEmpty(originalDisplayName))
                        originalDisplayName = Path.GetFileName(txtOriginalDoc.Text);
                    else
                        originalDisplayName = Path.GetFileName(originalDisplayName);

                    string modifiedDisplayName = txtModifiedDoc.Tag as string;
                    string originalModifiedPath = txtModifiedDoc.Tag as string; // Original path for output directory
                    if (string.IsNullOrEmpty(modifiedDisplayName))
                        modifiedDisplayName = Path.GetFileName(txtModifiedDoc.Text);
                    else
                        modifiedDisplayName = Path.GetFileName(modifiedDisplayName);

                    var options = new ComparisonOptions
                    {
                        OriginalPath = txtOriginalDoc.Text,
                        ModifiedPath = txtModifiedDoc.Text,
                        OriginalDisplayName = originalDisplayName,
                        ModifiedDisplayName = modifiedDisplayName,
                        OriginalModifiedPath = originalModifiedPath,
                        OutputDestination = OutputDestination.NewDocument
                    };

                    var engine = new ComparisonEngine();
                    engine.ProgressChanged += (progress, status) =>
                    {
                        this.Invoke((Action)(() =>
                        {
                            progressBar.Value = Math.Min(progress, 100);
                            lblStatus.Text = status;
                        }));
                    };

                    var result = await Task.Run(() => engine.Compare(options));

                    if (result.Success)
                    {
                        progressBar.Value = 100;
                        lblStatus.Text = "Comparison complete!";

                        // Open the Word redline first
                        if (!string.IsNullOrEmpty(result.OutputPath) && File.Exists(result.OutputPath))
                        {
                            // Open in Word and show task pane
                            var wordApp = Globals.ThisAddIn.Application;
                            var redlineDoc = wordApp.Documents.Open(result.OutputPath);

                            // Show the Changes task pane
                            ShowChangesTaskPane(redlineDoc);
                        }

                        // Then open the PDF so it's on top
                        if (!string.IsNullOrEmpty(result.PdfPath) && File.Exists(result.PdfPath))
                        {
                            // Small delay to ensure Word opens first
                            System.Threading.Thread.Sleep(500);
                            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                            {
                                FileName = result.PdfPath,
                                UseShellExecute = true
                            });
                        }

                        this.Close();
                    }
                    else
                    {
                        throw new Exception(result.ErrorMessage);
                    }
                }
                catch (Exception ex)
                {
                    lblStatus.Text = "Comparison failed.";
                    MessageBox.Show($"Comparison failed:\n\n{ex.Message}",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    SetUIEnabled(true);
                    progressBar.Visible = false;
                }
            }

            private bool ValidateInputs()
            {
                // Check original document
                if (string.IsNullOrWhiteSpace(txtOriginalDoc.Text))
                {
                    MessageBox.Show("Please select an original document.",
                        "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtOriginalDoc.Focus();
                    return false;
                }

                if (!File.Exists(txtOriginalDoc.Text))
                {
                    MessageBox.Show("The original document file was not found.",
                        "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtOriginalDoc.Focus();
                    return false;
                }

                // Check modified document
                if (string.IsNullOrWhiteSpace(txtModifiedDoc.Text))
                {
                    MessageBox.Show("Please select a modified document.",
                        "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtModifiedDoc.Focus();
                    return false;
                }

                if (!File.Exists(txtModifiedDoc.Text))
                {
                    MessageBox.Show("The modified document file was not found.",
                        "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtModifiedDoc.Focus();
                    return false;
                }

                // Check if files are the same (compare actual paths, not temp paths)
                string origPath = txtOriginalDoc.Tag as string ?? txtOriginalDoc.Text;
                string modPath = txtModifiedDoc.Tag as string ?? txtModifiedDoc.Text;
                if (origPath.Equals(modPath, StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("Please select two different documents to compare.",
                        "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                // Check file accessibility (skip for temp files created from active documents)
                bool isOrigTemp = txtOriginalDoc.Text.Contains(Path.GetTempPath());
                bool isModTemp = txtModifiedDoc.Text.Contains(Path.GetTempPath());

                if (!isOrigTemp)
                {
                    var accessError = CheckFileAccess(txtOriginalDoc.Text);
                    if (accessError != null)
                    {
                        MessageBox.Show($"Cannot access original document:\n{accessError}",
                            "File Access Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return false;
                    }
                }

                if (!isModTemp)
                {
                    var accessError = CheckFileAccess(txtModifiedDoc.Text);
                    if (accessError != null)
                    {
                        MessageBox.Show($"Cannot access modified document:\n{accessError}",
                            "File Access Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return false;
                    }
                }

                return true;
            }

            private string CheckFileAccess(string path)
            {
                try
                {
                    // Check if file is a valid Word document or PDF
                    var ext = Path.GetExtension(path)?.ToLowerInvariant();
                    if (ext != ".docx" && ext != ".doc" && ext != ".pdf")
                    {
                        return "File is not a valid Word document (.docx, .doc) or PDF (.pdf)";
                    }

                    // Try to open file for reading
                    using (var fs = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        // Check if it's a valid ZIP (for .docx)
                        if (ext == ".docx")
                        {
                            try
                            {
                                fs.Close();
                                using (var package = Package.Open(path, FileMode.Open, FileAccess.Read))
                                {
                                    var documentUri = new Uri("/word/document.xml", UriKind.Relative);
                                    if (!package.PartExists(documentUri))
                                    {
                                        return "File appears to be corrupted or is not a valid Word document.";
                                    }
                                }
                            }
                            catch (FileFormatException)
                            {
                                return "File appears to be corrupted or password protected.";
                            }
                        }
                        // For PDF, just check that we can read it
                        else if (ext == ".pdf")
                        {
                            // Basic check - file is readable
                            byte[] header = new byte[4];
                            fs.Read(header, 0, 4);
                            // PDF files start with %PDF
                            if (header[0] != 0x25 || header[1] != 0x50 || header[2] != 0x44 || header[3] != 0x46)
                            {
                                return "File does not appear to be a valid PDF.";
                            }
                        }
                    }
                    return null;
                }
                catch (IOException ex)
                {
                    if (ex.Message.Contains("being used by another process"))
                    {
                        return "File is open in another application. Please close it and try again.";
                    }
                    return ex.Message;
                }
                catch (UnauthorizedAccessException)
                {
                    return "Access denied. You may not have permission to read this file.";
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            }

            private void SetUIEnabled(bool enabled)
            {
                txtOriginalDoc.Enabled = enabled;
                txtModifiedDoc.Enabled = enabled;
                btnBrowseOriginal.Enabled = enabled;
                btnBrowseModified.Enabled = enabled;
                btnActiveOriginal.Enabled = enabled;
                btnActiveModified.Enabled = enabled;
                btnCompare.Enabled = enabled;
            }
        }

        #endregion

        #region Active File Selection Dialog

        private class OpenDocumentInfo
        {
            public string FileName { get; set; }
            public string FullPath { get; set; }
            public Word.Document Document { get; set; }
            public bool IsSaved { get; set; }
            public override string ToString() => FileName;
        }

        private class ActiveFileSelectDialog : Form
        {
            private ListBox lstDocuments;
            private Button btnOK;
            private Button btnCancel;
            private List<OpenDocumentInfo> documents;
            public OpenDocumentInfo SelectedDocument { get; private set; }

            public ActiveFileSelectDialog(List<OpenDocumentInfo> openDocuments, string activeDocumentName)
            {
                documents = openDocuments;

                this.Text = "Select Open Document";
                this.Size = new Size(400, 300);
                this.StartPosition = FormStartPosition.CenterParent;
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.Font = new Font("Segoe UI", 9F);
                this.BackColor = Color.White;

                var lblInstruction = new Label
                {
                    Text = "Select a document from the list of currently open files:",
                    Location = new Point(15, 15),
                    Size = new Size(360, 20)
                };
                this.Controls.Add(lblInstruction);

                lstDocuments = new ListBox
                {
                    Location = new Point(15, 40),
                    Size = new Size(355, 170),
                    Font = new Font("Segoe UI", 9F)
                };

                foreach (var doc in openDocuments)
                {
                    lstDocuments.Items.Add(doc);
                }

                // Select active document by default
                if (!string.IsNullOrEmpty(activeDocumentName))
                {
                    for (int i = 0; i < lstDocuments.Items.Count; i++)
                    {
                        var item = lstDocuments.Items[i] as OpenDocumentInfo;
                        if (item != null && item.FileName == activeDocumentName)
                        {
                            lstDocuments.SelectedIndex = i;
                            break;
                        }
                    }
                }

                if (lstDocuments.SelectedIndex < 0 && lstDocuments.Items.Count > 0)
                {
                    lstDocuments.SelectedIndex = 0;
                }

                lstDocuments.DoubleClick += (s, e) =>
                {
                    if (lstDocuments.SelectedItem != null)
                    {
                        SelectedDocument = lstDocuments.SelectedItem as OpenDocumentInfo;
                        DialogResult = DialogResult.OK;
                        Close();
                    }
                };
                this.Controls.Add(lstDocuments);

                btnOK = new Button
                {
                    Text = "OK",
                    Location = new Point(210, 220),
                    Size = new Size(75, 28),
                    DialogResult = DialogResult.OK,
                    BackColor = Color.FromArgb(0, 120, 212),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat
                };
                btnOK.FlatAppearance.BorderSize = 0;
                btnOK.Click += (s, e) => { if (lstDocuments.SelectedItem != null) SelectedDocument = lstDocuments.SelectedItem as OpenDocumentInfo; };
                this.Controls.Add(btnOK);

                btnCancel = new Button
                {
                    Text = "Cancel",
                    Location = new Point(295, 220),
                    Size = new Size(75, 28),
                    DialogResult = DialogResult.Cancel,
                    FlatStyle = FlatStyle.Flat
                };
                this.Controls.Add(btnCancel);

                this.AcceptButton = btnOK;
                this.CancelButton = btnCancel;
            }
        }

        #endregion

        #region Comparison Options and Results

        public enum OutputDestination
        {
            NewDocument,
            OriginalDocument,
            ModifiedDocument
        }

        public class ComparisonOptions
        {
            public string OriginalPath { get; set; }
            public string ModifiedPath { get; set; }
            public string OriginalDisplayName { get; set; }
            public string ModifiedDisplayName { get; set; }
            public string OriginalModifiedPath { get; set; } // Original path before temp copy (for output directory)
            public OutputDestination OutputDestination { get; set; }
        }

        public class ComparisonResult
        {
            public bool Success { get; set; }
            public string ErrorMessage { get; set; }
            public string OutputPath { get; set; }
            public string PdfPath { get; set; }
            public ChangeSummary Summary { get; set; }
        }

        public class ChangeSummary
        {
            public int Insertions { get; set; }      // Add
            public int Deletions { get; set; }       // Delete
            public int MoveFrom { get; set; }        // Move From (green strikethrough)
            public int MoveTo { get; set; }          // Move To (green underline)
            public int TableInserts { get; set; }    // Table Insert
            public int TableDeletes { get; set; }    // Table Delete
            public int FormatChanges { get; set; }   // Format changes
            public int TotalChanges => Insertions + Deletions + MoveFrom + MoveTo + TableInserts + TableDeletes + FormatChanges;
            public List<ChangeDetail> Details { get; set; } = new List<ChangeDetail>();
        }

        public class ChangeDetail
        {
            public ChangeType Type { get; set; }
            public string OriginalText { get; set; }
            public string ModifiedText { get; set; }
            public int ParagraphIndex { get; set; }
            public string Location { get; set; }
        }

        public enum ChangeType
        {
            Insertion,
            Deletion,
            Move,
            FormatChange,
            TableInsert,
            TableDelete,
            TableCellChange
        }

        #endregion

        #region Comparison Engine

        private class ComparisonEngine
        {
            public event Action<int, string> ProgressChanged;

            private void ReportProgress(int percent, string status)
            {
                ProgressChanged?.Invoke(percent, status);
            }

            public ComparisonResult Compare(ComparisonOptions options)
            {
                try
                {
                    ReportProgress(5, "Initializing comparison...");

                    // Extract document content
                    ReportProgress(10, "Extracting original document content...");
                    var originalContent = ExtractDocumentContent(options.OriginalPath, true);

                    ReportProgress(25, "Extracting modified document content...");
                    var modifiedContent = ExtractDocumentContent(options.ModifiedPath, true);

                    // Perform comparison
                    ReportProgress(40, "Comparing documents...");
                    var changes = CompareDocuments(originalContent, modifiedContent);

                    // Generate summary
                    ReportProgress(60, "Generating change summary...");
                    var summary = GenerateSummary(changes);

                    // Generate redline document
                    ReportProgress(70, "Generating redline document...");
                    var outputPath = GenerateRedlineDocument(options, originalContent, modifiedContent, changes, summary);

                    // Always generate PDF
                    ReportProgress(85, "Generating PDF report...");
                    string pdfPath = GeneratePdfReport(outputPath, summary, options);

                    ReportProgress(100, "Complete!");

                    return new ComparisonResult
                    {
                        Success = true,
                        OutputPath = outputPath,
                        PdfPath = pdfPath,
                        Summary = summary
                    };
                }
                catch (Exception ex)
                {
                    return new ComparisonResult
                    {
                        Success = false,
                        ErrorMessage = ex.Message,
                        Summary = new ChangeSummary()
                    };
                }
            }

            #region Document Content Extraction

            private DocumentContent ExtractDocumentContent(string path, bool cleanMetadata)
            {
                var content = new DocumentContent();
                var ext = Path.GetExtension(path)?.ToLowerInvariant();

                if (ext == ".docx")
                {
                    ExtractDocxContent(path, content, cleanMetadata);
                }
                else if (ext == ".doc")
                {
                    ExtractDocContent(path, content, cleanMetadata);
                }

                return content;
            }

            private void ExtractDocxContent(string path, DocumentContent content, bool cleanMetadata)
            {
                using (var package = Package.Open(path, FileMode.Open, FileAccess.Read))
                {
                    // Extract main document content
                    var documentUri = new Uri("/word/document.xml", UriKind.Relative);
                    if (package.PartExists(documentUri))
                    {
                        var documentPart = package.GetPart(documentUri);
                        using (var stream = documentPart.GetStream())
                        {
                            var doc = XDocument.Load(stream);
                            var ns = doc.Root.GetNamespaceOfPrefix("w") ??
                                XNamespace.Get("http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                            ExtractParagraphs(doc, ns, content);
                            ExtractTables(doc, ns, content);
                        }
                    }

                    // Extract headers
                    foreach (var part in package.GetParts())
                    {
                        var uriString = part.Uri.ToString();
                        if (uriString.StartsWith("/word/header") && uriString.EndsWith(".xml"))
                        {
                            using (var stream = part.GetStream())
                            {
                                var doc = XDocument.Load(stream);
                                var ns = doc.Root.GetNamespaceOfPrefix("w") ??
                                    XNamespace.Get("http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                                var headerContent = new DocumentContent();
                                ExtractParagraphs(doc, ns, headerContent);
                                content.Headers.Add(Path.GetFileName(uriString), headerContent.Paragraphs);
                            }
                        }
                    }

                    // Extract footers
                    foreach (var part in package.GetParts())
                    {
                        var uriString = part.Uri.ToString();
                        if (uriString.StartsWith("/word/footer") && uriString.EndsWith(".xml"))
                        {
                            using (var stream = part.GetStream())
                            {
                                var doc = XDocument.Load(stream);
                                var ns = doc.Root.GetNamespaceOfPrefix("w") ??
                                    XNamespace.Get("http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                                var footerContent = new DocumentContent();
                                ExtractParagraphs(doc, ns, footerContent);
                                content.Footers.Add(Path.GetFileName(uriString), footerContent.Paragraphs);
                            }
                        }
                    }

                    // Extract styles
                    var stylesUri = new Uri("/word/styles.xml", UriKind.Relative);
                    if (package.PartExists(stylesUri))
                    {
                        var stylesPart = package.GetPart(stylesUri);
                        using (var stream = stylesPart.GetStream())
                        using (var reader = new StreamReader(stream))
                        {
                            content.StylesXml = reader.ReadToEnd();
                        }
                    }

                    // Extract numbering
                    var numberingUri = new Uri("/word/numbering.xml", UriKind.Relative);
                    if (package.PartExists(numberingUri))
                    {
                        var numberingPart = package.GetPart(numberingUri);
                        using (var stream = numberingPart.GetStream())
                        {
                            var doc = XDocument.Load(stream);
                            var ns = doc.Root.GetNamespaceOfPrefix("w") ??
                                XNamespace.Get("http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                            ExtractNumbering(doc, ns, content);
                        }
                    }
                }
            }

            private void ExtractDocContent(string path, DocumentContent content, bool cleanMetadata)
            {
                // For .doc files, we'll use Word Interop
                Word.Application wordApp = null;
                Word.Document doc = null;

                try
                {
                    wordApp = new Word.Application { Visible = false };
                    doc = wordApp.Documents.Open(path, ReadOnly: true);

                    foreach (Word.Paragraph para in doc.Paragraphs)
                    {
                        var paragraphInfo = new ParagraphInfo
                        {
                            Text = para.Range.Text?.TrimEnd('\r', '\a') ?? "",
                            StyleName = GetParagraphStyleName(para),
                            ListLevel = GetListLevel(para),
                            ListNumber = GetListNumber(para),
                            IsNumbered = para.Range.ListFormat?.ListType == Word.WdListType.wdListSimpleNumbering ||
                                        para.Range.ListFormat?.ListType == Word.WdListType.wdListOutlineNumbering
                        };

                        ExtractRunsFromParagraph(para, paragraphInfo);
                        content.Paragraphs.Add(paragraphInfo);
                    }

                    // Extract tables
                    foreach (Word.Table table in doc.Tables)
                    {
                        var tableInfo = new TableInfo();
                        tableInfo.RowCount = table.Rows.Count;
                        tableInfo.ColumnCount = table.Columns.Count;

                        for (int r = 1; r <= table.Rows.Count; r++)
                        {
                            var row = new List<CellInfo>();
                            for (int c = 1; c <= table.Columns.Count; c++)
                            {
                                try
                                {
                                    var cell = table.Cell(r, c);
                                    row.Add(new CellInfo
                                    {
                                        Text = cell.Range.Text?.TrimEnd('\r', '\a', '\x07') ?? "",
                                        Row = r - 1,
                                        Column = c - 1
                                    });
                                }
                                catch
                                {
                                    row.Add(new CellInfo { Text = "", Row = r - 1, Column = c - 1 });
                                }
                            }
                            tableInfo.Cells.Add(row);
                        }
                        content.Tables.Add(tableInfo);
                    }
                }
                finally
                {
                    doc?.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                    wordApp?.Quit();
                    ReleaseComObject(doc);
                    ReleaseComObject(wordApp);
                }
            }

            private void ExtractParagraphs(XDocument doc, XNamespace ns, DocumentContent content)
            {
                var body = doc.Descendants(ns + "body").FirstOrDefault();
                if (body == null) return;

                foreach (var para in body.Descendants(ns + "p"))
                {
                    var paragraphInfo = new ParagraphInfo();

                    // Get paragraph properties
                    var pPr = para.Element(ns + "pPr");
                    if (pPr != null)
                    {
                        var pStyle = pPr.Element(ns + "pStyle");
                        paragraphInfo.StyleName = pStyle?.Attribute(ns + "val")?.Value ?? "";

                        var numPr = pPr.Element(ns + "numPr");
                        if (numPr != null)
                        {
                            var ilvl = numPr.Element(ns + "ilvl");
                            var numId = numPr.Element(ns + "numId");
                            paragraphInfo.ListLevel = int.TryParse(ilvl?.Attribute(ns + "val")?.Value, out int level) ? level : -1;
                            paragraphInfo.NumId = numId?.Attribute(ns + "val")?.Value;
                            paragraphInfo.IsNumbered = !string.IsNullOrEmpty(paragraphInfo.NumId) && paragraphInfo.NumId != "0";
                        }
                    }

                    // Extract runs (text with formatting)
                    var sb = new StringBuilder();
                    foreach (var run in para.Descendants(ns + "r"))
                    {
                        var runInfo = new RunInfo();

                        // Get run properties
                        var rPr = run.Element(ns + "rPr");
                        if (rPr != null)
                        {
                            runInfo.IsBold = rPr.Element(ns + "b") != null;
                            runInfo.IsItalic = rPr.Element(ns + "i") != null;
                            runInfo.IsUnderline = rPr.Element(ns + "u") != null;
                            runInfo.IsStrikethrough = rPr.Element(ns + "strike") != null;

                            var sz = rPr.Element(ns + "sz");
                            if (sz != null && int.TryParse(sz.Attribute(ns + "val")?.Value, out int size))
                            {
                                runInfo.FontSize = size / 2.0; // Word uses half-points
                            }

                            var rFonts = rPr.Element(ns + "rFonts");
                            runInfo.FontName = rFonts?.Attribute(ns + "ascii")?.Value ?? "";

                            var color = rPr.Element(ns + "color");
                            runInfo.Color = color?.Attribute(ns + "val")?.Value ?? "";
                        }

                        // Get text content
                        var texts = run.Elements(ns + "t");
                        foreach (var t in texts)
                        {
                            runInfo.Text += t.Value;
                            sb.Append(t.Value);
                        }

                        // Handle tabs and breaks
                        if (run.Element(ns + "tab") != null)
                        {
                            runInfo.Text += "\t";
                            sb.Append("\t");
                        }
                        if (run.Element(ns + "br") != null)
                        {
                            runInfo.Text += "\n";
                            sb.Append("\n");
                        }

                        if (!string.IsNullOrEmpty(runInfo.Text))
                        {
                            paragraphInfo.Runs.Add(runInfo);
                        }
                    }

                    paragraphInfo.Text = sb.ToString();
                    content.Paragraphs.Add(paragraphInfo);
                }
            }

            private void ExtractTables(XDocument doc, XNamespace ns, DocumentContent content)
            {
                foreach (var tbl in doc.Descendants(ns + "tbl"))
                {
                    var tableInfo = new TableInfo();
                    var rows = tbl.Elements(ns + "tr").ToList();
                    tableInfo.RowCount = rows.Count;

                    foreach (var row in rows)
                    {
                        var rowCells = new List<CellInfo>();
                        var cells = row.Elements(ns + "tc").ToList();
                        tableInfo.ColumnCount = Math.Max(tableInfo.ColumnCount, cells.Count);

                        int colIndex = 0;
                        foreach (var cell in cells)
                        {
                            var cellInfo = new CellInfo
                            {
                                Row = tableInfo.Cells.Count,
                                Column = colIndex++
                            };

                            // Get cell text
                            var sb = new StringBuilder();
                            foreach (var para in cell.Descendants(ns + "p"))
                            {
                                foreach (var t in para.Descendants(ns + "t"))
                                {
                                    sb.Append(t.Value);
                                }
                                sb.Append(" ");
                            }
                            cellInfo.Text = sb.ToString().Trim();

                            // Check for merged cells
                            var tcPr = cell.Element(ns + "tcPr");
                            if (tcPr != null)
                            {
                                var gridSpan = tcPr.Element(ns + "gridSpan");
                                if (gridSpan != null && int.TryParse(gridSpan.Attribute(ns + "val")?.Value, out int span))
                                {
                                    cellInfo.ColSpan = span;
                                }

                                var vMerge = tcPr.Element(ns + "vMerge");
                                if (vMerge != null)
                                {
                                    cellInfo.IsMerged = true;
                                    cellInfo.IsRowMergeStart = vMerge.Attribute(ns + "val")?.Value == "restart";
                                }
                            }

                            rowCells.Add(cellInfo);
                        }
                        tableInfo.Cells.Add(rowCells);
                    }
                    content.Tables.Add(tableInfo);
                }
            }

            private void ExtractNumbering(XDocument doc, XNamespace ns, DocumentContent content)
            {
                var abstractNums = doc.Descendants(ns + "abstractNum").ToList();
                var nums = doc.Descendants(ns + "num").ToList();

                foreach (var num in nums)
                {
                    var numId = num.Attribute(ns + "numId")?.Value;
                    var abstractNumIdRef = num.Element(ns + "abstractNumId")?.Attribute(ns + "val")?.Value;

                    if (numId != null && abstractNumIdRef != null)
                    {
                        var abstractNum = abstractNums.FirstOrDefault(a =>
                            a.Attribute(ns + "abstractNumId")?.Value == abstractNumIdRef);

                        if (abstractNum != null)
                        {
                            var levels = abstractNum.Elements(ns + "lvl").ToList();
                            var numInfo = new NumberingInfo { NumId = numId };

                            foreach (var lvl in levels)
                            {
                                var ilvl = lvl.Attribute(ns + "ilvl")?.Value ?? "0";
                                var numFmt = lvl.Element(ns + "numFmt")?.Attribute(ns + "val")?.Value ?? "decimal";
                                var lvlText = lvl.Element(ns + "lvlText")?.Attribute(ns + "val")?.Value ?? "";

                                numInfo.Levels[ilvl] = new NumberingLevelInfo
                                {
                                    Level = int.Parse(ilvl),
                                    Format = numFmt,
                                    Text = lvlText
                                };
                            }
                            content.Numbering[numId] = numInfo;
                        }
                    }
                }
            }

            private int GetListLevel(Word.Paragraph para)
            {
                try
                {
                    return para.Range.ListFormat?.ListLevelNumber ?? -1;
                }
                catch
                {
                    return -1;
                }
            }

            private string GetListNumber(Word.Paragraph para)
            {
                try
                {
                    return para.Range.ListFormat?.ListString ?? "";
                }
                catch
                {
                    return "";
                }
            }

            private string GetParagraphStyleName(Word.Paragraph para)
            {
                try
                {
                    object style = para.get_Style();
                    if (style is Word.Style wordStyle)
                    {
                        return wordStyle.NameLocal ?? "";
                    }
                    return "";
                }
                catch
                {
                    return "";
                }
            }

            private void ExtractRunsFromParagraph(Word.Paragraph para, ParagraphInfo paragraphInfo)
            {
                try
                {
                    foreach (Word.Range run in para.Range.Words)
                    {
                        var runInfo = new RunInfo
                        {
                            Text = run.Text,
                            IsBold = run.Bold == -1,
                            IsItalic = run.Italic == -1,
                            IsUnderline = run.Underline != Word.WdUnderline.wdUnderlineNone,
                            FontName = run.Font?.Name ?? "",
                            FontSize = run.Font?.Size ?? 0
                        };
                        paragraphInfo.Runs.Add(runInfo);
                    }
                }
                catch
                {
                    // Fall back to just using paragraph text
                }
            }

            #endregion

            #region Document Comparison Algorithm

            private List<Change> CompareDocuments(DocumentContent original, DocumentContent modified)
            {
                var changes = new List<Change>();

                // Compare paragraphs using LCS-based diff
                var paragraphChanges = CompareParagraphs(original.Paragraphs, modified.Paragraphs);
                changes.AddRange(paragraphChanges);

                // Detect moved paragraphs
                DetectMovedParagraphs(changes, original.Paragraphs, modified.Paragraphs);

                // Compare tables
                var tableChanges = CompareTables(original.Tables, modified.Tables);
                changes.AddRange(tableChanges);

                // Compare headers
                foreach (var header in original.Headers)
                {
                    if (modified.Headers.TryGetValue(header.Key, out var modHeader))
                    {
                        var headerChanges = CompareParagraphs(header.Value, modHeader);
                        foreach (var change in headerChanges)
                        {
                            change.Location = $"Header ({header.Key})";
                        }
                        changes.AddRange(headerChanges);
                    }
                }

                // Compare footers
                foreach (var footer in original.Footers)
                {
                    if (modified.Footers.TryGetValue(footer.Key, out var modFooter))
                    {
                        var footerChanges = CompareParagraphs(footer.Value, modFooter);
                        foreach (var change in footerChanges)
                        {
                            change.Location = $"Footer ({footer.Key})";
                        }
                        changes.AddRange(footerChanges);
                    }
                }

                return changes;
            }

            private List<Change> CompareParagraphs(List<ParagraphInfo> original, List<ParagraphInfo> modified)
            {
                var changes = new List<Change>();

                // Use Myers diff algorithm for paragraph-level comparison
                var lcs = ComputeLCS(original, modified, (a, b) => ParagraphsSimilar(a, b, 0.8));

                int origIndex = 0;
                int modIndex = 0;
                int lcsIndex = 0;

                while (origIndex < original.Count || modIndex < modified.Count)
                {
                    if (lcsIndex < lcs.Count &&
                        origIndex < original.Count && modIndex < modified.Count &&
                        ParagraphsSimilar(original[origIndex], lcs[lcsIndex].OriginalPara, 0.95) &&
                        ParagraphsSimilar(modified[modIndex], lcs[lcsIndex].ModifiedPara, 0.95))
                    {
                        // Check for word-level changes within matched paragraphs
                        if (!original[origIndex].Text.Equals(modified[modIndex].Text))
                        {
                            var wordChanges = CompareWords(original[origIndex], modified[modIndex], origIndex);
                            changes.AddRange(wordChanges);
                        }

                        // Check for formatting changes
                        var formatChanges = CompareFormatting(original[origIndex], modified[modIndex], origIndex);
                        changes.AddRange(formatChanges);

                        // Check for list number changes
                        if (original[origIndex].IsNumbered || modified[modIndex].IsNumbered)
                        {
                            if (original[origIndex].ListNumber != modified[modIndex].ListNumber)
                            {
                                changes.Add(new Change
                                {
                                    Type = ChangeType.FormatChange,
                                    OriginalText = $"List: {original[origIndex].ListNumber}",
                                    ModifiedText = $"List: {modified[modIndex].ListNumber}",
                                    ParagraphIndex = origIndex,
                                    Location = $"Paragraph {origIndex + 1}"
                                });
                            }
                        }

                        origIndex++;
                        modIndex++;
                        lcsIndex++;
                    }
                    else if (modIndex >= modified.Count ||
                             (origIndex < original.Count && lcsIndex < lcs.Count &&
                              !ParagraphsSimilar(original[origIndex], lcs[lcsIndex].OriginalPara, 0.8)))
                    {
                        // Deletion
                        changes.Add(new Change
                        {
                            Type = ChangeType.Deletion,
                            OriginalText = original[origIndex].Text,
                            OriginalParagraph = original[origIndex],
                            ParagraphIndex = origIndex,
                            Location = $"Paragraph {origIndex + 1}"
                        });
                        origIndex++;
                    }
                    else
                    {
                        // Insertion
                        changes.Add(new Change
                        {
                            Type = ChangeType.Insertion,
                            ModifiedText = modified[modIndex].Text,
                            ModifiedParagraph = modified[modIndex],
                            ParagraphIndex = modIndex,
                            Location = $"Paragraph {modIndex + 1}"
                        });
                        modIndex++;
                    }
                }

                return changes;
            }

            private List<LCSMatch> ComputeLCS<T>(List<T> original, List<T> modified, Func<T, T, bool> equals)
            {
                int m = original.Count;
                int n = modified.Count;
                var dp = new int[m + 1, n + 1];

                // Build LCS matrix
                for (int i = 1; i <= m; i++)
                {
                    for (int j = 1; j <= n; j++)
                    {
                        if (equals(original[i - 1], modified[j - 1]))
                        {
                            dp[i, j] = dp[i - 1, j - 1] + 1;
                        }
                        else
                        {
                            dp[i, j] = Math.Max(dp[i - 1, j], dp[i, j - 1]);
                        }
                    }
                }

                // Backtrack to find LCS
                var result = new List<LCSMatch>();
                int x = m, y = n;
                while (x > 0 && y > 0)
                {
                    if (equals(original[x - 1], modified[y - 1]))
                    {
                        if (original[x - 1] is ParagraphInfo origPara && modified[y - 1] is ParagraphInfo modPara)
                        {
                            result.Insert(0, new LCSMatch
                            {
                                OriginalIndex = x - 1,
                                ModifiedIndex = y - 1,
                                OriginalPara = origPara,
                                ModifiedPara = modPara
                            });
                        }
                        x--;
                        y--;
                    }
                    else if (dp[x - 1, y] > dp[x, y - 1])
                    {
                        x--;
                    }
                    else
                    {
                        y--;
                    }
                }

                return result;
            }

            private bool ParagraphsSimilar(ParagraphInfo a, ParagraphInfo b, double threshold)
            {
                if (string.IsNullOrEmpty(a.Text) && string.IsNullOrEmpty(b.Text))
                    return true;

                if (string.IsNullOrEmpty(a.Text) || string.IsNullOrEmpty(b.Text))
                    return false;

                var similarity = ComputeSimilarity(a.Text, b.Text);
                return similarity >= threshold;
            }

            private double ComputeSimilarity(string a, string b)
            {
                if (a == b) return 1.0;
                if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b)) return 0.0;

                // Use Levenshtein distance-based similarity
                int distance = ComputeLevenshteinDistance(a, b);
                int maxLength = Math.Max(a.Length, b.Length);
                return 1.0 - (double)distance / maxLength;
            }

            private int ComputeLevenshteinDistance(string a, string b)
            {
                int[,] dp = new int[a.Length + 1, b.Length + 1];

                for (int i = 0; i <= a.Length; i++) dp[i, 0] = i;
                for (int j = 0; j <= b.Length; j++) dp[0, j] = j;

                for (int i = 1; i <= a.Length; i++)
                {
                    for (int j = 1; j <= b.Length; j++)
                    {
                        int cost = a[i - 1] == b[j - 1] ? 0 : 1;
                        dp[i, j] = Math.Min(Math.Min(
                            dp[i - 1, j] + 1,       // deletion
                            dp[i, j - 1] + 1),      // insertion
                            dp[i - 1, j - 1] + cost // substitution
                        );
                    }
                }

                return dp[a.Length, b.Length];
            }

            private List<Change> CompareWords(ParagraphInfo original, ParagraphInfo modified, int paragraphIndex)
            {
                var changes = new List<Change>();

                var origWords = TokenizeText(original.Text);
                var modWords = TokenizeText(modified.Text);

                var wordLcs = ComputeWordLCS(origWords, modWords);

                int origIdx = 0;
                int modIdx = 0;
                int lcsIdx = 0;

                var deletions = new StringBuilder();
                var insertions = new StringBuilder();
                int deleteStart = -1;
                int insertStart = -1;

                while (origIdx < origWords.Count || modIdx < modWords.Count)
                {
                    if (lcsIdx < wordLcs.Count &&
                        origIdx < origWords.Count && modIdx < modWords.Count &&
                        origWords[origIdx] == wordLcs[lcsIdx] &&
                        modWords[modIdx] == wordLcs[lcsIdx])
                    {
                        // Flush pending changes
                        if (deletions.Length > 0 || insertions.Length > 0)
                        {
                            changes.Add(new Change
                            {
                                Type = deletions.Length > 0 && insertions.Length > 0
                                    ? ChangeType.Deletion  // Will be paired with insertion
                                    : (deletions.Length > 0 ? ChangeType.Deletion : ChangeType.Insertion),
                                OriginalText = deletions.ToString().Trim(),
                                ModifiedText = insertions.ToString().Trim(),
                                ParagraphIndex = paragraphIndex,
                                WordStartIndex = deleteStart >= 0 ? deleteStart : insertStart,
                                Location = $"Paragraph {paragraphIndex + 1}"
                            });
                            deletions.Clear();
                            insertions.Clear();
                            deleteStart = -1;
                            insertStart = -1;
                        }
                        origIdx++;
                        modIdx++;
                        lcsIdx++;
                    }
                    else if (modIdx >= modWords.Count ||
                             (origIdx < origWords.Count && lcsIdx < wordLcs.Count &&
                              origWords[origIdx] != wordLcs[lcsIdx]))
                    {
                        if (deleteStart < 0) deleteStart = origIdx;
                        deletions.Append(origWords[origIdx] + " ");
                        origIdx++;
                    }
                    else
                    {
                        if (insertStart < 0) insertStart = modIdx;
                        insertions.Append(modWords[modIdx] + " ");
                        modIdx++;
                    }
                }

                // Flush remaining changes
                if (deletions.Length > 0 || insertions.Length > 0)
                {
                    changes.Add(new Change
                    {
                        Type = deletions.Length > 0 ? ChangeType.Deletion : ChangeType.Insertion,
                        OriginalText = deletions.ToString().Trim(),
                        ModifiedText = insertions.ToString().Trim(),
                        ParagraphIndex = paragraphIndex,
                        WordStartIndex = deleteStart >= 0 ? deleteStart : insertStart,
                        Location = $"Paragraph {paragraphIndex + 1}"
                    });
                }

                return changes;
            }

            private List<string> TokenizeText(string text)
            {
                if (string.IsNullOrEmpty(text)) return new List<string>();

                // Tokenize preserving punctuation as separate tokens
                var tokens = new List<string>();
                var currentToken = new StringBuilder();

                foreach (char c in text)
                {
                    if (char.IsWhiteSpace(c))
                    {
                        if (currentToken.Length > 0)
                        {
                            tokens.Add(currentToken.ToString());
                            currentToken.Clear();
                        }
                    }
                    else if (char.IsPunctuation(c) && c != '\'' && c != '-')
                    {
                        if (currentToken.Length > 0)
                        {
                            tokens.Add(currentToken.ToString());
                            currentToken.Clear();
                        }
                        tokens.Add(c.ToString());
                    }
                    else
                    {
                        currentToken.Append(c);
                    }
                }

                if (currentToken.Length > 0)
                {
                    tokens.Add(currentToken.ToString());
                }

                return tokens;
            }

            private List<string> ComputeWordLCS(List<string> a, List<string> b)
            {
                int m = a.Count;
                int n = b.Count;
                var dp = new int[m + 1, n + 1];

                for (int i = 1; i <= m; i++)
                {
                    for (int j = 1; j <= n; j++)
                    {
                        if (a[i - 1] == b[j - 1])
                        {
                            dp[i, j] = dp[i - 1, j - 1] + 1;
                        }
                        else
                        {
                            dp[i, j] = Math.Max(dp[i - 1, j], dp[i, j - 1]);
                        }
                    }
                }

                var result = new List<string>();
                int x = m, y = n;
                while (x > 0 && y > 0)
                {
                    if (a[x - 1] == b[y - 1])
                    {
                        result.Insert(0, a[x - 1]);
                        x--;
                        y--;
                    }
                    else if (dp[x - 1, y] > dp[x, y - 1])
                    {
                        x--;
                    }
                    else
                    {
                        y--;
                    }
                }

                return result;
            }

            private List<Change> CompareFormatting(ParagraphInfo original, ParagraphInfo modified, int paragraphIndex)
            {
                var changes = new List<Change>();

                // Compare style changes
                if (original.StyleName != modified.StyleName)
                {
                    changes.Add(new Change
                    {
                        Type = ChangeType.FormatChange,
                        OriginalText = $"Style: {original.StyleName}",
                        ModifiedText = $"Style: {modified.StyleName}",
                        ParagraphIndex = paragraphIndex,
                        Location = $"Paragraph {paragraphIndex + 1}"
                    });
                }

                // Compare run-level formatting
                var origRuns = original.Runs;
                var modRuns = modified.Runs;

                for (int i = 0; i < Math.Min(origRuns.Count, modRuns.Count); i++)
                {
                    var origRun = origRuns[i];
                    var modRun = modRuns[i];

                    if (origRun.Text == modRun.Text)
                    {
                        var formatDiffs = new List<string>();

                        if (origRun.IsBold != modRun.IsBold)
                            formatDiffs.Add(modRun.IsBold ? "Bold added" : "Bold removed");
                        if (origRun.IsItalic != modRun.IsItalic)
                            formatDiffs.Add(modRun.IsItalic ? "Italic added" : "Italic removed");
                        if (origRun.IsUnderline != modRun.IsUnderline)
                            formatDiffs.Add(modRun.IsUnderline ? "Underline added" : "Underline removed");
                        if (Math.Abs(origRun.FontSize - modRun.FontSize) > 0.1)
                            formatDiffs.Add($"Font size: {origRun.FontSize} → {modRun.FontSize}");
                        if (origRun.FontName != modRun.FontName && !string.IsNullOrEmpty(modRun.FontName))
                            formatDiffs.Add($"Font: {origRun.FontName} → {modRun.FontName}");

                        if (formatDiffs.Count > 0)
                        {
                            changes.Add(new Change
                            {
                                Type = ChangeType.FormatChange,
                                OriginalText = origRun.Text,
                                ModifiedText = string.Join(", ", formatDiffs),
                                ParagraphIndex = paragraphIndex,
                                Location = $"Paragraph {paragraphIndex + 1}"
                            });
                        }
                    }
                }

                return changes;
            }

            private void DetectMovedParagraphs(List<Change> changes, List<ParagraphInfo> original, List<ParagraphInfo> modified)
            {
                var deletions = changes.Where(c => c.Type == ChangeType.Deletion && c.OriginalParagraph != null).ToList();
                var insertions = changes.Where(c => c.Type == ChangeType.Insertion && c.ModifiedParagraph != null).ToList();

                foreach (var deletion in deletions.ToList())
                {
                    foreach (var insertion in insertions.ToList())
                    {
                        var similarity = ComputeSimilarity(
                            deletion.OriginalText ?? "",
                            insertion.ModifiedText ?? "");

                        if (similarity >= 0.85) // High similarity threshold for moves
                        {
                            // Convert to move
                            changes.Remove(deletion);
                            changes.Remove(insertion);

                            changes.Add(new Change
                            {
                                Type = ChangeType.Move,
                                OriginalText = deletion.OriginalText,
                                ModifiedText = insertion.ModifiedText,
                                ParagraphIndex = deletion.ParagraphIndex,
                                TargetParagraphIndex = insertion.ParagraphIndex,
                                Location = $"Moved from Paragraph {deletion.ParagraphIndex + 1} to {insertion.ParagraphIndex + 1}"
                            });

                            insertions.Remove(insertion);
                            break;
                        }
                    }
                }
            }

            private List<Change> CompareTables(List<TableInfo> original, List<TableInfo> modified)
            {
                var changes = new List<Change>();

                int minTables = Math.Min(original.Count, modified.Count);

                // Compare corresponding tables
                for (int t = 0; t < minTables; t++)
                {
                    var origTable = original[t];
                    var modTable = modified[t];

                    // Check for row/column changes
                    if (origTable.RowCount != modTable.RowCount)
                    {
                        changes.Add(new Change
                        {
                            Type = origTable.RowCount < modTable.RowCount ? ChangeType.TableInsert : ChangeType.TableDelete,
                            OriginalText = $"Table {t + 1}: {origTable.RowCount} rows",
                            ModifiedText = $"Table {t + 1}: {modTable.RowCount} rows",
                            Location = $"Table {t + 1}"
                        });
                    }

                    if (origTable.ColumnCount != modTable.ColumnCount)
                    {
                        changes.Add(new Change
                        {
                            Type = origTable.ColumnCount < modTable.ColumnCount ? ChangeType.TableInsert : ChangeType.TableDelete,
                            OriginalText = $"Table {t + 1}: {origTable.ColumnCount} columns",
                            ModifiedText = $"Table {t + 1}: {modTable.ColumnCount} columns",
                            Location = $"Table {t + 1}"
                        });
                    }

                    // Compare cells
                    int minRows = Math.Min(origTable.Cells.Count, modTable.Cells.Count);
                    for (int r = 0; r < minRows; r++)
                    {
                        int minCols = Math.Min(origTable.Cells[r].Count, modTable.Cells[r].Count);
                        for (int c = 0; c < minCols; c++)
                        {
                            var origCell = origTable.Cells[r][c];
                            var modCell = modTable.Cells[r][c];

                            if (origCell.Text != modCell.Text)
                            {
                                changes.Add(new Change
                                {
                                    Type = ChangeType.TableCellChange,
                                    OriginalText = origCell.Text,
                                    ModifiedText = modCell.Text,
                                    Location = $"Table {t + 1}, Cell [{r + 1},{c + 1}]"
                                });
                            }
                        }
                    }
                }

                // Handle added tables
                for (int t = minTables; t < modified.Count; t++)
                {
                    changes.Add(new Change
                    {
                        Type = ChangeType.TableInsert,
                        ModifiedText = $"New table with {modified[t].RowCount} rows × {modified[t].ColumnCount} columns",
                        Location = $"Table {t + 1}"
                    });
                }

                // Handle deleted tables
                for (int t = minTables; t < original.Count; t++)
                {
                    changes.Add(new Change
                    {
                        Type = ChangeType.TableDelete,
                        OriginalText = $"Deleted table with {original[t].RowCount} rows × {original[t].ColumnCount} columns",
                        Location = $"Table {t + 1}"
                    });
                }

                return changes;
            }

            #endregion

            #region Summary Generation

            private ChangeSummary GenerateSummary(List<Change> changes)
            {
                var summary = new ChangeSummary();

                foreach (var change in changes)
                {
                    switch (change.Type)
                    {
                        case ChangeType.Insertion:
                            summary.Insertions++;
                            break;
                        case ChangeType.Deletion:
                            summary.Deletions++;
                            break;
                        case ChangeType.Move:
                            summary.MoveFrom++;
                            summary.MoveTo++;
                            break;
                        case ChangeType.FormatChange:
                            summary.FormatChanges++;
                            break;
                        case ChangeType.TableInsert:
                            summary.TableInserts++;
                            break;
                        case ChangeType.TableDelete:
                            summary.TableDeletes++;
                            break;
                        case ChangeType.TableCellChange:
                            summary.Insertions++; // Cell changes count as insertions
                            break;
                    }

                    summary.Details.Add(new ChangeDetail
                    {
                        Type = change.Type,
                        OriginalText = TruncateText(change.OriginalText, 100),
                        ModifiedText = TruncateText(change.ModifiedText, 100),
                        ParagraphIndex = change.ParagraphIndex,
                        Location = change.Location
                    });
                }

                return summary;
            }

            private string TruncateText(string text, int maxLength)
            {
                if (string.IsNullOrEmpty(text)) return "";
                if (text.Length <= maxLength) return text;
                return text.Substring(0, maxLength - 3) + "...";
            }

            #endregion

            #region Redline Document Generation

            private string GenerateRedlineDocument(ComparisonOptions options, DocumentContent original,
                DocumentContent modified, List<Change> changes, ChangeSummary summary)
            {
                Word.Application wordApp = null;
                Word.Document comparedDoc = null;
                Word.Document originalDoc = null;
                Word.Document modifiedDoc = null;
                string tempOriginalPath = null;
                string tempModifiedPath = null;

                try
                {
                    wordApp = new Word.Application { Visible = false };

                    // Convert PDFs to DOCX if needed
                    string originalPath = options.OriginalPath;
                    string modifiedPath = options.ModifiedPath;

                    if (Path.GetExtension(options.OriginalPath)?.ToLowerInvariant() == ".pdf")
                    {
                        tempOriginalPath = ConvertPdfToDocx(wordApp, options.OriginalPath);
                        originalPath = tempOriginalPath;
                    }

                    if (Path.GetExtension(options.ModifiedPath)?.ToLowerInvariant() == ".pdf")
                    {
                        tempModifiedPath = ConvertPdfToDocx(wordApp, options.ModifiedPath);
                        modifiedPath = tempModifiedPath;
                    }

                    // Determine output path - use "Redline - Original File Name" format
                    string originalFileName = options.OriginalDisplayName != null
                        ? Path.GetFileNameWithoutExtension(options.OriginalDisplayName)
                        : Path.GetFileNameWithoutExtension(options.OriginalPath);

                    // Use original modified path directory if available (for active files), otherwise use modified path directory
                    string outputDir;
                    if (!string.IsNullOrEmpty(options.OriginalModifiedPath))
                    {
                        outputDir = Path.GetDirectoryName(options.OriginalModifiedPath);
                    }
                    else
                    {
                        outputDir = Path.GetDirectoryName(options.ModifiedPath);
                    }

                    string outputPath = Path.Combine(outputDir, $"Redline - {originalFileName}.docx");

                    // Make unique if file exists
                    if (File.Exists(outputPath))
                    {
                        int counter = 1;
                        string basePath = Path.Combine(outputDir, $"Redline - {originalFileName}");
                        while (File.Exists(outputPath))
                        {
                            outputPath = $"{basePath} ({counter}).docx";
                            counter++;
                        }
                    }

                    // Open documents
                    originalDoc = wordApp.Documents.Open(originalPath, ReadOnly: true, AddToRecentFiles: false);
                    modifiedDoc = wordApp.Documents.Open(modifiedPath, ReadOnly: true, AddToRecentFiles: false);

                    // Use Word's built-in compare feature
                    comparedDoc = wordApp.CompareDocuments(
                        originalDoc,
                        modifiedDoc,
                        Word.WdCompareDestination.wdCompareDestinationNew,
                        Word.WdGranularity.wdGranularityWordLevel,
                        true,   // CompareFormatting
                        true,   // CompareCaseChanges
                        true,   // CompareWhitespace
                        true,   // CompareTables
                        true,   // CompareHeaders
                        true,   // CompareFootnotes
                        true,   // CompareTextboxes
                        true,   // CompareFields
                        true,   // CompareComments
                        true    // RevisedAuthor - track moves
                    );

                    // Set revision colors - Blue for insertions, Red for deletions
                    wordApp.Options.InsertedTextColor = Word.WdColorIndex.wdBlue;
                    wordApp.Options.DeletedTextColor = Word.WdColorIndex.wdRed;
                    wordApp.Options.InsertedTextMark = Word.WdInsertedTextMark.wdInsertedTextMarkUnderline;
                    wordApp.Options.DeletedTextMark = Word.WdDeletedTextMark.wdDeletedTextMarkStrikeThrough;

                    // Enable changed lines in margin (the vertical bar indicating changes)
                    wordApp.Options.RevisedLinesColor = Word.WdColorIndex.wdAuto;
                    wordApp.Options.RevisedLinesMark = Word.WdRevisedLinesMark.wdRevisedLinesMarkOutsideBorder;

                    // Set view to show inline markup only (no balloons)
                    var view = comparedDoc.ActiveWindow.View;
                    view.RevisionsFilter.Markup = Word.WdRevisionsMarkup.wdRevisionsMarkupAll;
                    view.ShowRevisionsAndComments = true;
                    view.RevisionsMode = Word.WdRevisionsMode.wdInLineRevisions;

                    // Ensure changed lines bar shows in margin
                    try
                    {
                        view.ShowInsertionsAndDeletions = true;
                        view.ShowFormatChanges = true;
                    }
                    catch { /* Ignore if not supported */ }

                    // Save the compared document (no summary table in Word version - only in PDF)
                    comparedDoc.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);

                    // Close all open documents
                    originalDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                    modifiedDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                    comparedDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);

                    return outputPath;
                }
                catch (Exception ex)
                {
                    // Clean up on error
                    try
                    {
                        originalDoc?.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                        modifiedDoc?.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                        comparedDoc?.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                    }
                    catch { }

                    throw new Exception($"Failed to generate redline document: {ex.Message}", ex);
                }
                finally
                {
                    // Clean up temp files
                    try
                    {
                        if (tempOriginalPath != null && File.Exists(tempOriginalPath))
                            File.Delete(tempOriginalPath);
                        if (tempModifiedPath != null && File.Exists(tempModifiedPath))
                            File.Delete(tempModifiedPath);
                    }
                    catch { }

                    ReleaseComObject(originalDoc);
                    ReleaseComObject(modifiedDoc);
                    ReleaseComObject(comparedDoc);
                    wordApp?.Quit();
                    ReleaseComObject(wordApp);
                }
            }

            private string ConvertPdfToDocx(Word.Application wordApp, string pdfPath)
            {
                // Create temp file path for converted docx
                string tempPath = Path.Combine(Path.GetTempPath(), $"PDFConvert_{Guid.NewGuid()}.docx");

                // Open PDF in Word (Word will convert it automatically)
                var doc = wordApp.Documents.Open(
                    pdfPath,
                    ConfirmConversions: false,
                    ReadOnly: true,
                    AddToRecentFiles: false
                );

                // Save as DOCX
                doc.SaveAs2(tempPath, Word.WdSaveFormat.wdFormatXMLDocument);
                doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);

                return tempPath;
            }

            private void AddSummaryTableAtEnd(Word.Document doc, ChangeSummary summary, ComparisonOptions options)
            {
                // Move to end of document and add page break
                var endRange = doc.Content;
                endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                endRange.InsertBreak(Word.WdBreakType.wdPageBreak);

                // Get range at end for new content
                endRange = doc.Content;
                endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                // Add centered title
                endRange.Text = "Summary Report\r\n";
                endRange.Font.Bold = 1;
                endRange.Font.Size = 12;
                endRange.Font.Color = Word.WdColor.wdColorBlack;
                endRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                // Move to end and add time
                endRange = doc.Content;
                endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                endRange.Text = $"{DateTime.Now:MM/dd/yyyy h:mm:ss tt}\r\n";
                endRange.Font.Bold = 0;
                endRange.Font.Size = 11;
                endRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                // Move to end again for table
                endRange = doc.Content;
                endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                endRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                endRange.InsertParagraphAfter();

                endRange = doc.Content;
                endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                // Add summary table - 13 rows (removed Style name row)
                var table = doc.Tables.Add(endRange, 13, 2);
                table.Borders.Enable = 1;
                table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                table.PreferredWidth = 70;

                // Set column widths - first column narrower, second column gets extra space
                table.Columns[1].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                table.Columns[1].PreferredWidth = 22;
                table.Columns[2].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                table.Columns[2].PreferredWidth = 78;

                // Row 1: Original Document
                table.Cell(1, 1).Range.Text = "Original Document:";
                table.Cell(1, 1).Range.Font.Bold = 1;
                table.Cell(1, 2).Range.Text = options.OriginalDisplayName ?? Path.GetFileName(options.OriginalPath);

                // Row 2: Modified Document
                table.Cell(2, 1).Range.Text = "Modified Document:";
                table.Cell(2, 1).Range.Font.Bold = 1;
                table.Cell(2, 2).Range.Text = options.ModifiedDisplayName ?? Path.GetFileName(options.ModifiedPath);

                // Row 3: Changes header
                table.Cell(3, 1).Range.Text = "Changes:";
                table.Cell(3, 1).Range.Font.Bold = 1;
                table.Cell(3, 2).Range.Text = "";

                // Row 4: Add (blue underline)
                table.Cell(4, 1).Range.Text = "Add";
                table.Cell(4, 1).Range.Font.Color = Word.WdColor.wdColorBlue;
                table.Cell(4, 1).Range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                table.Cell(4, 2).Range.Text = summary.Insertions.ToString();

                // Row 5: Delete (red strikethrough)
                table.Cell(5, 1).Range.Text = "Delete";
                table.Cell(5, 1).Range.Font.Color = Word.WdColor.wdColorRed;
                table.Cell(5, 1).Range.Font.StrikeThrough = 1;
                table.Cell(5, 2).Range.Text = summary.Deletions.ToString();

                // Row 6: Move From (green strikethrough)
                table.Cell(6, 1).Range.Text = "Move From";
                table.Cell(6, 1).Range.Font.Color = Word.WdColor.wdColorGreen;
                table.Cell(6, 1).Range.Font.StrikeThrough = 1;
                table.Cell(6, 2).Range.Text = summary.MoveFrom.ToString();

                // Row 7: Move To (green underline)
                table.Cell(7, 1).Range.Text = "Move To";
                table.Cell(7, 1).Range.Font.Color = Word.WdColor.wdColorGreen;
                table.Cell(7, 1).Range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                table.Cell(7, 2).Range.Text = summary.MoveTo.ToString();

                // Row 8: Table Insert (blue underline)
                table.Cell(8, 1).Range.Text = "Table Insert";
                table.Cell(8, 1).Range.Font.Color = Word.WdColor.wdColorBlue;
                table.Cell(8, 1).Range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                table.Cell(8, 2).Range.Text = summary.TableInserts.ToString();

                // Row 9: Table Delete (red strikethrough)
                table.Cell(9, 1).Range.Text = "Table Delete";
                table.Cell(9, 1).Range.Font.Color = Word.WdColor.wdColorRed;
                table.Cell(9, 1).Range.Font.StrikeThrough = 1;
                table.Cell(9, 2).Range.Text = summary.TableDeletes.ToString();

                // Row 10: Table moves to (green underline)
                table.Cell(10, 1).Range.Text = "Table moves to";
                table.Cell(10, 1).Range.Font.Color = Word.WdColor.wdColorGreen;
                table.Cell(10, 1).Range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                table.Cell(10, 2).Range.Text = "0";

                // Row 11: Table moves from (green strikethrough)
                table.Cell(11, 1).Range.Text = "Table moves from";
                table.Cell(11, 1).Range.Font.Color = Word.WdColor.wdColorGreen;
                table.Cell(11, 1).Range.Font.StrikeThrough = 1;
                table.Cell(11, 2).Range.Text = "0";

                // Row 12: Format changes
                table.Cell(12, 1).Range.Text = "Format changes";
                table.Cell(12, 1).Range.Font.Color = Word.WdColor.wdColorBlack;
                table.Cell(12, 2).Range.Text = summary.FormatChanges.ToString();

                // Row 13: Total Changes (bold)
                table.Cell(13, 1).Range.Text = "Total Changes:";
                table.Cell(13, 1).Range.Font.Bold = 1;
                table.Cell(13, 1).Range.Font.Color = Word.WdColor.wdColorBlack;
                table.Cell(13, 2).Range.Text = summary.TotalChanges.ToString();
                table.Cell(13, 2).Range.Font.Bold = 1;

                // Set column widths
                table.Columns[1].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                table.Columns[1].PreferredWidth = 70;
                table.Columns[2].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                table.Columns[2].PreferredWidth = 30;
            }

            #endregion

            #region PDF Report Generation

            private string GeneratePdfReport(string wordPath, ChangeSummary summary, ComparisonOptions options)
            {
                Word.Application wordApp = null;
                Word.Document doc = null;

                try
                {
                    var pdfPath = Path.ChangeExtension(wordPath, ".pdf");

                    wordApp = new Word.Application { Visible = false };
                    doc = wordApp.Documents.Open(wordPath, ReadOnly: false);

                    // Add summary table at the end for PDF only
                    AddSummaryTableAtEnd(doc, summary, options);

                    // Export to PDF with track changes visible
                    doc.ExportAsFixedFormat(
                        pdfPath,
                        Word.WdExportFormat.wdExportFormatPDF,
                        false,
                        Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                        Word.WdExportRange.wdExportAllDocument,
                        0, 0,
                        Word.WdExportItem.wdExportDocumentWithMarkup,
                        true,
                        true,
                        Word.WdExportCreateBookmarks.wdExportCreateHeadingBookmarks
                    );

                    return pdfPath;
                }
                catch (Exception ex)
                {
                    // PDF generation failed, but comparison succeeded
                    System.Diagnostics.Debug.WriteLine($"PDF generation failed: {ex.Message}");
                    return null;
                }
                finally
                {
                    doc?.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                    ReleaseComObject(doc);
                    wordApp?.Quit();
                    ReleaseComObject(wordApp);
                }
            }

            #endregion

            #region Helpers

            private static void ReleaseComObject(object obj)
            {
                if (obj != null)
                {
                    try
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                    }
                    catch
                    {
                        // Ignore
                    }
                }
            }

            #endregion
        }

        #endregion

        #region Data Structures

        private class DocumentContent
        {
            public List<ParagraphInfo> Paragraphs { get; set; } = new List<ParagraphInfo>();
            public List<TableInfo> Tables { get; set; } = new List<TableInfo>();
            public Dictionary<string, List<ParagraphInfo>> Headers { get; set; } = new Dictionary<string, List<ParagraphInfo>>();
            public Dictionary<string, List<ParagraphInfo>> Footers { get; set; } = new Dictionary<string, List<ParagraphInfo>>();
            public Dictionary<string, NumberingInfo> Numbering { get; set; } = new Dictionary<string, NumberingInfo>();
            public string StylesXml { get; set; }
        }

        private class ParagraphInfo
        {
            public string Text { get; set; } = "";
            public string StyleName { get; set; } = "";
            public int ListLevel { get; set; } = -1;
            public string ListNumber { get; set; } = "";
            public string NumId { get; set; }
            public bool IsNumbered { get; set; }
            public List<RunInfo> Runs { get; set; } = new List<RunInfo>();
        }

        private class RunInfo
        {
            public string Text { get; set; } = "";
            public bool IsBold { get; set; }
            public bool IsItalic { get; set; }
            public bool IsUnderline { get; set; }
            public bool IsStrikethrough { get; set; }
            public string FontName { get; set; } = "";
            public double FontSize { get; set; }
            public string Color { get; set; } = "";
        }

        private class TableInfo
        {
            public int RowCount { get; set; }
            public int ColumnCount { get; set; }
            public List<List<CellInfo>> Cells { get; set; } = new List<List<CellInfo>>();
        }

        private class CellInfo
        {
            public string Text { get; set; } = "";
            public int Row { get; set; }
            public int Column { get; set; }
            public int ColSpan { get; set; } = 1;
            public bool IsMerged { get; set; }
            public bool IsRowMergeStart { get; set; }
        }

        private class NumberingInfo
        {
            public string NumId { get; set; }
            public Dictionary<string, NumberingLevelInfo> Levels { get; set; } = new Dictionary<string, NumberingLevelInfo>();
        }

        private class NumberingLevelInfo
        {
            public int Level { get; set; }
            public string Format { get; set; }
            public string Text { get; set; }
        }

        private class Change
        {
            public ChangeType Type { get; set; }
            public string OriginalText { get; set; }
            public string ModifiedText { get; set; }
            public int ParagraphIndex { get; set; }
            public int TargetParagraphIndex { get; set; }
            public int WordStartIndex { get; set; }
            public string Location { get; set; }
            public ParagraphInfo OriginalParagraph { get; set; }
            public ParagraphInfo ModifiedParagraph { get; set; }
        }

        private class LCSMatch
        {
            public int OriginalIndex { get; set; }
            public int ModifiedIndex { get; set; }
            public ParagraphInfo OriginalPara { get; set; }
            public ParagraphInfo ModifiedPara { get; set; }
        }

        #endregion

        #region Changes Task Pane

        public class ChangesTaskPaneControl : UserControl
        {
            private DataGridView dgvChanges;
            private Label lblSummary;
            private Button btnAcceptAll;
            private Button btnRejectAll;

            private Word.Document document;
            private List<RevisionInfo> revisions;

            // === Keyboard capture window ===
            private DateTime _lastPaneInteractionUtc = DateTime.MinValue;
            private static readonly TimeSpan CaptureWindow = TimeSpan.FromSeconds(10);
            private bool _suppressSelectionChanged = false;

            // === Thread keyboard hook ===
            private IntPtr _kbdHook = IntPtr.Zero;
            private NativeMethods.HookProc _hookProc; // keep delegate alive

            public ChangesTaskPaneControl()
            {
                InitializeComponents();
            }

            protected override void OnHandleCreated(EventArgs e)
            {
                base.OnHandleCreated(e);
                InstallKeyboardHook();
            }

            protected override void OnHandleDestroyed(EventArgs e)
            {
                RemoveKeyboardHook();
                base.OnHandleDestroyed(e);
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing) RemoveKeyboardHook();
                base.Dispose(disposing);
            }

            private void MarkPaneInteracted()
            {
                _lastPaneInteractionUtc = DateTime.UtcNow;
            }

            private bool ShouldCaptureHotkeys()
            {
                if (IsDisposed || !Visible) return false;

                // Always capture if grid is focused.
                if (dgvChanges != null && dgvChanges.ContainsFocus) return true;

                // Otherwise capture briefly after the user interacts with the pane.
                return (DateTime.UtcNow - _lastPaneInteractionUtc) <= CaptureWindow;
            }

            private void InstallKeyboardHook()
            {
                if (_kbdHook != IntPtr.Zero) return;

                // Must be called on Word UI thread (task pane is created there in VSTO)
                _hookProc = KeyboardHookCallback;
                uint tid = NativeMethods.GetCurrentThreadId();

                _kbdHook = NativeMethods.SetWindowsHookEx(NativeMethods.WH_KEYBOARD, _hookProc, IntPtr.Zero, tid);

                // Optional debug:
                System.Diagnostics.Debug.WriteLine(_kbdHook != IntPtr.Zero
                    ? $"[DocComparison] Keyboard hook installed on TID={tid}"
                    : $"[DocComparison] Keyboard hook FAILED on TID={tid}, err={System.Runtime.InteropServices.Marshal.GetLastWin32Error()}");
            }

            private void RemoveKeyboardHook()
            {
                if (_kbdHook == IntPtr.Zero) return;
                try { NativeMethods.UnhookWindowsHookEx(_kbdHook); } catch { }
                _kbdHook = IntPtr.Zero;
                _hookProc = null;
            }

            private IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
            {
                if (nCode == NativeMethods.HC_ACTION)
                {
                    // WH_KEYBOARD: wParam = VK code, lParam bit 31 = key up
                    int vk = wParam.ToInt32();
                    bool isKeyUp = ((lParam.ToInt64() & 0x80000000) != 0);

                    if (!isKeyUp)
                    {
                        Keys key = (Keys)vk;
                        if (key == Keys.Up || key == Keys.Down || key == Keys.Enter || key == Keys.Delete || key == Keys.Back)
                        {
                            if (ShouldCaptureHotkeys())
                            {
                                try
                                {
                                    // Handle on UI thread
                                    BeginInvoke((Action)(() =>
                                    {
                                        try { HandleHotkey(key); } catch { }
                                    }));
                                }
                                catch { }

                                // Swallow so Word doesn't consume it.
                                return (IntPtr)1;
                            }
                        }
                    }
                }

                return NativeMethods.CallNextHookEx(_kbdHook, nCode, wParam, lParam);
            }

            private void InitializeComponents()
            {
                BackColor = Color.White;
                Dock = DockStyle.Fill;
                Font = new Font("Segoe UI", 9F);
                Padding = new Padding(10);

                btnAcceptAll = new Button
                {
                    Text = "Accept All",
                    Location = new Point(10, 10),
                    Size = new Size(100, 35),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.FromArgb(0, 120, 212),
                    ForeColor = Color.White,
                    Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
                };
                btnAcceptAll.FlatAppearance.BorderSize = 0;
                btnAcceptAll.Click += BtnAcceptAll_Click;
                btnAcceptAll.MouseDown += (s, e) => MarkPaneInteracted();
                Controls.Add(btnAcceptAll);

                btnRejectAll = new Button
                {
                    Text = "Reject All",
                    Location = new Point(115, 10),
                    Size = new Size(100, 35),
                    FlatStyle = FlatStyle.Flat,
                    Anchor = AnchorStyles.Top | AnchorStyles.Right
                };
                btnRejectAll.Click += BtnRejectAll_Click;
                btnRejectAll.MouseDown += (s, e) => MarkPaneInteracted();
                Controls.Add(btnRejectAll);

                lblSummary = new Label
                {
                    Text = "Loading changes...",
                    Location = new Point(10, 55),
                    Size = new Size(430, 25),
                    Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                    ForeColor = Color.Gray
                };
                Controls.Add(lblSummary);

                dgvChanges = new DataGridView
                {
                    Location = new Point(10, 87),
                    Size = new Size(430, 400),
                    Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                    AllowUserToAddRows = false,
                    AllowUserToDeleteRows = false,
                    ReadOnly = false,
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                    MultiSelect = false,
                    RowHeadersVisible = false,
                    BackgroundColor = Color.White,
                    BorderStyle = BorderStyle.FixedSingle,
                    CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                    GridColor = Color.LightGray,
                    ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing,
                    ColumnHeadersHeight = 36,
                    RowTemplate = { Height = 36 }
                };

                dgvChanges.Columns.Add(new DataGridViewTextBoxColumn
                {
                    Name = "No",
                    HeaderText = "#",
                    Width = 40,
                    ReadOnly = true,
                    SortMode = DataGridViewColumnSortMode.NotSortable,
                    DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter }
                });

                dgvChanges.Columns.Add(new DataGridViewTextBoxColumn
                {
                    Name = "Item",
                    HeaderText = "Item",
                    Width = 125,
                    ReadOnly = true,
                    SortMode = DataGridViewColumnSortMode.NotSortable
                });

                dgvChanges.Columns.Add(new DataGridViewButtonColumn
                {
                    Name = "Accept",
                    HeaderText = "Accept",
                    Text = "✓",
                    UseColumnTextForButtonValue = true,
                    Width = 80
                });

                dgvChanges.Columns.Add(new DataGridViewButtonColumn
                {
                    Name = "Reject",
                    HeaderText = "Reject",
                    Text = "✗",
                    UseColumnTextForButtonValue = true,
                    Width = 80
                });

                dgvChanges.Columns.Add(new DataGridViewTextBoxColumn
                {
                    Name = "ChangeString",
                    HeaderText = "Change",
                    AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                    ReadOnly = true,
                    SortMode = DataGridViewColumnSortMode.NotSortable
                });

                dgvChanges.CellClick += DgvChanges_CellClick;
                dgvChanges.SelectionChanged += DgvChanges_SelectionChanged;

                dgvChanges.MouseDown += (s, e) => MarkPaneInteracted();
                dgvChanges.Enter += (s, e) => MarkPaneInteracted();
                dgvChanges.GotFocus += (s, e) => MarkPaneInteracted();

                Controls.Add(dgvChanges);

                Resize += OnResize;
                Load += (s, e) => OnResize(s, e);
                MouseDown += (s, e) => MarkPaneInteracted();
            }

            private void OnResize(object sender, EventArgs e)
            {
                int availableWidth = Width - 25;
                int buttonWidth = Math.Max(80, availableWidth / 2);

                btnAcceptAll.Width = buttonWidth;
                btnAcceptAll.Location = new Point(10, 10);

                btnRejectAll.Width = buttonWidth;
                btnRejectAll.Location = new Point(10 + buttonWidth + 5, 10);

                lblSummary.Width = Width - 20;
                dgvChanges.Size = new Size(Width - 20, Height - 97);
            }

            public void LoadDocument(Word.Document doc)
            {
                document = doc;

                // Ensure hook is installed (in case handle timing is odd)
                InstallKeyboardHook();

                var refreshTimer = new System.Windows.Forms.Timer();
                refreshTimer.Interval = 1000;
                refreshTimer.Tick += (s, e) =>
                {
                    try
                    {
                        if (document != null)
                        {
                            int currentRevisionCount = document.Revisions.Count;
                            if (currentRevisionCount != dgvChanges.Rows.Count)
                                RefreshChanges();
                        }
                    }
                    catch { refreshTimer.Stop(); }
                };
                refreshTimer.Start();
                Disposed += (s, e) => refreshTimer.Stop();

                RefreshChanges();

                if (dgvChanges.Rows.Count > 0)
                {
                    SelectRow(0);
                    dgvChanges.Focus();
                    MarkPaneInteracted();
                }
            }

            public void RefreshChanges()
            {
                if (document == null) return;

                try
                {
                    revisions = new List<RevisionInfo>();
                    dgvChanges.Rows.Clear();

                    int index = 1;
                    foreach (Word.Revision rev in document.Revisions)
                    {
                        var revInfo = new RevisionInfo
                        {
                            Index = index,
                            OriginalIndex = index,
                            Revision = rev,
                            Type = GetRevisionType(rev.Type),
                            Text = TruncateText(rev.Range?.Text ?? "", 100)
                        };
                        revisions.Add(revInfo);

                        int rowIndex = dgvChanges.Rows.Add();
                        var row = dgvChanges.Rows[rowIndex];
                        row.Cells["No"].Value = index;
                        row.Cells["Item"].Value = revInfo.Type;
                        row.Cells["ChangeString"].Value = revInfo.Text;
                        row.Tag = revInfo;

                        switch (rev.Type)
                        {
                            case Word.WdRevisionType.wdRevisionInsert:
                                row.Cells["Item"].Style.ForeColor = Color.Blue;
                                break;
                            case Word.WdRevisionType.wdRevisionDelete:
                                row.Cells["Item"].Style.ForeColor = Color.Red;
                                break;
                            case Word.WdRevisionType.wdRevisionMovedFrom:
                            case Word.WdRevisionType.wdRevisionMovedTo:
                                row.Cells["Item"].Style.ForeColor = Color.Green;
                                break;
                        }

                        index++;
                    }

                    RefreshSummary();
                }
                catch (Exception ex)
                {
                    lblSummary.Text = $"Error loading changes: {ex.Message}";
                }
            }

            private void NavigateToRevision(Word.Revision revision)
            {
                if (document == null || revision == null) return;

                try
                {
                    // Make sure the correct doc/window is active; otherwise Word may not scroll.
                    document.Activate();

                    var app = document.Application;
                    var win = app?.ActiveWindow;
                    if (win == null) return;

                    // Duplicate the range to avoid COM weirdness if Word mutates it during select.
                    Word.Range rng = revision.Range?.Duplicate;
                    if (rng == null) return;

                    // Scroll first (reduces “blink”), then select to force caret/page navigation,
                    // then scroll again to ensure it’s actually in view.
                    win.ScrollIntoView(rng, true);
                    rng.Select();
                    win.ScrollIntoView(rng, true);
                }
                catch
                {
                    // Ignore navigation failures (revision might be stale/removed)
                }
            }

            private void DgvChanges_SelectionChanged(object sender, EventArgs e)
            {
                if (_suppressSelectionChanged) return;
                if (dgvChanges.SelectedRows.Count == 0) return;

                MarkPaneInteracted();

                var row = dgvChanges.SelectedRows[0];
                var revInfo = row.Tag as RevisionInfo;
                if (revInfo?.Revision == null) return;

                NavigateToRevision(revInfo.Revision);
            }

            private void DgvChanges_CellClick(object sender, DataGridViewCellEventArgs e)
            {
                if (e.RowIndex < 0) return;

                MarkPaneInteracted();

                var row = dgvChanges.Rows[e.RowIndex];
                var revInfo = row.Tag as RevisionInfo;
                if (revInfo?.Revision == null) return;

                try
                {
                    if (e.ColumnIndex == dgvChanges.Columns["Accept"].Index)
                    {
                        revInfo.Revision.Accept();
                        dgvChanges.Rows.RemoveAt(e.RowIndex);
                        RefreshSummary();
                        ReselectAfterRemove(e.RowIndex);
                    }
                    else if (e.ColumnIndex == dgvChanges.Columns["Reject"].Index)
                    {
                        revInfo.Revision.Reject();
                        dgvChanges.Rows.RemoveAt(e.RowIndex);
                        RefreshSummary();
                        ReselectAfterRemove(e.RowIndex);
                    }
                }
                catch { RefreshChanges(); }
            }

            private void ReselectAfterRemove(int removedIndex)
            {
                if (dgvChanges.Rows.Count == 0) return;
                int next = Math.Min(removedIndex, dgvChanges.Rows.Count - 1);
                SelectRow(next);
            }

            private void SelectRow(int rowIndex)
            {
                if (rowIndex < 0 || rowIndex >= dgvChanges.Rows.Count) return;

                _suppressSelectionChanged = true;
                try
                {
                    dgvChanges.ClearSelection();
                    dgvChanges.Rows[rowIndex].Selected = true;
                    dgvChanges.CurrentCell = dgvChanges.Rows[rowIndex].Cells[0];
                }
                finally
                {
                    _suppressSelectionChanged = false;
                }

                // Behave like SelectionChanged
                try
                {
                    var revInfo = dgvChanges.Rows[rowIndex].Tag as RevisionInfo;
                    if (revInfo?.Revision != null)
                    {
                        MarkPaneInteracted();
                        NavigateToRevision(revInfo.Revision);
                    }
                }
                catch { }
            }

            private void HandleHotkey(Keys key)
            {
                if (dgvChanges == null || dgvChanges.Rows.Count == 0) return;

                int current =
                    dgvChanges.CurrentCell?.RowIndex
                    ?? (dgvChanges.SelectedRows.Count > 0 ? dgvChanges.SelectedRows[0].Index : 0);

                if (key == Keys.Down)
                {
                    SelectRow(Math.Min(current + 1, dgvChanges.Rows.Count - 1));
                    return;
                }

                if (key == Keys.Up)
                {
                    SelectRow(Math.Max(current - 1, 0));
                    return;
                }

                if (key == Keys.Enter)
                {
                    AcceptSelected();
                    return;
                }

                if (key == Keys.Delete || key == Keys.Back)
                {
                    RejectSelected();
                    return;
                }
            }

            private void AcceptSelected()
            {
                if (dgvChanges.SelectedRows.Count == 0) return;

                var row = dgvChanges.SelectedRows[0];
                var revInfo = row.Tag as RevisionInfo;
                if (revInfo?.Revision == null) return;

                try
                {
                    int idx = row.Index;
                    revInfo.Revision.Accept();
                    dgvChanges.Rows.RemoveAt(idx);
                    RefreshSummary();
                    ReselectAfterRemove(idx);
                }
                catch { RefreshChanges(); }
            }

            private void RejectSelected()
            {
                if (dgvChanges.SelectedRows.Count == 0) return;

                var row = dgvChanges.SelectedRows[0];
                var revInfo = row.Tag as RevisionInfo;
                if (revInfo?.Revision == null) return;

                try
                {
                    int idx = row.Index;
                    revInfo.Revision.Reject();
                    dgvChanges.Rows.RemoveAt(idx);
                    RefreshSummary();
                    ReselectAfterRemove(idx);
                }
                catch { RefreshChanges(); }
            }

            private void RefreshSummary()
            {
                int count = dgvChanges.Rows.Count;
                int adds = 0, deletes = 0, moves = 0;

                foreach (DataGridViewRow row in dgvChanges.Rows)
                {
                    var type = row.Cells["Item"].Value?.ToString();
                    if (type == "Add") adds++;
                    else if (type == "Delete") deletes++;
                    else if (type == "Move From" || type == "Move To") moves++;
                }

                lblSummary.Text = $"{count} changes: {adds} adds, {deletes} deletes, {moves} moves";
            }

            private void BtnAcceptAll_Click(object sender, EventArgs e)
            {
                if (document == null) return;
                MarkPaneInteracted();

                try
                {
                    document.Revisions.AcceptAll();
                    RefreshChanges();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error accepting changes: {ex.Message}",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            private void BtnRejectAll_Click(object sender, EventArgs e)
            {
                if (document == null) return;
                MarkPaneInteracted();

                try
                {
                    document.Revisions.RejectAll();
                    RefreshChanges();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error rejecting changes: {ex.Message}",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            private string GetRevisionType(Word.WdRevisionType type)
            {
                switch (type)
                {
                    case Word.WdRevisionType.wdRevisionInsert: return "Add";
                    case Word.WdRevisionType.wdRevisionDelete: return "Delete";
                    case Word.WdRevisionType.wdRevisionMovedFrom: return "Move From";
                    case Word.WdRevisionType.wdRevisionMovedTo: return "Move To";
                    case Word.WdRevisionType.wdRevisionProperty: return "Format";
                    default: return "Other";
                }
            }

            private string TruncateText(string text, int maxLength)
            {
                if (string.IsNullOrEmpty(text)) return "";
                text = text.Replace("\r", " ").Replace("\n", " ").Replace("\a", "").Trim();
                if (text.Length <= maxLength) return text;
                return text.Substring(0, maxLength - 3) + "...";
            }

            private class RevisionInfo
            {
                public int Index { get; set; }
                public int OriginalIndex { get; set; }
                public Word.Revision Revision { get; set; }
                public string Type { get; set; }
                public string Text { get; set; }
            }

            private static class NativeMethods
            {
                public const int WH_KEYBOARD = 2;
                public const int HC_ACTION = 0;

                public delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

                [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
                public static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hmod, uint dwThreadId);

                [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
                public static extern bool UnhookWindowsHookEx(IntPtr hhk);

                [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
                public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

                [System.Runtime.InteropServices.DllImport("kernel32.dll")]
                public static extern uint GetCurrentThreadId();
            }
        }

        /// <summary>
        /// Shows the Changes task pane for the specified document
        /// </summary>
        public static void ShowChangesTaskPane(Word.Document doc)
        {
            try
            {
                var control = new ChangesTaskPaneControl();
                var taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(control, "Tracked Changes");
                taskPane.Width = 550;
                taskPane.Visible = true;
                control.LoadDocument(doc);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error showing task pane: {ex.Message}");
            }
        }

        #endregion
    }
}