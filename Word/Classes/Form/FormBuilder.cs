using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using OfficePragmatics_Excel.Settings;

namespace Office_Pragmatics_Word
{
    public static class FormEditor
    {
        private static CustomTaskPane _taskPane;
        private static FormBuilderTaskPaneHost _taskPaneHost;

        public static void ShowFormEditor()
        {
            try
            {
                Word.Application wordApp = Globals.ThisAddIn.Application;
                if (wordApp == null)
                {
                    MessageBox.Show("Word application not available.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var customTaskPanes = Globals.ThisAddIn.CustomTaskPanes;
                if (customTaskPanes == null)
                {
                    MessageBox.Show("CustomTaskPanes not available.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (_taskPane != null)
                {
                    _taskPane.Visible = !_taskPane.Visible;
                    if (_taskPane.Visible)
                    {
                        _taskPaneHost?.RefreshContent(wordApp);
                        try { _taskPane.Width = CalculateInitialPaneWidth(wordApp); } catch { }
                    }
                    return;
                }

                _taskPaneHost = new FormBuilderTaskPaneHost(wordApp);
                _taskPane = customTaskPanes.Add(_taskPaneHost, "Form Builder");

                _taskPane.Width = CalculateInitialPaneWidth(wordApp);
                _taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                _taskPane.Visible = true;
                _taskPane.VisibleChanged += TaskPane_VisibleChanged;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error showing Form Editor: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static int CalculateInitialPaneWidth(Word.Application wordApp)
        {
            try
            {
                int hwnd = 0;
                try { hwnd = wordApp?.ActiveWindow?.Hwnd ?? 0; } catch { hwnd = 0; }

                Screen screen = (hwnd != 0) ? Screen.FromHandle(new IntPtr(hwnd)) : Screen.PrimaryScreen;
                int screenW = screen.WorkingArea.Width;

                float scale = 1f;
                if (hwnd != 0)
                {
                    using (Graphics g = Graphics.FromHwnd(new IntPtr(hwnd)))
                        scale = g.DpiX / 96f;
                }

                int min = (int)Math.Round(300 * scale);
                int max = (int)Math.Round(520 * scale);

                int desired = (int)Math.Round(screenW * 0.30);

                int hardMax = (int)Math.Round(screenW * 0.42);
                max = Math.Min(max, hardMax);

                if (desired < min) desired = min;
                if (desired > max) desired = max;

                return desired;
            }
            catch
            {
                return Theme.Px(360);
            }
        }

        private static void TaskPane_VisibleChanged(object sender, EventArgs e)
        {
            try
            {
                if (_taskPane != null && _taskPane.Visible)
                {
                    Word.Application wordApp = Globals.ThisAddIn.Application;
                    _taskPaneHost?.RefreshContent(wordApp);
                }
            }
            catch { }
        }
    }

    public class FormBuilderTaskPaneHost : UserControl
    {
        private Word.Application _wordApp;
        private FormBuilderTaskPane _innerPane;

        public FormBuilderTaskPaneHost(Word.Application wordApp)
        {
            Dpi.ResetToDesign();

            AutoScaleMode = AutoScaleMode.Font;
            _wordApp = wordApp;
            BackColor = Theme.Surface;
            Dock = DockStyle.Fill;
            Margin = new Padding(0);

            _innerPane = new FormBuilderTaskPane(_wordApp) { Dock = DockStyle.Fill };
            Controls.Add(_innerPane);
        }

        public void RefreshContent(Word.Application wordApp)
        {
            _wordApp = wordApp;
            _innerPane?.RefreshFields();
        }
    }

    public class FormBuilderTaskPane : UserControl
    {
        private readonly Word.Application _wordApp;
        private TabControl tabControl;
        private TabPage tabEdit;
        private TabPage tabFill;

        // Edit tab controls
        private TextBox txtFieldLabel;
        private ComboBox cboFieldType;
        private TextBox txtChoices;
        private Panel pnlChoices;
        private Button btnAddField;
        private Button btnRemoveField;
        private ComboBox cboExistingField;
        private Button btnLinkToField;
        private ListBox lstFields;
        private readonly Dictionary<string, string> existingFieldTags = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // Fill tab controls
        private Panel pnlQuestions;
        private Button btnFillDocument;
        private Button btnClearForm;
        private Dictionary<string, Control> fieldControls;
        private bool _isRefreshing = false;

        // Layout roots
        private TableLayoutPanel _editRoot;
        private TableLayoutPanel _fillRoot;

        public FormBuilderTaskPane(Word.Application wordApp)
        {
            AutoScaleMode = AutoScaleMode.Font;
            Font = new Font("Segoe UI", 9f);

            _wordApp = wordApp;
            InitializeComponents();
            LoadExistingFields();
            if (lstFields.Items.Count > 0)
            {
                tabControl.SelectedTab = tabFill;
                RefreshQuestionPanel();
            }
        }

        public void RefreshFields()
        {
            LoadExistingFields();
            if (tabControl.SelectedTab == tabFill)
                RefreshQuestionPanel();
        }

        private void InitializeComponents()
        {
            BackColor = Theme.Surface;
            Dock = DockStyle.Fill;
            Margin = new Padding(0);
            Theme.DoubleBuffer(this);

            tabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = Font,
                Margin = new Padding(0),
                Padding = new Point(Math.Max(6, Theme.PadMd), Math.Max(4, Theme.PadSm))
            };

            tabEdit = new TabPage("Edit") { BackColor = Theme.Surface, Padding = new Padding(0) };
            tabFill = new TabPage("Fill") { BackColor = Theme.Surface, Padding = new Padding(0) };

            InitializeEditTab();
            InitializeFillTab();

            tabControl.TabPages.Add(tabEdit);
            tabControl.TabPages.Add(tabFill);
            tabControl.SelectedIndexChanged += (s, e) =>
            {
                if (tabControl.SelectedTab == tabFill) RefreshQuestionPanel();
            };

            Controls.Add(tabControl);

            SizeChanged += (s, e) => ApplyResponsiveMetrics();
            ApplyResponsiveMetrics();
        }

        private void ApplyResponsiveMetrics()
        {
            int w = ClientSize.Width;
            int pad = Theme.PadMd;

            if (w > 0)
            {
                if (w <= Theme.Px(320)) pad = Theme.PadSm;
                else if (w >= Theme.Px(520)) pad = Theme.PadLg;
            }

            if (_editRoot != null) _editRoot.Padding = new Padding(pad);
            if (_fillRoot != null) _fillRoot.Padding = new Padding(pad);
        }

        private void InitializeEditTab()
        {
            _editRoot = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                BackColor = Theme.Surface,
                Padding = new Padding(Theme.PadMd),
                Margin = new Padding(0)
            };
            _editRoot.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            _editRoot.RowStyles.Add(new RowStyle(SizeType.AutoSize));     // Add
            _editRoot.RowStyles.Add(new RowStyle(SizeType.AutoSize));     // Link
            _editRoot.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // List (fill)

            // ===== ADD NEW FIELD =====
            var addCard = CreateCard("Add New Field", fill: false, out TableLayoutPanel addBody);
            addCard.Margin = new Padding(0, 0, 0, Theme.PadMd);

            AddAutoRow(addBody, CreateLabel("Question / Label"));

            txtFieldLabel = CreateTextBox("e.g., What is the company name?");
            txtFieldLabel.Multiline = true;
            txtFieldLabel.AcceptsReturn = true;
            txtFieldLabel.WordWrap = true;
            txtFieldLabel.ScrollBars = ScrollBars.Vertical;

            int min3Lines = CalcMinMultilineHeight(txtFieldLabel.Font, 3);
            txtFieldLabel.MinimumSize = new Size(0, min3Lines);
            txtFieldLabel.Height = min3Lines;
            txtFieldLabel.Margin = new Padding(0, 0, 0, Theme.PadSm);
            AddAutoRow(addBody, txtFieldLabel);

            AddAutoRow(addBody, CreateLabel("Field Type"));

            cboFieldType = CreateComboBox();
            cboFieldType.Items.AddRange(new object[] { "Text", "Date", "Yes/No", "Multiple Choice" });
            cboFieldType.SelectedIndex = 0;
            cboFieldType.SelectedIndexChanged += CboFieldType_SelectedIndexChanged;
            cboFieldType.Margin = new Padding(0, 0, 0, Theme.PadSm);
            AddAutoRow(addBody, cboFieldType);

            pnlChoices = new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Visible = false,
                Margin = new Padding(0, 0, 0, Theme.PadSm)
            };

            var choicesBody = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 1,
                RowCount = 0,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = new Padding(0)
            };
            choicesBody.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            AddAutoRow(choicesBody, CreateLabel("Choices (comma-separated)"));

            txtChoices = CreateTextBox("e.g., Option A, Option B, Option C");
            txtChoices.Margin = new Padding(0);
            AddAutoRow(choicesBody, txtChoices);

            pnlChoices.Controls.Add(choicesBody);
            AddAutoRow(addBody, pnlChoices);

            // Add/Remove row
            var addRemoveRow = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                RowCount = 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = new Padding(0, Theme.PadSm, 0, 0)
            };
            addRemoveRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            addRemoveRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

            btnAddField = CreatePrimaryButton("Add Field");
            btnAddField.Dock = DockStyle.Fill;
            btnAddField.Click += BtnAddField_Click;
            btnAddField.Margin = new Padding(0, 0, Theme.PadSm / 2, 0);

            btnRemoveField = CreateSecondaryButton("Remove Field");
            btnRemoveField.Dock = DockStyle.Fill;
            btnRemoveField.Click += BtnRemoveField_Click;
            btnRemoveField.Margin = new Padding(Theme.PadSm / 2, 0, 0, 0);

            addRemoveRow.Controls.Add(btnAddField, 0, 0);
            addRemoveRow.Controls.Add(btnRemoveField, 1, 0);

            AddAutoRow(addBody, addRemoveRow);

            _editRoot.Controls.Add(addCard, 0, 0);

            // ===== LINK TO EXISTING FIELD =====
            var linkCard = CreateCard("Link to Existing Field", fill: false, out TableLayoutPanel linkBody);
            linkCard.Margin = new Padding(0, 0, 0, Theme.PadMd);

            AddAutoRow(linkBody, CreateLabel("Select Field"));

            cboExistingField = CreateComboBox();
            cboExistingField.Margin = new Padding(0, 0, 0, Theme.PadSm);
            AddAutoRow(linkBody, cboExistingField);

            btnLinkToField = CreateSecondaryButton("Link Selected Text");
            btnLinkToField.Dock = DockStyle.Top;
            btnLinkToField.Click += BtnLinkToField_Click;
            AddAutoRow(linkBody, btnLinkToField);

            _editRoot.Controls.Add(linkCard, 0, 1);

            // ===== DOCUMENT FIELDS (FILL) =====
            var listCard = CreateCard("Document Fields", fill: true, out TableLayoutPanel listBody);
            listCard.Margin = new Padding(0);

            var listLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Margin = new Padding(0)
            };
            listLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            listLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            listLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            lstFields = new ListBox
            {
                Dock = DockStyle.Fill,
                Font = Font,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White,
                IntegralHeight = false,
                Margin = new Padding(0)
            };
            lstFields.SelectedIndexChanged += LstFields_SelectedIndexChanged;
            listLayout.Controls.Add(lstFields, 0, 0);

            var btnRefresh = CreateSecondaryButton("↻ Refresh");
            btnRefresh.Dock = DockStyle.Fill; // full-width stretch
            btnRefresh.Margin = new Padding(0, Theme.PadSm, 0, 0);
            btnRefresh.Click += (s, e) => LoadExistingFields();
            listLayout.Controls.Add(btnRefresh, 0, 1);

            // listBody is a TableLayoutPanel; make it host a single fill control
            listBody.Dock = DockStyle.Fill;
            listBody.ColumnStyles.Clear();
            listBody.RowStyles.Clear();
            listBody.ColumnCount = 1;
            listBody.RowCount = 1;
            listBody.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            listBody.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            listBody.Controls.Add(listLayout, 0, 0);

            _editRoot.Controls.Add(listCard, 0, 2);

            tabEdit.Controls.Add(_editRoot);
        }

        private void InitializeFillTab()
        {
            _fillRoot = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                BackColor = Theme.Surface,
                Padding = new Padding(Theme.PadMd),
                Margin = new Padding(0)
            };
            _fillRoot.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            _fillRoot.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            _fillRoot.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            pnlQuestions = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0)
            };
            Theme.DoubleBuffer(pnlQuestions);
            _fillRoot.Controls.Add(pnlQuestions, 0, 0);

            var buttons = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                RowCount = 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = new Padding(0, Theme.PadMd, 0, 0)
            };
            buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

            btnFillDocument = CreatePrimaryButton("Fill Document");
            btnFillDocument.Dock = DockStyle.Fill;
            btnFillDocument.Click += BtnFillDocument_Click;
            btnFillDocument.Margin = new Padding(0, 0, Theme.PadSm / 2, 0);

            btnClearForm = CreateSecondaryButton("Clear Form");
            btnClearForm.Dock = DockStyle.Fill;
            btnClearForm.Click += (s, e) => RefreshQuestionPanel();
            btnClearForm.Margin = new Padding(Theme.PadSm / 2, 0, 0, 0);

            buttons.Controls.Add(btnFillDocument, 0, 0);
            buttons.Controls.Add(btnClearForm, 1, 0);

            _fillRoot.Controls.Add(buttons, 0, 1);

            fieldControls = new Dictionary<string, Control>();
            tabFill.Controls.Add(_fillRoot);
        }

        // ===== UI HELPERS =====

        private Panel CreateCard(string title, bool fill, out TableLayoutPanel body)
        {
            var card = new Panel
            {
                Dock = fill ? DockStyle.Fill : DockStyle.Top,
                BackColor = Color.FromArgb(250, 250, 250),
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(Theme.PadSm),
                Margin = new Padding(0),
                AutoSize = !fill,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };

            var outer = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Margin = new Padding(0),
                AutoSize = !fill,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            outer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            outer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            outer.RowStyles.Add(new RowStyle(fill ? SizeType.Percent : SizeType.AutoSize, fill ? 100 : 0));

            var titleLabel = new Label
            {
                Text = title,
                Font = new Font(Font.FontFamily, Font.Size, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 102, 178),
                Dock = DockStyle.Top,
                AutoSize = true,
                Margin = new Padding(0, 0, 0, Theme.PadSm)
            };

            body = new TableLayoutPanel
            {
                Dock = fill ? DockStyle.Fill : DockStyle.Top,
                ColumnCount = 1,
                RowCount = 0,
                Margin = new Padding(0),
                AutoSize = !fill,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            outer.Controls.Add(titleLabel, 0, 0);
            outer.Controls.Add(body, 0, 1);
            card.Controls.Add(outer);

            return card;
        }

        private static void AddAutoRow(TableLayoutPanel tlp, Control control)
        {
            int row = tlp.RowStyles.Count;
            tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tlp.RowCount = tlp.RowStyles.Count;

            if (control.Dock == DockStyle.None)
                control.Dock = DockStyle.Top;

            tlp.Controls.Add(control, 0, row);
        }

        private static int CalcMinMultilineHeight(Font font, int lines)
        {
            int lineH = TextRenderer.MeasureText("Wg", font).Height;
            return (int)Math.Ceiling((lineH * lines) + (lineH * 0.75));
        }

        private Label CreateLabel(string text)
        {
            return new Label
            {
                Text = text,
                Font = new Font(Font.FontFamily, Math.Max(8f, Font.Size - 0.5f)),
                ForeColor = Color.FromArgb(80, 80, 80),
                Dock = DockStyle.Top,
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 2)
            };
        }

        private TextBox CreateTextBox(string placeholder)
        {
            var txt = new TextBox
            {
                Dock = DockStyle.Top,
                Font = Font,
                ForeColor = Color.Gray,
                Text = placeholder,
                BorderStyle = BorderStyle.FixedSingle
            };

            string ph = placeholder;
            txt.GotFocus += (s, e) =>
            {
                if (txt.ForeColor == Color.Gray)
                {
                    txt.Text = "";
                    txt.ForeColor = Color.Black;
                }
            };
            txt.LostFocus += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(txt.Text))
                {
                    txt.ForeColor = Color.Gray;
                    txt.Text = ph;
                }
            };

            return txt;
        }

        private ComboBox CreateComboBox()
        {
            return new ComboBox
            {
                Dock = DockStyle.Top,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = Font,
                IntegralHeight = false
            };
        }

        private Button CreatePrimaryButton(string text)
        {
            var btn = new Button
            {
                Text = text,
                FlatStyle = FlatStyle.Flat,
                Font = new Font(Font.FontFamily, Font.Size, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 102, 178),
                ForeColor = Color.White,
                Cursor = Cursors.Hand,
                Padding = new Padding(Theme.PadMd, Theme.PadSm, Theme.PadMd, Theme.PadSm),
                UseCompatibleTextRendering = true,
                AutoSize = false
            };
            btn.FlatAppearance.BorderSize = 0;

            int minH = TextRenderer.MeasureText("Wg", btn.Font).Height + (Theme.PadSm * 2);
            btn.MinimumSize = new Size(0, minH);

            return btn;
        }

        private Button CreateSecondaryButton(string text)
        {
            var btn = new Button
            {
                Text = text,
                FlatStyle = FlatStyle.Flat,
                Font = Font,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(80, 80, 80),
                Cursor = Cursors.Hand,
                Padding = new Padding(Theme.PadMd, Theme.PadSm, Theme.PadMd, Theme.PadSm),
                UseCompatibleTextRendering = true,
                AutoSize = false
            };
            btn.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
            btn.FlatAppearance.BorderSize = 1;

            int minH = TextRenderer.MeasureText("Wg", btn.Font).Height + (Theme.PadSm * 2);
            btn.MinimumSize = new Size(0, minH);

            return btn;
        }

        // ===== EVENT HANDLERS (UNCHANGED) =====

        private void CboFieldType_SelectedIndexChanged(object sender, EventArgs e)
        {
            pnlChoices.Visible = cboFieldType.SelectedItem?.ToString() == "Multiple Choice";
        }

        private void BtnAddField_Click(object sender, EventArgs e)
        {
            try
            {
                if (_wordApp?.Selection == null)
                    throw new InvalidOperationException("Word selection not available.");

                string fieldLabel = GetTextBoxValue(txtFieldLabel, "e.g., What is the company name?");
                if (string.IsNullOrWhiteSpace(fieldLabel))
                {
                    MessageBox.Show("Please enter a field label/question.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string fieldName = new string(fieldLabel.Where(c => char.IsLetterOrDigit(c) || c == '_').ToArray());
                if (fieldName.Length > 50) fieldName = fieldName.Substring(0, 50);

                Word.Selection selection = _wordApp.Selection;
                if (selection.Range == null || selection.Range.Text.Trim().Length == 0)
                {
                    MessageBox.Show("Please select some text in the document to convert to a field.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string fieldType = cboFieldType?.Text ?? "Text";
                string choices = "";
                if (fieldType == "Multiple Choice")
                {
                    choices = GetTextBoxValue(txtChoices, "e.g., Option A, Option B, Option C");
                    if (string.IsNullOrWhiteSpace(choices))
                    {
                        MessageBox.Show("Please enter choices for the multiple choice field.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }

                Word.ContentControl cc = selection.Range.ContentControls.Add(Word.WdContentControlType.wdContentControlText);
                cc.Title = fieldName;
                cc.Tag = string.IsNullOrEmpty(choices)
                  ? $"{fieldName}|{fieldType}|{fieldLabel}"
                  : $"{fieldName}|{fieldType}|{fieldLabel}|{choices}";
                cc.Appearance = Word.WdContentControlAppearance.wdContentControlBoundingBox;
                cc.LockContents = false;
                cc.SetPlaceholderText(Text: $"[{fieldLabel}]");

                MessageBox.Show("Field added successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ResetAddFieldForm();
                LoadExistingFields();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adding field: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ResetAddFieldForm()
        {
            txtFieldLabel.ForeColor = Color.Gray;
            txtFieldLabel.Text = "e.g., What is the company name?";
            txtChoices.ForeColor = Color.Gray;
            txtChoices.Text = "e.g., Option A, Option B, Option C";
            cboFieldType.SelectedIndex = 0;
            pnlChoices.Visible = false;
        }

        private void BtnRemoveField_Click(object sender, EventArgs e)
        {
            try
            {
                if (lstFields.SelectedItem == null)
                {
                    MessageBox.Show("Please select a field to remove.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string fieldName = lstFields.SelectedItem.ToString().Split('•')[0].Trim();

                foreach (Word.ContentControl cc in _wordApp.ActiveDocument.ContentControls)
                {
                    if (cc.Title == fieldName)
                    {
                        string text = cc.Range.Text;
                        Word.Range range = cc.Range;
                        cc.Delete();
                        range.Text = text;
                        break;
                    }
                }

                MessageBox.Show($"Field '{fieldName}' removed.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadExistingFields();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error removing field: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnLinkToField_Click(object sender, EventArgs e)
        {
            try
            {
                if (cboExistingField.SelectedItem == null)
                {
                    MessageBox.Show("Select an existing field to link to.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (_wordApp?.Selection?.Range == null || _wordApp.Selection.Range.Text.Trim().Length == 0)
                {
                    MessageBox.Show("Select some text in the document to link.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string fieldName = cboExistingField.SelectedItem.ToString();
                if (!existingFieldTags.TryGetValue(fieldName, out string tagToReuse))
                {
                    MessageBox.Show("Could not find the selected field's metadata.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                Word.ContentControl cc = _wordApp.Selection.Range.ContentControls.Add(Word.WdContentControlType.wdContentControlText);
                cc.Title = fieldName;
                cc.Tag = tagToReuse;
                cc.Appearance = Word.WdContentControlAppearance.wdContentControlBoundingBox;
                cc.LockContents = false;

                var parts = tagToReuse.Split('|');
                if (parts.Length >= 3)
                    cc.SetPlaceholderText(Text: $"[{parts[2]}]");

                MessageBox.Show("Selection linked to existing field.", "Linked", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error linking to field: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LstFields_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (lstFields.SelectedItem == null) return;
                string fieldName = lstFields.SelectedItem.ToString().Split('•')[0].Trim();

                foreach (Word.ContentControl cc in _wordApp.ActiveDocument.ContentControls)
                {
                    if (cc.Title == fieldName)
                    {
                        cc.Range.Select();
                        break;
                    }
                }
            }
            catch { }
        }

        private void LoadExistingFields()
        {
            try
            {
                lstFields.Items.Clear();
                cboExistingField.Items.Clear();
                existingFieldTags.Clear();

                if (_wordApp.ActiveDocument == null) return;

                foreach (Word.ContentControl cc in _wordApp.ActiveDocument.ContentControls)
                {
                    if (!string.IsNullOrEmpty(cc.Tag))
                    {
                        string[] parts = cc.Tag.Split('|');
                        string fieldName = parts.Length > 0 ? parts[0] : cc.Title;
                        string fieldType = parts.Length > 1 ? parts[1] : "Text";

                        lstFields.Items.Add($"{fieldName} • {fieldType}");

                        if (!existingFieldTags.ContainsKey(fieldName))
                        {
                            existingFieldTags[fieldName] = cc.Tag;
                            cboExistingField.Items.Add(fieldName);
                        }
                    }
                }

                if (cboExistingField.Items.Count > 0)
                    cboExistingField.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading fields: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RefreshQuestionPanel()
        {
            if (_isRefreshing) return;
            _isRefreshing = true;

            var currentValues = new Dictionary<string, object>();
            foreach (var kvp in fieldControls)
            {
                if (kvp.Value is TextBox tb && tb.ForeColor != Color.Gray)
                    currentValues[kvp.Key] = tb.Text;
                else if (kvp.Value is DateTimePicker dtp)
                    currentValues[kvp.Key] = dtp.Value;
                else if (kvp.Value is ComboBox cbo && cbo.SelectedItem != null)
                    currentValues[kvp.Key] = cbo.SelectedItem.ToString();
            }

            pnlQuestions.SuspendLayout();
            pnlQuestions.Controls.Clear();
            fieldControls.Clear();

            try
            {
                if (_wordApp.ActiveDocument == null)
                {
                    pnlQuestions.ResumeLayout();
                    _isRefreshing = false;
                    return;
                }

                var questionsLayout = new TableLayoutPanel
                {
                    AutoSize = true,
                    AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    ColumnCount = 1,
                    Dock = DockStyle.Top,
                    Padding = new Padding(Theme.PadMd)
                };
                questionsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

                int index = 0;
                var seenFieldNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (Word.ContentControl cc in _wordApp.ActiveDocument.ContentControls)
                {
                    if (string.IsNullOrEmpty(cc.Tag)) continue;

                    string[] parts = cc.Tag.Split('|');
                    if (parts.Length < 3) continue;

                    string fieldName = parts[0];
                    string fieldType = parts[1];
                    string prompt = parts[2];
                    string choices = parts.Length > 3 ? parts[3] : "";

                    if (!seenFieldNames.Add(fieldName)) continue;

                    var lblQuestion = new Label
                    {
                        Text = $"{index + 1}. {prompt}",
                        Dock = DockStyle.Fill,
                        AutoSize = true,
                        Font = new Font(Font.FontFamily, Font.Size, FontStyle.Bold),
                        ForeColor = Color.FromArgb(60, 60, 60),
                        Margin = new Padding(0, index > 0 ? Theme.PadMd : 0, 0, 4)
                    };
                    questionsLayout.Controls.Add(lblQuestion);

                    Control inputControl;
                    switch (fieldType)
                    {
                        case "Date":
                            inputControl = new DateTimePicker
                            {
                                Dock = DockStyle.Fill,
                                Format = DateTimePickerFormat.Short,
                                Font = Font,
                                Margin = new Padding(0, 0, 0, Theme.PadSm)
                            };
                            break;

                        case "Yes/No":
                            var cboYesNo = CreateComboBox();
                            cboYesNo.Items.AddRange(new object[] { "Yes", "No" });
                            cboYesNo.SelectedIndex = 0;
                            cboYesNo.Margin = new Padding(0, 0, 0, Theme.PadSm);
                            inputControl = cboYesNo;
                            break;

                        case "Multiple Choice":
                            var cboMulti = CreateComboBox();
                            if (!string.IsNullOrWhiteSpace(choices))
                            {
                                var choiceArray = choices.Split(',').Select(c => c.Trim()).Where(c => !string.IsNullOrEmpty(c)).ToArray();
                                cboMulti.Items.AddRange(choiceArray);
                                if (cboMulti.Items.Count > 0) cboMulti.SelectedIndex = 0;
                            }
                            cboMulti.Margin = new Padding(0, 0, 0, Theme.PadSm);
                            inputControl = cboMulti;
                            break;

                        default:
                            inputControl = CreateTextBox(fieldName);
                            inputControl.Margin = new Padding(0, 0, 0, Theme.PadSm);
                            break;
                    }

                    questionsLayout.Controls.Add(inputControl);
                    fieldControls[fieldName] = inputControl;

                    if (currentValues.TryGetValue(fieldName, out var prevValue))
                    {
                        if (inputControl is TextBox textBox && prevValue is string strVal)
                        {
                            textBox.Text = strVal;
                            textBox.ForeColor = Color.Black;
                        }
                        else if (inputControl is DateTimePicker dtPicker && prevValue is DateTime dtVal)
                        {
                            dtPicker.Value = dtVal;
                        }
                        else if (inputControl is ComboBox combo && prevValue is string comboVal)
                        {
                            int idx = combo.Items.IndexOf(comboVal);
                            if (idx >= 0) combo.SelectedIndex = idx;
                        }
                    }

                    index++;
                }

                if (index == 0)
                {
                    var lblNoFields = new Label
                    {
                        Text = "No fields found.\n\nSwitch to the Edit tab to add fields to your document.",
                        Dock = DockStyle.Fill,
                        ForeColor = Color.Gray,
                        TextAlign = ContentAlignment.MiddleCenter,
                        Font = Font,
                        Padding = new Padding(Theme.PadLg)
                    };
                    questionsLayout.Controls.Add(lblNoFields);
                }

                pnlQuestions.Controls.Add(questionsLayout);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading questions: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                pnlQuestions.ResumeLayout();
                _isRefreshing = false;
            }
        }

        private void BtnFillDocument_Click(object sender, EventArgs e)
        {
            try
            {
                if (_wordApp.ActiveDocument == null) return;

                int fieldsUpdated = 0;
                foreach (Word.ContentControl cc in _wordApp.ActiveDocument.ContentControls)
                {
                    if (string.IsNullOrEmpty(cc.Tag)) continue;

                    string[] parts = cc.Tag.Split('|');
                    if (parts.Length < 3) continue;

                    string fieldName = parts[0];

                    if (fieldControls.TryGetValue(fieldName, out Control ctrl))
                    {
                        string value = "";

                        if (ctrl is TextBox textBox)
                            value = GetTextBoxValue(textBox, fieldName);
                        else if (ctrl is DateTimePicker dtp)
                            value = dtp.Value.ToShortDateString();
                        else if (ctrl is ComboBox cbo)
                            value = cbo.SelectedItem?.ToString() ?? "";

                        if (!string.IsNullOrWhiteSpace(value))
                        {
                            cc.Range.Text = value;
                            cc.Range.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
                            fieldsUpdated++;
                        }
                    }
                }

                MessageBox.Show($"Document updated. {fieldsUpdated} field(s) filled.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error filling document: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GetTextBoxValue(TextBox textBox, string placeholder)
        {
            if (textBox == null) return "";
            if (textBox.ForeColor == Color.Gray && textBox.Text == placeholder) return "";
            return textBox.Text;
        }
    }
}
