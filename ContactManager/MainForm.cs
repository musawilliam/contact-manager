using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace ContactManager
{
    public partial class MainForm : Form
    {
        private DataService _dataService;
        private ExcelService _excelService;
        private List<Contact> _contacts;

        private MenuStrip menuStrip1;
        private ToolStripMenuItem fileMenuItem;
        private ToolStripMenuItem importExcelMenuItem;
        private ToolStripMenuItem exportExcelMenuItem;
        private ToolStripMenuItem exitMenuItem;
        private ToolStripMenuItem recordsMenuItem;
        private ToolStripMenuItem addRecordMenuItem;
        private ToolStripMenuItem editRecordMenuItem;
        private ToolStripMenuItem deleteRecordMenuItem;
        private DataGridView dataGridView1;
        private Button btnAdd;
        private Button btnEdit;
        private Button btnDelete;
        private StatusStrip statusStrip1;
        private ToolStripStatusLabel toolStripStatusLabel1;

        public MainForm()
        {
            InitializeComponent();
            _dataService = new DataService();
            _excelService = new ExcelService();
            _contacts = new List<Contact>();
            LoadContacts();
            SetupDataGridView();
        }

        private void InitializeComponent()
        {
            this.menuStrip1 = new MenuStrip();
            this.fileMenuItem = new ToolStripMenuItem();
            this.importExcelMenuItem = new ToolStripMenuItem();
            this.exportExcelMenuItem = new ToolStripMenuItem();
            this.exitMenuItem = new ToolStripMenuItem();
            this.recordsMenuItem = new ToolStripMenuItem();
            this.addRecordMenuItem = new ToolStripMenuItem();
            this.editRecordMenuItem = new ToolStripMenuItem();
            this.deleteRecordMenuItem = new ToolStripMenuItem();
            this.dataGridView1 = new DataGridView();
            this.btnAdd = new Button();
            this.btnEdit = new Button();
            this.btnDelete = new Button();
            this.statusStrip1 = new StatusStrip();
            this.toolStripStatusLabel1 = new ToolStripStatusLabel();
            
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();

            // MenuStrip
            this.menuStrip1.Items.AddRange(new ToolStripItem[] {
                this.fileMenuItem,
                this.recordsMenuItem
            });
            this.menuStrip1.Location = new Point(0, 0);
            this.menuStrip1.Size = new Size(800, 24);

            // File Menu
            this.fileMenuItem.Text = "File";
            this.fileMenuItem.DropDownItems.AddRange(new ToolStripItem[] {
                this.importExcelMenuItem,
                this.exportExcelMenuItem,
                new ToolStripSeparator(),
                this.exitMenuItem
            });

            this.importExcelMenuItem.Text = "Import Excel";
            this.importExcelMenuItem.Click += ImportExcelMenuItem_Click;

            this.exportExcelMenuItem.Text = "Export Excel";
            this.exportExcelMenuItem.Click += ExportExcelMenuItem_Click;

            this.exitMenuItem.Text = "Exit";
            this.exitMenuItem.Click += ExitMenuItem_Click;

            // Records Menu
            this.recordsMenuItem.Text = "Records";
            this.recordsMenuItem.DropDownItems.AddRange(new ToolStripItem[] {
                this.addRecordMenuItem,
                this.editRecordMenuItem,
                this.deleteRecordMenuItem
            });

            this.addRecordMenuItem.Text = "Add";
            this.addRecordMenuItem.Click += BtnAdd_Click;

            this.editRecordMenuItem.Text = "Edit";
            this.editRecordMenuItem.Click += BtnEdit_Click;

            this.deleteRecordMenuItem.Text = "Delete";
            this.deleteRecordMenuItem.Click += BtnDelete_Click;

            // DataGridView
            this.dataGridView1.Location = new Point(12, 35);
            this.dataGridView1.Size = new Size(776, 350);
            this.dataGridView1.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            this.dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.ColumnHeaderMouseClick += DataGridView1_ColumnHeaderMouseClick;

            // Buttons
            this.btnAdd.Text = "Add";
            this.btnAdd.Location = new Point(12, 395);
            this.btnAdd.Size = new Size(75, 30);
            this.btnAdd.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += BtnAdd_Click;

            this.btnEdit.Text = "Edit";
            this.btnEdit.Location = new Point(100, 395);
            this.btnEdit.Size = new Size(75, 30);
            this.btnEdit.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            this.btnEdit.UseVisualStyleBackColor = true;
            this.btnEdit.Click += BtnEdit_Click;

            this.btnDelete.Text = "Delete";
            this.btnDelete.Location = new Point(188, 395);
            this.btnDelete.Size = new Size(75, 30);
            this.btnDelete.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Click += BtnDelete_Click;

            // StatusStrip
            this.statusStrip1.Items.AddRange(new ToolStripItem[] {
                this.toolStripStatusLabel1
            });
            this.statusStrip1.Location = new Point(0, 435);
            this.statusStrip1.Size = new Size(800, 22);

            this.toolStripStatusLabel1.Text = "Ready";

            // MainForm
            this.ClientSize = new Size(800, 457);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnEdit);
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Text = "Contact Manager";
            this.WindowState = FormWindowState.Maximized;

            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private void LoadContacts()
        {
            _contacts = _dataService.GetAllContacts();
            RefreshGrid();
        }

        private void RefreshGrid()
        {
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = _contacts;
            
            // Hide ID column
            if (dataGridView1.Columns["Id"] != null)
                dataGridView1.Columns["Id"].Visible = false;
            
            // Format columns
            if (dataGridView1.Columns["CreatedDate"] != null)
            {
                dataGridView1.Columns["CreatedDate"].HeaderText = "Created Date";
                dataGridView1.Columns["CreatedDate"].DefaultCellStyle.Format = "yyyy-MM-dd HH:mm";
            }
            
            ApplyRowColors();
        }

        private void ApplyRowColors()
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.DataBoundItem is Contact contact)
                {
                    if (contact.Used)
                    {
                        row.DefaultCellStyle.BackColor = Color.LightGreen;
                    }
                    else
                    {
                        row.DefaultCellStyle.BackColor = Color.LightCoral;
                    }
                }
            }
        }

        private void SetupDataGridView()
        {
            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.MultiSelect = false;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.ReadOnly = true;
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            using (var form = new ContactForm())
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    _dataService.AddContact(form.Contact);
                    LoadContacts();
                }
            }
        }

        private void BtnEdit_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                var selectedContact = (Contact)dataGridView1.SelectedRows[0].DataBoundItem;
                using (var form = new ContactForm(selectedContact))
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        _dataService.UpdateContact(form.Contact);
                        LoadContacts();
                    }
                }
            }
            else
            {
                MessageBox.Show("Please select a contact to edit.", "No Selection", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                var selectedContact = (Contact)dataGridView1.SelectedRows[0].DataBoundItem;
                var result = MessageBox.Show($"Are you sure you want to delete {selectedContact.Name}?", 
                    "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                
                if (result == DialogResult.Yes)
                {
                    _dataService.DeleteContact(selectedContact.Id);
                    LoadContacts();
                }
            }
            else
            {
                MessageBox.Show("Please select a contact to delete.", "No Selection", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ImportExcelMenuItem_Click(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                openFileDialog.Title = "Select Excel file to import";
                
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var importedContacts = _excelService.ImportFromExcel(openFileDialog.FileName);
                    
                    if (importedContacts.Count > 0)
                    {
                        var result = MessageBox.Show(
                            $"Found {importedContacts.Count} contacts. Do you want to clear existing data first?",
                            "Import Contacts", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                        
                        if (result == DialogResult.Cancel) return;
                        
                        if (result == DialogResult.Yes)
                        {
                            _dataService.ClearAllContacts();
                        }
                        
                        foreach (var contact in importedContacts)
                        {
                            _dataService.AddContact(contact);
                        }
                        
                        LoadContacts();
                        MessageBox.Show($"Successfully imported {importedContacts.Count} contacts.", 
                            "Import Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }

        private void ExportExcelMenuItem_Click(object sender, EventArgs e)
        {
            using (var saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                saveFileDialog.Title = "Export contacts to Excel";
                saveFileDialog.FileName = $"contacts_export_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    _excelService.ExportToExcel(_contacts, saveFileDialog.FileName);
                }
            }
        }

        private void ExitMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void DataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            // Handle sorting
            var column = dataGridView1.Columns[e.ColumnIndex];
            var sortOrder = ListSortDirection.Ascending;
            
            if (dataGridView1.Tag != null && dataGridView1.Tag.ToString() == column.Name)
            {
                sortOrder = ListSortDirection.Descending;
                dataGridView1.Tag = null;
            }
            else
            {
                dataGridView1.Tag = column.Name;
            }
            
            // Sort the data
            switch (column.Name)
            {
                case "Name":
                    _contacts = sortOrder == ListSortDirection.Ascending 
                        ? _contacts.OrderBy(c => c.Name).ToList()
                        : _contacts.OrderByDescending(c => c.Name).ToList();
                    break;
                case "Surname":
                    _contacts = sortOrder == ListSortDirection.Ascending 
                        ? _contacts.OrderBy(c => c.Surname).ToList()
                        : _contacts.OrderByDescending(c => c.Surname).ToList();
                    break;
                case "Number":
                    _contacts = sortOrder == ListSortDirection.Ascending 
                        ? _contacts.OrderBy(c => c.Number).ToList()
                        : _contacts.OrderByDescending(c => c.Number).ToList();
                    break;
                case "Used":
                    _contacts = sortOrder == ListSortDirection.Ascending 
                        ? _contacts.OrderBy(c => c.Used).ToList()
                        : _contacts.OrderByDescending(c => c.Used).ToList();
                    break;
                case "CreatedDate":
                    _contacts = sortOrder == ListSortDirection.Ascending 
                        ? _contacts.OrderBy(c => c.CreatedDate).ToList()
                        : _contacts.OrderByDescending(c => c.CreatedDate).ToList();
                    break;
            }
            
            RefreshGrid();
        }
    }
}