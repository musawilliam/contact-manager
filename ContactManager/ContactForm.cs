using System;
using System.Drawing;
using System.Windows.Forms;

namespace ContactManager
{
    public partial class ContactForm : Form
    {
        public Contact Contact { get; set; }

        private Label lblName;
        private Label lblSurname;
        private Label lblNumber;
        
        private TextBox txtName;
        private TextBox txtSurname;
        private TextBox txtNumber;
        
        private CheckBox chkUsed;
        private Button btnSave;
        private Button btnCancel;

        public ContactForm(Contact contact = null)
        {
            InitializeComponent();
            
            if (contact != null)
            {
                Contact = contact;
                LoadContactData();
                this.Text = "Edit Contact";
            }
            else
            {
                Contact = new Contact();
                this.Text = "Add Contact";
            }
        }

        private void InitializeComponent()
        {
            this.lblName = new Label();
            this.lblSurname = new Label();
            this.lblNumber = new Label();
            
            this.txtName = new TextBox();
            this.txtSurname = new TextBox();
            this.txtNumber = new TextBox();
            
            this.chkUsed = new CheckBox();
            this.btnSave = new Button();
            this.btnCancel = new Button();
            this.SuspendLayout();

            // Labels
            this.lblName.Text = "Name:";
            this.lblName.Location = new Point(12, 15);
            this.lblName.Size = new Size(60, 23);

            this.lblSurname.Text = "Surname:";
            this.lblSurname.Location = new Point(12, 50);
            this.lblSurname.Size = new Size(60, 23);

            this.lblNumber.Text = "Number:";
            this.lblNumber.Location = new Point(12, 85);
            this.lblNumber.Size = new Size(60, 23);

            // TextBoxes
            this.txtName.Location = new Point(80, 12);
            this.txtName.Size = new Size(250, 23);

            this.txtSurname.Location = new Point(80, 47);
            this.txtSurname.Size = new Size(250, 23);

            this.txtNumber.Location = new Point(80, 82);
            this.txtNumber.Size = new Size(250, 23);

            // CheckBox
            this.chkUsed.Text = "Used";
            this.chkUsed.Location = new Point(80, 155);
            this.chkUsed.Size = new Size(100, 24);

            // Buttons
            this.btnSave.Text = "Save";
            this.btnSave.Location = new Point(175, 200);
            this.btnSave.Size = new Size(75, 30);
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += btnSave_Click;

            this.btnCancel.Text = "Cancel";
            this.btnCancel.Location = new Point(255, 200);
            this.btnCancel.Size = new Size(75, 30);
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += btnCancel_Click;

            // Form
            this.ClientSize = new Size(350, 250);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            
            this.Controls.Add(this.lblName);
            this.Controls.Add(this.lblSurname);
            this.Controls.Add(this.lblNumber);
            
            this.Controls.Add(this.txtName);
            this.Controls.Add(this.txtSurname);
            this.Controls.Add(this.txtNumber);
            
            this.Controls.Add(this.chkUsed);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnCancel);

            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private void LoadContactData()
        {
            txtName.Text = Contact.Name;
            txtSurname.Text = Contact.Surname;
            txtNumber.Text = Contact.Number;
            
            chkUsed.Checked = Contact.Used;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (ValidateInput())
            {
                Contact.Name = txtName.Text.Trim();
                Contact.Surname = txtSurname.Text.Trim();
                Contact.Number = txtNumber.Text.Trim();
                
                Contact.Used = chkUsed.Checked;
                
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        private bool ValidateInput()
        {
            if (string.IsNullOrWhiteSpace(txtName.Text))
            {
                MessageBox.Show("Name is required.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtName.Focus();
                return false;
            }
            
            return true;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}