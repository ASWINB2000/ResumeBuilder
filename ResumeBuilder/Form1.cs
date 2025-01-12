using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Drawing;
using Color = System.Drawing.Color;
using Control = System.Windows.Forms.Control;
using Font = System.Drawing.Font;


namespace ResumeBuilder
{
    public partial class Form1 : Form
    {
        private TextBox txtName;
        private TextBox txtEmail;
        private TextBox txtPhone;
        private TextBox txtAddress;
        private TextBox txtSummary;
        private TextBox txtEducation;
        private TextBox txtExperience;
        private TextBox txtSkills;
        private Button btnGenerate;
        private RichTextBox previewBox;


        public Form1()
        {
            InitializeComponent();
            SetupUI();
            SetupEventHandlers();
        }


        private void SetupUI()
        {
            this.Text = "Resume Builder with Live Preview";
            this.Size = new Size(1000, 800);


            Panel inputPanel = new Panel
            {
                Location = new Point(20, 20),
                Size = new Size(450, 700),
                AutoScroll = true
            };
            Panel previewPanel = new Panel
            {
                Location = new Point(490, 20),
                Size = new Size(450, 700),
                BorderStyle = BorderStyle.FixedSingle
            };


            int currentY = 10;
            int labelX = 10;
            int controlX = 100;
            int spacing = 30;


            Label lblName = new Label { Text = "Full Name:", Location = new Point(labelX, currentY) };
            txtName = new TextBox { Location = new Point(controlX, currentY), Width = 300 };
            currentY += spacing;


            Label lblEmail = new Label { Text = "Email:", Location = new Point(labelX, currentY) };
            txtEmail = new TextBox { Location = new Point(controlX, currentY), Width = 300 };
            currentY += spacing;


            Label lblPhone = new Label { Text = "Phone:", Location = new Point(labelX, currentY) };
            txtPhone = new TextBox { Location = new Point(controlX, currentY), Width = 300 };
            currentY += spacing;


            Label lblAddress = new Label { Text = "Address:", Location = new Point(labelX, currentY) };
            txtAddress = new TextBox { Location = new Point(controlX, currentY), Width = 300 };
            currentY += spacing;


            Label lblSummary = new Label { Text = "Summary:", Location = new Point(labelX, currentY) };
            txtSummary = new TextBox
            {
                Location = new Point(controlX, currentY),
                Width = 300,
                Multiline = true,
                Height = 100
            };
            currentY += spacing + 100;


            Label lblEducation = new Label { Text = "Education:", Location = new Point(labelX, currentY) };
            txtEducation = new TextBox
            {
                Location = new Point(controlX, currentY),
                Width = 300,
                Multiline = true,
                Height = 100
            };
            currentY += spacing + 100;


            Label lblExperience = new Label { Text = "Experience:", Location = new Point(labelX, currentY) };
            txtExperience = new TextBox
            {
                Location = new Point(controlX, currentY),
                Width = 300,
                Multiline = true,
                Height = 100
            };
            currentY += spacing + 100;


            Label lblSkills = new Label { Text = "Skills:", Location = new Point(labelX, currentY) };
            txtSkills = new TextBox
            {
                Location = new Point(controlX, currentY),
                Width = 300,
                Multiline = true,
                Height = 100
            };
            currentY += spacing + 100;


            btnGenerate = new Button
            {
                Text = "Generate Resume",
                Location = new Point(controlX, currentY),
                Width = 150,
                Height = 40
            };


            previewBox = new RichTextBox
            {
                Location = new Point(10, 10),
                Size = new Size(430, 680),
                ReadOnly = true,
                BackColor = Color.White
            };


            inputPanel.Controls.AddRange(new Control[] {
                lblName, txtName,
                lblEmail, txtEmail,
                lblPhone, txtPhone,
                lblAddress, txtAddress,
                lblSummary, txtSummary,
                lblEducation, txtEducation,
                lblExperience, txtExperience,
                lblSkills, txtSkills,
                btnGenerate
            });


            previewPanel.Controls.Add(previewBox);
            this.Controls.AddRange(new Control[] { inputPanel, previewPanel });
        }


        private void SetupEventHandlers()
        {
            txtName.TextChanged += UpdatePreview;
            txtEmail.TextChanged += UpdatePreview;
            txtPhone.TextChanged += UpdatePreview;
            txtAddress.TextChanged += UpdatePreview;
            txtSummary.TextChanged += UpdatePreview;
            txtEducation.TextChanged += UpdatePreview;
            txtExperience.TextChanged += UpdatePreview;
            txtSkills.TextChanged += UpdatePreview;
            btnGenerate.Click += BtnGenerate_Click;
        }


        private void UpdatePreview(object sender, EventArgs e)
        {
            previewBox.Clear();
            previewBox.SelectionFont = new Font("Arial", 16, FontStyle.Bold);
            previewBox.AppendText($"{txtName.Text}\n\n");


            previewBox.SelectionFont = new Font("Arial", 10, FontStyle.Regular);
            previewBox.AppendText($"Email: {txtEmail.Text}\n");
            previewBox.AppendText($"Phone: {txtPhone.Text}\n");
            previewBox.AppendText($"Address: {txtAddress.Text}\n\n");


            previewBox.SelectionFont = new Font("Arial", 12, FontStyle.Bold);
            previewBox.AppendText("Professional Summary\n");
            previewBox.SelectionFont = new Font("Arial", 10, FontStyle.Regular);
            previewBox.AppendText($"{txtSummary.Text}\n\n");


            previewBox.SelectionFont = new Font("Arial", 12, FontStyle.Bold);
            previewBox.AppendText("Education\n");
            previewBox.SelectionFont = new Font("Arial", 10, FontStyle.Regular);
            previewBox.AppendText($"{txtEducation.Text}\n\n");


            previewBox.SelectionFont = new Font("Arial", 12, FontStyle.Bold);
            previewBox.AppendText("Experience\n");
            previewBox.SelectionFont = new Font("Arial", 10, FontStyle.Regular);
            previewBox.AppendText($"{txtExperience.Text}\n\n");


            previewBox.SelectionFont = new Font("Arial", 12, FontStyle.Bold);
            previewBox.AppendText("Skills\n");
            previewBox.SelectionFont = new Font("Arial", 10, FontStyle.Regular);
            previewBox.AppendText($"{txtSkills.Text}\n");
        }


        private void BtnGenerate_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "Word Document|*.docx",
                Title = "Save Resume",
                FileName = "Resume.docx"
            };


            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                GenerateWordDocument(saveDialog.FileName);
                MessageBox.Show("Resume generated successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }


        private void GenerateWordDocument(string filePath)
{
    using (WordprocessingDocument doc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
    {
        MainDocumentPart mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());


        Body body = mainPart.Document.Body;


        // Add content to the document
        AddParagraph(body, txtName.Text, true, "36");
        AddParagraph(body, $"Email: {txtEmail.Text}");
        AddParagraph(body, $"Phone: {txtPhone.Text}");
        AddParagraph(body, $"Address: {txtAddress.Text}");


        AddParagraph(body, "Professional Summary", true, "28");
        AddParagraph(body, txtSummary.Text);


        AddParagraph(body, "Education", true, "28");
        foreach (var edu in txtEducation.Text.Split('\n'))
        {
            if (!string.IsNullOrWhiteSpace(edu))
                AddParagraph(body, "• " + edu.Trim());
        }


        AddParagraph(body, "Experience", true, "28");
        foreach (var exp in txtExperience.Text.Split('\n'))
        {
            if (!string.IsNullOrWhiteSpace(exp))
                AddParagraph(body, "• " + exp.Trim());
        }


        AddParagraph(body, "Skills", true, "28");
        foreach (var skill in txtSkills.Text.Split('\n'))
        {
            if (!string.IsNullOrWhiteSpace(skill))
                AddParagraph(body, "• " + skill.Trim());
        }
    }
}


private void AddParagraph(Body body, string text, bool isBold = false, string fontSize = "24")
{
    Paragraph para = body.AppendChild(new Paragraph());
    Run run = para.AppendChild(new Run());
    RunProperties runProperties = run.AppendChild(new RunProperties());
    if (isBold)
    {
        runProperties.AppendChild(new Bold());
    }
    runProperties.AppendChild(new FontSize { Val = fontSize });
    run.AppendChild(new Text(text));
}
    }
}

