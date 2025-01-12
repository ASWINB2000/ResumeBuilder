using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Drawing;
using Color = System.Drawing.Color;
using Control = System.Windows.Forms.Control;
using Font = System.Drawing.Font;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;



namespace ResumeBuilder
{
    public partial class Form1 : Form
    {
        private TextBox txtName;
        private TextBox txtEmail;
        private TextBox txtPhone;
        private TextBox txtAddress;
        private TextBox txtLink;
        private PictureBox profilePicture;
        private Button btnSelectImage;
        private TextBox txtSummary;
        private TextBox txtEducation;
        private TextBox txtExperience;
        private TextBox txtSkills;
        private TextBox txtLanguages;
        private TextBox txtCertifications;
        private TextBox txtProjects;
        private Button btnGenerate;
        private RichTextBox previewBox;
        private StatusStrip statusStrip;

        int a4Width = 794;  // 210mm
        int a4Height = 1123;
        public Form1()
        {
            InitializeComponent();
            SetupUI();
            SetupEventHandlers();
            AddTooltips();
        }
        private void SetupUI()
        {
            this.Text = "Resume Builder with Live Preview";
            this.Size = new Size(1200, 800);
            this.BackColor = Color.FromArgb(240, 240, 240);

            Panel inputPanel = new Panel
            {
                Location = new Point(20, 20),
                Size = new Size(550, 700),
                AutoScroll = true,
                BorderStyle = BorderStyle.None,
                BackColor = Color.White,
                Padding = new Padding(20),
                Margin = new Padding(10)
            };

            Panel previewPanel = new Panel
            {
                Location = new Point(590, 20),
                Size = new Size(a4Width + 20, a4Height + 20), // Add small margin
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White,
                Padding = new Padding(10)
            };

            int currentY = 10;
            int labelX = 10;
            int controlX = 150;
            int spacing = 50;

            // Personal Information Section
            Label lblPersonalInfo = CreateSectionHeader("PERSONAL INFORMATION", new Point(labelX, currentY));
            currentY += 30;

            Label lblName = CreateLabel("Full Name:", new Point(labelX, currentY));
            txtName = CreateStyledTextBox(new Point(controlX, currentY), 350, false);
            currentY += spacing;

            Label lblEmail = CreateLabel("Email:", new Point(labelX, currentY));
            txtEmail = CreateStyledTextBox(new Point(controlX, currentY), 350, false);
            currentY += spacing;

            Label lblPhone = CreateLabel("Phone:", new Point(labelX, currentY));
            txtPhone = CreateStyledTextBox(new Point(controlX, currentY), 350, false);
            currentY += spacing;

            Label lblAddress = CreateLabel("Address:", new Point(labelX, currentY));
            txtAddress = CreateStyledTextBox(new Point(controlX, currentY), 350, true, 60);
            currentY += spacing + 40;

            Label lblLink = CreateLabel("Portfolio Link:", new Point(labelX, currentY));
            txtLink = CreateStyledTextBox(new Point(controlX, currentY), 350, false);
            currentY += spacing + 20;
            Label lblProfilePic = CreateSectionHeader("PROFILE PICTURE", new Point(labelX, currentY));
            currentY += 30;

            profilePicture = new PictureBox
            {
                Location = new Point(controlX, currentY),
                Size = new Size(150, 150),
                SizeMode = PictureBoxSizeMode.StretchImage,
                BorderStyle = BorderStyle.FixedSingle
            };
            btnSelectImage = new Button
            {
                Text = "Select Image",
                Location = new Point(controlX + 160, currentY + 60),
                Size = new Size(100, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White
            };
            currentY += 170;
            // Summary Section
            Label lblSummary = CreateSectionHeader("SUMMARY", new Point(labelX, currentY));
            currentY += 30;
            txtSummary = CreateStyledTextBox(new Point(controlX, currentY), 350, true, 100);
            currentY += spacing + 80;

            // Skills Section
            Label lblSkills = CreateSectionHeader("SKILLS AND PROFICIENCIES", new Point(labelX, currentY));
            currentY += 30;
            txtSkills = CreateStyledTextBox(new Point(controlX, currentY), 350, true, 150);
            currentY += spacing + 130;

            // Languages Section
            Label lblLanguages = CreateSectionHeader("LANGUAGES", new Point(labelX, currentY));
            currentY += 30;
            txtLanguages = CreateStyledTextBox(new Point(controlX, currentY), 350, true, 100);
            currentY += spacing + 80;

            // Education Section
            Label lblEducation = CreateSectionHeader("EDUCATION", new Point(labelX, currentY));
            currentY += 30;
            txtEducation = CreateStyledTextBox(new Point(controlX, currentY), 350, true, 100);
            currentY += spacing + 80;

            // Experience Section
            Label lblExperience = CreateSectionHeader("WORK EXPERIENCE", new Point(labelX, currentY));
            currentY += 30;
            txtExperience = CreateStyledTextBox(new Point(controlX, currentY), 350, true, 150);
            currentY += spacing + 130;

            // Projects Section
            Label lblProjects = CreateSectionHeader("PROJECTS", new Point(labelX, currentY));
            currentY += 30;
            txtProjects = CreateStyledTextBox(new Point(controlX, currentY), 350, true, 200);
            currentY += spacing + 180;

            // Certifications Section
            Label lblCertifications = CreateSectionHeader("CERTIFICATIONS", new Point(labelX, currentY));
            currentY += 30;
            txtCertifications = CreateStyledTextBox(new Point(controlX, currentY), 350, true, 100);
            currentY += spacing + 80;

            // Generate Button
            btnGenerate = new Button
            {
                Text = "Generate Resume",
                Location = new Point(controlX, currentY),
                Width = 150,
                Height = 40,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10f, FontStyle.Bold)
            };

            // Preview Box
            previewBox = new RichTextBox
            {
                Location = new Point(10, 10),
                Size = new Size(a4Width - 20, a4Height - 20), // Account for padding
                ReadOnly = true,
                BackColor = Color.White,
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 10f),
                Padding = new Padding(40, 10, 40, 10),
                ScrollBars = RichTextBoxScrollBars.Vertical
            };

            inputPanel.Controls.Add(lblProfilePic);
            inputPanel.Controls.Add(profilePicture);
            inputPanel.Controls.Add(btnSelectImage);

            // Add all controls to the input panel
            inputPanel.Controls.AddRange(new Control[] {
            lblPersonalInfo,
            lblName, txtName,
            lblEmail, txtEmail,
            lblPhone, txtPhone,
            lblAddress, txtAddress,
            lblLink, txtLink,
            lblSummary, txtSummary,
            lblSkills, txtSkills,
            lblLanguages, txtLanguages,
            lblEducation, txtEducation,
            lblExperience, txtExperience,
            lblProjects, txtProjects,
            lblCertifications, txtCertifications,
            btnGenerate
        });

            previewPanel.Controls.Add(previewBox);
            this.Controls.AddRange(new Control[] { inputPanel, previewPanel, statusStrip });
        }


        private Label CreateLabel(string text, Point location)
        {
            return new Label
            {
                Text = text,
                Location = location,
                Font = new Font("Segoe UI", 9f),
                ForeColor = Color.FromArgb(60, 60, 60),
                AutoSize = true
            };
        }

        private Label CreateSectionHeader(string text, Point location)
        {
            return new Label
            {
                Text = text,
                Location = location,
                Font = new Font("Segoe UI", 11f, FontStyle.Bold),
                ForeColor = Color.FromArgb(40, 40, 40),
                AutoSize = true
            };
        }

        private TextBox CreateStyledTextBox(Point location, int width, bool multiline, int height = 25)
        {
            TextBox textBox = new TextBox
            {
                Location = location,
                Width = width,
                Height = height,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9.5f),
                Padding = new Padding(5)
            };

            if (multiline)
            {
                textBox.Multiline = true;
                textBox.ScrollBars = ScrollBars.Vertical;
            }

            return textBox;
        }

        private void AddTooltips()
        {
            ToolTip tooltip = new ToolTip();

            tooltip.SetToolTip(txtName, "Enter your full name");
            tooltip.SetToolTip(txtEmail, "Enter your professional email address");
            tooltip.SetToolTip(txtPhone, "Enter your contact number");
            tooltip.SetToolTip(txtAddress, "Enter your current address");
            tooltip.SetToolTip(txtLink, "Enter your portfolio or LinkedIn URL");

            tooltip.SetToolTip(txtSummary, "Write a brief overview of your professional background, skills, and career goals");
            tooltip.SetToolTip(txtSkills, "List your technical skills, tools, programming languages, etc. (one per line)");
            tooltip.SetToolTip(txtLanguages, "List languages you know and proficiency level (e.g., 'English - Native', 'Japanese - N5')");

            tooltip.SetToolTip(txtEducation, "List your educational qualifications with institution names and years");
            tooltip.SetToolTip(txtExperience, "List your work experience with company name, position, duration, and key responsibilities");
            tooltip.SetToolTip(txtProjects, "Describe your significant projects with technology stack and key features");
            tooltip.SetToolTip(txtCertifications, "List relevant certifications with issuing organization and date");
        }

        private void SetupEventHandlers()
        {
            txtName.TextChanged += UpdatePreview!;
            txtEmail.TextChanged += UpdatePreview!;
            txtPhone.TextChanged += UpdatePreview!;
            txtAddress.TextChanged += UpdatePreview!;
            txtLink.TextChanged += UpdatePreview!;
            btnSelectImage.Click += (s, e) =>
{
    using (OpenFileDialog openFileDialog = new OpenFileDialog())
    {
        openFileDialog.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.gif;*.bmp";
        openFileDialog.Title = "Select Profile Picture";

        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
            profilePicture.Image = Image.FromFile(openFileDialog.FileName);
        }
    }
};
            txtSummary.TextChanged += UpdatePreview!;
            txtSkills.TextChanged += UpdatePreview!;
            txtLanguages.TextChanged += UpdatePreview!;
            txtEducation.TextChanged += UpdatePreview!;
            txtExperience.TextChanged += UpdatePreview!;
            txtProjects.TextChanged += UpdatePreview!;
            txtCertifications.TextChanged += UpdatePreview!;
            btnGenerate.Click += BtnGenerate_Click!;

        }


        private void UpdatePreview(object sender, EventArgs e)
        {
            previewBox.Clear();
            previewBox.SelectionFont = new Font("Arial", 12, FontStyle.Bold);
            previewBox.AppendText("_________________________________________________\n\n");
            // Name
            previewBox.SelectionFont = new Font("Arial", 18, FontStyle.Bold);
            previewBox.AppendText($"{txtName.Text.ToUpper()}\n\n");

            // Contact Information
            previewBox.SelectionFont = new Font("Arial", 10, FontStyle.Regular);
            previewBox.AppendText($"Phone: {txtPhone.Text}\n");
            previewBox.AppendText($"Email: {txtEmail.Text}\n");
            previewBox.AppendText($"Address: {txtAddress.Text}\n");
            if (!string.IsNullOrWhiteSpace(txtLink.Text))
            {
                previewBox.AppendText($"Link: {txtLink.Text}\n");
            }
            previewBox.AppendText("\n");

            // Summary Section
            previewBox.SelectionFont = new Font("Arial", 12, FontStyle.Bold);
            previewBox.AppendText("SUMMARY\n");
            previewBox.SelectionFont = new Font("Arial", 10, FontStyle.Regular);
            previewBox.AppendText($"{txtSummary.Text}\n\n");

            // Skills Section
            previewBox.SelectionFont = new Font("Arial", 12, FontStyle.Bold);
            previewBox.AppendText("SKILLS AND PROFICIENCIES\n");
            previewBox.SelectionFont = new Font("Arial", 10, FontStyle.Regular);
            string[] skills = txtSkills.Text.Split('\n');
            foreach (var skill in skills)
            {
                if (!string.IsNullOrWhiteSpace(skill))
                {
                    previewBox.AppendText($"• {skill.Trim()}\n");
                }
            }
            previewBox.AppendText("\n");

            // Languages Section
            previewBox.SelectionFont = new Font("Arial", 12, FontStyle.Bold);
            previewBox.AppendText("LANGUAGES\n");
            previewBox.SelectionFont = new Font("Arial", 10, FontStyle.Regular);
            string[] languages = txtLanguages.Text.Split('\n');
            foreach (var language in languages)
            {
                if (!string.IsNullOrWhiteSpace(language))
                {
                    previewBox.AppendText($"{language.Trim()}\n");
                }
            }
            previewBox.AppendText("\n");

            // Education Section
            previewBox.SelectionFont = new Font("Arial", 12, FontStyle.Bold);
            previewBox.AppendText("EDUCATION\n");
            previewBox.SelectionFont = new Font("Arial", 10, FontStyle.Regular);
            string[] education = txtEducation.Text.Split('\n');
            foreach (var edu in education)
            {
                if (!string.IsNullOrWhiteSpace(edu))
                {
                    previewBox.AppendText($"{edu.Trim()}\n");
                }
            }
            previewBox.AppendText("\n");

            // Work Experience Section
            previewBox.SelectionFont = new Font("Arial", 12, FontStyle.Bold);
            previewBox.AppendText("WORK EXPERIENCE\n");
            previewBox.SelectionFont = new Font("Arial", 10, FontStyle.Regular);
            string[] experiences = txtExperience.Text.Split('\n');
            foreach (var exp in experiences)
            {
                if (!string.IsNullOrWhiteSpace(exp))
                {
                    previewBox.AppendText($"{exp.Trim()}\n");
                }
            }
            previewBox.AppendText("\n");

            // Projects Section
            previewBox.SelectionFont = new Font("Arial", 12, FontStyle.Bold);
            previewBox.AppendText("PROJECTS\n");
            previewBox.SelectionFont = new Font("Arial", 10, FontStyle.Regular);
            string[] projects = txtProjects.Text.Split('\n');
            foreach (var project in projects)
            {
                if (!string.IsNullOrWhiteSpace(project))
                {
                    previewBox.AppendText($"{project.Trim()}\n");
                }
            }
            previewBox.AppendText("\n");

            // Certifications Section
            previewBox.SelectionFont = new Font("Arial", 12, FontStyle.Bold);
            previewBox.AppendText("CERTIFICATIONS\n");
            previewBox.SelectionFont = new Font("Arial", 10, FontStyle.Regular);
            string[] certifications = txtCertifications.Text.Split('\n');
            foreach (var cert in certifications)
            {
                if (!string.IsNullOrWhiteSpace(cert))
                {
                    previewBox.AppendText($"{cert.Trim()}\n");
                }
            }
            previewBox.SelectionFont = new Font("Arial", 12, FontStyle.Bold);
            previewBox.AppendText("\n_________________________________________________\n\n");

            previewBox.SelectionFont = new Font("Copperplate Gothic Std", 12, FontStyle.Regular);
            previewBox.AppendText("COSQ NETWORK PVT LTD");
            previewBox.SelectionAlignment = HorizontalAlignment.Right;
            previewBox.SelectionFont = new Font("Arial", 12, FontStyle.Bold);
            previewBox.AppendText("1");

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


        private void AddParagraph(Body body, string text, bool isBold = false, string fontSize = "24",
    JustificationValues? justification = null)
        {
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            RunProperties runProperties = run.AppendChild(new RunProperties());

            if (isBold)
            {
                runProperties.AppendChild(new Bold());
            }

            runProperties.AppendChild(new FontSize { Val = fontSize });
            runProperties.AppendChild(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" });

            // Add justification if specified
            if (justification.HasValue)
            {
                para.AppendChild(new Justification { Val = justification.Value });
            }

            run.AppendChild(new Text(text));
        }

        private void GenerateWordDocument(string filePath)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                Body body = mainPart.Document.Body;
                Paragraph headerLinePara = body.AppendChild(new Paragraph());
                Run headerLineRun = headerLinePara.AppendChild(new Run());
                headerLineRun.AppendChild(new Break { Type = BreakValues.TextWrapping });
                headerLinePara.AppendChild(
                    new ParagraphProperties(
                        new ParagraphBorders(
                            new TopBorder { Val = BorderValues.Single, Size = 24, Space = 1, Color = "000000" }
                        )
                    )
                );

                // Name (Large, Bold)
                AddParagraph(body, txtName.Text.ToUpper(), true, "36", JustificationValues.Left);

                // Contact Information
                AddParagraph(body, $"Phone: {txtPhone.Text}", false, "24");
                AddParagraph(body, $"Email: {txtEmail.Text}", false, "24");
                AddParagraph(body, $"Address: {txtAddress.Text}", false, "24");
                if (!string.IsNullOrWhiteSpace(txtLink.Text))
                {
                    AddParagraph(body, $"Link: {txtLink.Text}", false, "24");
                }
                AddParagraph(body, "", false, "24"); // Empty line for spacing
                                                     // Add this after creating the document body
                if (profilePicture.Image != null)
                {
                    ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                    using (MemoryStream ms = new MemoryStream())
                    {
                        profilePicture.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                        ms.Position = 0;
                        imagePart.FeedData(ms);
                    }

                    var element =
                         new Drawing(
                             new DW.Inline(
                                 new DW.Extent() { Cx = 990000L, Cy = 990000L },
                                 new DW.EffectExtent()
                                 {
                                     LeftEdge = 0L,
                                     TopEdge = 0L,
                                     RightEdge = 0L,
                                     BottomEdge = 0L
                                 },
                                 new DW.DocProperties()
                                 {
                                     Id = 1U,
                                     Name = "Profile Picture"
                                 },
                                 new DW.NonVisualGraphicFrameDrawingProperties(
                                     new A.GraphicFrameLocks() { NoChangeAspect = true }),
                                 new A.Graphic(
                                     new A.GraphicData(
                                         new PIC.Picture(
                                             new PIC.NonVisualPictureProperties(
                                                 new PIC.NonVisualDrawingProperties()
                                                 {
                                                     Id = 0U,
                                                     Name = "Profile.jpg"
                                                 },
                                                 new PIC.NonVisualPictureDrawingProperties()),
                                             new PIC.BlipFill(
                                                 new A.Blip(
                                                     new A.BlipExtensionList(
                                                         new A.BlipExtension()
                                                         {
                                                             Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                         })
                                                 )
                                                 {
                                                     Embed = mainPart.GetIdOfPart(imagePart),
                                                     CompressionState =
                                                     A.BlipCompressionValues.Print
                                                 },
                                                 new A.Stretch(
                                                     new A.FillRectangle())),
                                             new PIC.ShapeProperties(
                                                 new A.Transform2D(
                                                     new A.Offset() { X = 0L, Y = 0L },
                                                     new A.Extents() { Cx = 990000L, Cy = 990000L }),
                                                 new A.PresetGeometry(
                                                     new A.AdjustValueList()
                                                 )
                                                 { Preset = A.ShapeTypeValues.Rectangle }))
                                     )
                                     { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                             )
                             {
                                 DistanceFromTop = 0U,
                                 DistanceFromBottom = 0U,
                                 DistanceFromLeft = 0U,
                                 DistanceFromRight = 0U,
                                 EditId = "50D07946"
                             });

                    body.AppendChild(new Paragraph(new Run(element)));
                }

                // Summary Section
                AddParagraph(body, "SUMMARY", true, "28", JustificationValues.Left);
                AddParagraph(body, txtSummary.Text, false, "24");
                AddParagraph(body, "", false, "24");

                // Skills Section
                AddParagraph(body, "SKILLS AND PROFICIENCIES", true, "28", JustificationValues.Left);
                string[] skills = txtSkills.Text.Split('\n');
                foreach (var skill in skills)
                {
                    if (!string.IsNullOrWhiteSpace(skill))
                    {
                        AddParagraph(body, "• " + skill.Trim(), false, "24");
                    }
                }
                AddParagraph(body, "", false, "24");

                // Languages Section
                AddParagraph(body, "LANGUAGES", true, "28", JustificationValues.Left);
                string[] languages = txtLanguages.Text.Split('\n');
                foreach (var language in languages)
                {
                    if (!string.IsNullOrWhiteSpace(language))
                    {
                        AddParagraph(body, language.Trim(), false, "24");
                    }
                }
                AddParagraph(body, "", false, "24");

                // Education Section
                AddParagraph(body, "EDUCATION", true, "28", JustificationValues.Left);
                string[] education = txtEducation.Text.Split('\n');
                foreach (var edu in education)
                {
                    if (!string.IsNullOrWhiteSpace(edu))
                    {
                        AddParagraph(body, edu.Trim(), false, "24");
                    }
                }
                AddParagraph(body, "", false, "24");

                // Work Experience Section
                AddParagraph(body, "WORK EXPERIENCE", true, "28", JustificationValues.Left);
                string[] experiences = txtExperience.Text.Split('\n');
                foreach (var exp in experiences)
                {
                    if (!string.IsNullOrWhiteSpace(exp))
                    {
                        AddParagraph(body, exp.Trim(), false, "24");
                    }
                }
                AddParagraph(body, "", false, "24");

                // Projects Section
                AddParagraph(body, "PROJECTS", true, "28", JustificationValues.Left);
                string[] projects = txtProjects.Text.Split('\n');
                foreach (var project in projects)
                {
                    if (!string.IsNullOrWhiteSpace(project))
                    {
                        AddParagraph(body, project.Trim(), false, "24");
                    }
                }
                AddParagraph(body, "", false, "24");

                // Certifications Section
                AddParagraph(body, "CERTIFICATIONS", true, "28", JustificationValues.Left);
                string[] certifications = txtCertifications.Text.Split('\n');
                foreach (var cert in certifications)
                {
                    if (!string.IsNullOrWhiteSpace(cert))
                    {
                        AddParagraph(body, cert.Trim(), false, "24");
                    }
                }
                Paragraph footerPara = body.AppendChild(new Paragraph(
           new ParagraphProperties(
               new ParagraphBorders(
                   new BottomBorder { Val = BorderValues.Single, Size = 24, Space = 1, Color = "000000" }
               )
           )
       ));

                Run footerRun = footerPara.AppendChild(new Run());
                footerRun.AppendChild(new Break { Type = BreakValues.TextWrapping });

                // Add COSQ text and page number
                Paragraph cosqPara = body.AppendChild(new Paragraph(
                    new Run(
                        new RunProperties(new FontSize { Val = "24" }),
                        new Text("COSQ NETWORK PVT LTD")
                    ),
                    new Run(
                        new RunProperties(new FontSize { Val = "24" }),
                        new Text("1") { Space = SpaceProcessingModeValues.Preserve }
                    )
                ));

                cosqPara.ParagraphProperties = new ParagraphProperties(
                    new Justification { Val = JustificationValues.Left }
                );
            }
        }
    }
}


