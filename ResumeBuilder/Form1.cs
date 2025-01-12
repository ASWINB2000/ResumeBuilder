using System;
using System.Drawing;
using System.Windows.Forms;
using System.Drawing.Text;

namespace ResumeBuilder
{
    public partial class Form1 : Form
    {
        private Panel previewPanel;
        private const float FooterPositionRatio = 0.95f;

        public Form1()
        {
            InitializeComponent();
            SetupPanels();
            this.Resize += Form1_Resize;
            this.AutoScroll = true;
        }

        private void SetupPanels()
        {
            this.Size = new Size(1600, 1200);
            this.BackColor = Color.Gray;
            int a4Width = 794;
            int a4Height = 1123;
            this.MinimumSize = new Size(800, 600);
            this.AutoScrollMinSize = new Size(1600, 1200);

            previewPanel = new Panel()
            {
                Size = new Size(a4Width, a4Height),
                Location = new Point(this.Width - a4Width - 40, 20),
                BackColor = Color.White,
                AutoScroll = true,
                AutoScrollMinSize = new Size(a4Width, a4Height + 200)
            };

            SetupPanelContents(previewPanel, a4Width, a4Height);
            this.Controls.Add(previewPanel);
            UpdateLayout();
        }

        private void SetupPanelContents(Panel panel, int width, int height)
        {
            PrivateFontCollection pfc = new PrivateFontCollection();
            pfc.AddFontFile("C:/Users/Aswin/Desktop/F-projects/ASP.NET/ResumeBuilder/ResumeBuilder/Fonts/OpenSans-VariableFont_wdth,wght.ttf");
            pfc.AddFontFile("C:/Users/Aswin/Desktop/F-projects/ASP.NET/ResumeBuilder/ResumeBuilder/Fonts/Copperplate Gothic Std 29 AB/Copperplate Gothic Std 29 AB.otf");

            var headerLine = new Panel()
            {
                Size = new Size(width - 80, 7),
                Location = new Point(40, 10),
                BackColor = Color.Black
            };

            var footerLine = new Panel()
            {
                Size = new Size(width - 80, 7),
                BackColor = Color.Black
            };

            var cosq = new Label()
            {
                Text = "COSQ NETWORK PVT LTD",
                Font = new Font(pfc.Families[1], 12, FontStyle.Regular),
                AutoSize = true
            };

            var pageNumber = new Label()
            {
                Text = "1",
                AutoSize = true,
                Font = new Font(pfc.Families[0], 12, FontStyle.Bold)
            };

            panel.Controls.Add(headerLine);
            panel.Controls.Add(footerLine);
            panel.Controls.Add(cosq);
            panel.Controls.Add(pageNumber);

            panel.Tag = new { FooterLine = footerLine, COSQ = cosq, PageNumber = pageNumber };
        }

        private void UpdateLayout()
        {
            UpdatePanelLayout(previewPanel);
        }

        private void UpdatePanelLayout(Panel panel)
        {
            var footerElements = (dynamic)panel.Tag;
            int panelHeight = panel.ClientSize.Height;
            int footerY = (int)(panelHeight * FooterPositionRatio);

            footerElements.FooterLine.Location = new Point(40, footerY);
            footerElements.COSQ.Location = new Point(40, footerY + 15);
            footerElements.PageNumber.Location = new Point(panel.ClientSize.Width - 60, footerY + 15);
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            UpdateLayout();
        }
    }
}
