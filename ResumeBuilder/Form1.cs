
// using System.Drawing.Text;

// namespace ResumeBuilder
// {
//     public partial class Form1 : Form
//     {
//         private Panel pagePanel;
//         private Panel previewPanel;
//         private Panel headerLine;
//         private Panel footerLine;
//         private Label pageNumber;
//         private Label COSQ;
//         public Form1()
//     {
//         InitializeComponent();
//         SetupPanels();
//     }

//         private void SetupPanels()
//         {
//             this.Size = new Size(1600, 1200);
//             this.BackColor = Color.Gray;
//             int a4Width = 794;
//             int a4Height = 1123;

//             pagePanel = new Panel()
//             {
//                 Size = new Size(a4Width, a4Height),
//                 Location= new Point(20,20),
//                 BackColor = Color.White,
//                 AutoScroll = true,
//                 AutoScrollMinSize = new Size(0, a4Height + 50)
//             };
//             previewPanel = new Panel()
//             {
//                 Size = new Size(a4Width, a4Height),
//                 Location = new Point(a4Width + 40, 20),
//                 BackColor = Color.White,
//                 AutoScroll = true,
//                 AutoScrollMinSize = new Size(0, a4Height + 50)
//             };
//             headerLine = new Panel()
//             {
//                 Size = new Size(a4Width - 80, 7),
//                 Location = new Point(40, 10),
//                 BackColor = Color.Black
//             };
//             var previewHeaderLine = new Panel()
//             {
//                 Size = new Size(a4Width - 80, 7),
//                 Location = new Point(40, 10),
//                 BackColor = Color.Black
//             };
//             footerLine = new Panel()
//             {
//                 Size = new Size(a4Width - 80, 7),
//                 Location = new Point(40, (int)(a4Height * 0.82)), 
//                 BackColor = Color.Black
//             };
//             var previewFooterLine  = new Panel()
//             {
//                 Size = new Size(a4Width - 80, 7),
//                 Location = new Point(40, (int)(a4Height * 0.82)), 
//                 BackColor = Color.Black
//             };
//             PrivateFontCollection pfc = new PrivateFontCollection();
//             pfc.AddFontFile("C:/Users/Aswin/Desktop/F-projects/ASP.NET/ResumeBuilder/ResumeBuilder/Fonts/OpenSans-VariableFont_wdth,wght.ttf");
//             pfc.AddFontFile("C:/Users/Aswin/Desktop/F-projects/ASP.NET/ResumeBuilder/ResumeBuilder/Fonts/Copperplate Gothic Std 29 AB/Copperplate Gothic Std 29 AB.otf");
//             COSQ = new Label()
//             {
//                 Text = "COSQ NETWORK PVT LTD",
//                 Font = new Font(pfc.Families[1], 12, FontStyle.Regular),
//                 Location = new Point(40, (int)(a4Height * 0.85)),
//                 AutoSize = true
//             };
//             var previewCOSQ = new Label
//         {
//             Text = "COSQ NETWORK PVT LTD",
//             Location = new Point(40, (int)(a4Height * 0.85)),
//             AutoSize = true,
//             Font = new Font(pfc.Families[1], 12, FontStyle.Regular)
//         };
//             pageNumber = new Label()
//             {
//                 Text="1",
//                 Location = new Point(a4Width - 60, (int)(a4Height * 0.85)),
//                 AutoSize= true,
//                 Font = new Font(pfc.Families[0], 12, FontStyle.Bold)
//             };
//             var previewPageNumber = new Label
//             {
//                 Text = "1",
//                 Location = new Point(a4Width - 60, (int)(a4Height * 0.85)),
//                 AutoSize = true,
//                 Font = new Font(pfc.Families[0], 12, FontStyle.Bold)
//             };
//                 // Add controls to the editing panel
//             pagePanel.Controls.Add(headerLine);
//             pagePanel.Controls.Add(footerLine);
//             pagePanel.Controls.Add(COSQ);
//             pagePanel.Controls.Add(pageNumber);

//             // Add controls to the preview panel
//             previewPanel.Controls.Add(previewHeaderLine);
//             previewPanel.Controls.Add(previewFooterLine);
//             previewPanel.Controls.Add(previewCOSQ);
//             previewPanel.Controls.Add(previewPageNumber);

//             // Add both panels to the form
//             this.Controls.Add(pagePanel);
//             this.Controls.Add(previewPanel);

//         }
//     }

// }




