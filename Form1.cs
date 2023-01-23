using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using MsoTriState = Microsoft.Office.Core.MsoTriState;
using Microsoft.Office.Interop.PowerPoint;

namespace WinFormPowerPoint
{
    public partial class Form1 : Form
    { 
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            // ���� ���� �ʱ�ȭ
            string text = lyricText.Text;

            // �Ŀ�����Ʈ �ν��Ͻ� ���� �� �߰�
            PowerPoint.Application objApp = new PowerPoint.Application();
            PowerPoint._Presentation objPres = objApp.Presentations.Add(MsoTriState.msoTrue);
            objApp.Visible = MsoTriState.msoTrue;

            // �����̵� ���� ����
            PowerPoint.Slides objSlides = objPres.Slides;

            // �����̵� ��� ���������� ���� ����
            Color backgroundColor = Color.Black;

            // �� ���� ���� �� text�� �迭 element�� �и�
            string[] textArr = text.Split(new[] { "\n\n" }, StringSplitOptions.RemoveEmptyEntries);

            for (int i = textArr.Length - 1; i >= 0; i--)
            {
                // �����̵� ���� �߰�
                PowerPoint._Slide objSlide = objSlides.Add(1, PpSlideLayout.ppLayoutBlank);

                // �����̵� ��� ���������� ����
                objSlide.FollowMasterBackground = MsoTriState.msoFalse;
                objSlide.Background.Fill.ForeColor.RGB = backgroundColor.ToArgb();

                // �����̵忡 �ؽ�Ʈ ���� �߰� �� ��Ÿ�� ����
                objSlide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 10, 215, 940, 100);
                PowerPoint.TextRange objTextRng = objSlide.Shapes[1].TextFrame.TextRange;
                objTextRng.Text = textArr[i];
                objTextRng.Font.Size = 50;
                objTextRng.Font.Color.RGB = Color.White.ToArgb();
                objTextRng.Font.Bold = MsoTriState.msoTrue;
                objTextRng.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // ���� ��� ���� �ʱ�ȭ
            var filePath = string.Empty;

            // ���� ���� �˾�
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                // �ʱ� ��� ����
                openFileDialog.InitialDirectory = "c:\\";

                // ���� ����
                openFileDialog.Filter = "ppt files (*.ppt)|*.ppt|pptx files (*.pptx)|*.pptx";

                // �ε��� ����
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // ���� ��� ����
                    filePath = openFileDialog.FileName;
                }
            }

            // ���ο� �Ŀ�����Ʈ ���� ���� �� ���� ����
            PowerPoint.Application objApp = new PowerPoint.Application();
            objApp.Visible = MsoTriState.msoTrue;
            Presentations objPresens = objApp.Presentations;
            PowerPoint._Presentation objPres = objApp.Presentations.Open(filePath, MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoTrue);

            // ���̵� ��ũ������ �����ϱ�
            objApp.ActivePresentation.PageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeOnScreen16x9;
        }
    }
}