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
            // 가사 변수 초기화
            string text = lyricText.Text;

            // 파워포인트 인스턴스 생성 및 추가
            PowerPoint.Application objApp = new PowerPoint.Application();
            PowerPoint._Presentation objPres = objApp.Presentations.Add(MsoTriState.msoTrue);
            objApp.Visible = MsoTriState.msoTrue;

            // 슬라이드 변수 선언
            PowerPoint.Slides objSlides = objPres.Slides;

            // 슬라이드 배경 검정색으로 변수 지정
            Color backgroundColor = Color.Black;

            // 빈 줄이 있을 때 text를 배열 element로 분리
            string[] textArr = text.Split(new[] { "\n\n" }, StringSplitOptions.RemoveEmptyEntries);

            for (int i = textArr.Length - 1; i >= 0; i--)
            {
                // 슬라이드 새로 추가
                PowerPoint._Slide objSlide = objSlides.Add(1, PpSlideLayout.ppLayoutBlank);

                // 슬라이드 배경 검정색으로 설정
                objSlide.FollowMasterBackground = MsoTriState.msoFalse;
                objSlide.Background.Fill.ForeColor.RGB = backgroundColor.ToArgb();

                // 슬라이드에 텍스트 상자 추가 및 스타일 지정
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
            // 파일 경로 변수 초기화
            var filePath = string.Empty;

            // 파일 열기 팝업
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                // 초기 경로 세팅
                openFileDialog.InitialDirectory = "c:\\";

                // 파일 필터
                openFileDialog.Filter = "ppt files (*.ppt)|*.ppt|pptx files (*.pptx)|*.pptx";

                // 인덱스 설정
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // 파일 경로 지정
                    filePath = openFileDialog.FileName;
                }
            }

            // 새로운 파워포인트 변수 설정 및 파일 열기
            PowerPoint.Application objApp = new PowerPoint.Application();
            objApp.Visible = MsoTriState.msoTrue;
            Presentations objPresens = objApp.Presentations;
            PowerPoint._Presentation objPres = objApp.Presentations.Open(filePath, MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoTrue);

            // 와이드 스크린으로 변경하기
            objApp.ActivePresentation.PageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeOnScreen16x9;
        }
    }
}