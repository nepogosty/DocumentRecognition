namespace DocumentRecognition
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.Drawing;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using AForge.Imaging;
    using AForge.Imaging.Filters;
    using DocumentRecognition.Recognition;
    using OpenCvSharp;
    using OpenCvSharp.Extensions;
    using Tesseract;
    using static Alignment;
    using Rect = OpenCvSharp.Rect;
    using Point = OpenCvSharp.Point;
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Загрузка изображения.
        /// </summary>
        private void button1_Click(object sender, EventArgs e)
        {
            // openFileDialog1.Filter = "JPG|*.jpg|PNG|*.png|";
            if (openFileDialog1.ShowDialog() ==  DialogResult.OK )
            {
                pictureBox1.Image = System.Drawing.Image.FromFile(openFileDialog1.FileName);
            }
        }

        /// <summary>
        /// Приведение к нормальному виду.
        /// </summary>
        private void button2_Click(object sender, EventArgs e)
        {
            if (pictureBox1.Image != null)
            {
                var pictureToAlignment = new Bitmap(pictureBox1.Image);
                var rdyPicture = DocumentAngleCorrection(pictureToAlignment);
                pictureBox2.Image = rdyPicture;

                var selectionedPicture = SelectionDocument.DocumentCropInfo(rdyPicture);
                pictureBox3.Image = selectionedPicture;

                //var recognitionedText = Recognition.Recognition.RecognizeTable(selectionedPicture);

                Mat image = BitmapConverter.ToMat(selectionedPicture);
                // Преобразуем изображение в черно-белый формат
                Mat grayImage = new Mat();
                Cv2.CvtColor(image, grayImage, ColorConversionCodes.BGR2GRAY);
                Cv2.AdaptiveThreshold(grayImage, grayImage, 255, AdaptiveThresholdTypes.GaussianC, ThresholdTypes.Binary, 75, 10);

                // Находим контуры на изображении
                OpenCvSharp.Point[][] contours;
                HierarchyIndex[] hierarchy;
                Cv2.FindContours(grayImage, out contours, out hierarchy, RetrievalModes.Tree, ContourApproximationModes.ApproxSimple);

                // Получаем прямоугольник, описывающий каждый контур
                List<Rect> rectangles = new List<Rect>();
                foreach (OpenCvSharp.Point[] contour in contours)
                {
                    // Выполняем аппроксимацию контура многоугольником
                    OpenCvSharp.Point[] approx = Cv2.ApproxPolyDP(contour, Cv2.ArcLength(contour, true) * 0.02, true);

                    // Определяем, является ли многоугольник прямоугольником
                    if (approx.Length == 4 && Cv2.IsContourConvex(approx))
                    {
                        double maxCosine = 0;

                        for (int i = 2; i < 5; i++)
                        {
                            double cosine = Math.Abs(angle(approx[i % 4], approx[i - 2], approx[i - 1]));
                            maxCosine = Math.Max(maxCosine, cosine);
                        }

                        if (maxCosine < 0.3)
                        {
                            Rect rectangle = Cv2.BoundingRect(contour);
                            rectangles.Add(rectangle);
                        }
                    }
                }

                // Отображаем прямоугольники на изображении
                foreach (Rect rectangle in rectangles)
                {
                    Cv2.Rectangle(image, rectangle, new Scalar(0, 255, 0), 2);
                }

                var doc = new Document();
                // Распознаем текст в каждом прямоугольнике
                StringBuilder resultText = new StringBuilder();
                using (TesseractEngine engine = new TesseractEngine(@"DataWords", "rus", EngineMode.Default))
                {
                    foreach (Rect rectangle in rectangles)
                    {
                        Mat roi = new Mat(image, rectangle);
                        Bitmap bitmapImage = roi.ToBitmap();

                        pictureBox4.Image = bitmapImage;

                        // Конвертируем Mat в Pix
                        using (Pix pix = PixConverter.ToPix(bitmapImage))
                        {
                            using (Page page = engine.Process(pix, PageSegMode.SingleBlock))
                            {
                                string text = page.GetText();
                                resultText.AppendLine(text);
                                doc.Accuracy = page.GetMeanConfidence();
                            }
                        }
                    }
                }

                // Выводим результат 

                doc.Text = resultText.ToString();



                textBox1.Text = doc.Text;

                textBox2.Text = Convert.ToString(doc.Accuracy);

            }
        }
        private double angle(Point p1, Point p2, Point p0)
        {
            double dx1 = p1.X - p0.X;
            double dy1 = p1.Y - p0.Y;
            double dx2 = p2.X - p0.X;
            double dy2 = p2.Y - p0.Y;
            return (dx1 * dx2 + dy1 * dy2) / Math.Sqrt((dx1 * dx1 + dy1 * dy1) * (dx2 * dx2 + dy2 * dy2) + 1e-10);
        }
    }
}
