namespace DocumentRecognition
{
    using System;
    using System.Collections.Generic;

    using System.Drawing;
    using System.Text;
    using System.Windows.Forms;
    using DocumentRecognition.Recognition;
    using OpenCvSharp;
    using OpenCvSharp.Extensions;
    using static Alignment;
    using RectO = OpenCvSharp.Rect;
    using PointO = OpenCvSharp.Point;
    using MatO = OpenCvSharp.Mat;
    using Mat = Emgu.CV.Mat;
    //using Tesseract;
    using Emgu.CV;
    using Emgu.CV.Structure;
    using Emgu.CV.Util;
    using Emgu.CV.CvEnum;
    using System.Linq;
    using ClosedXML.Excel;
    using Emgu.CV.OCR;

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



                //первый способ
                var recognitionedText = Recognition.Recognition.RecognizeTable(selectionedPicture);
                //второй способ
                var recognitionTable = TableRecognition(selectionedPicture);




                textBox1.Text = recognitionedText.Text;

                textBox2.Text = Convert.ToString(recognitionedText.Accuracy);

            }
        }

        private void TableDetection(int NumberCols = 4, float MorphThrehold = 30f,
            int binaryThreshold = 200, int offset = 5)
        {
            try
            {
                if (pictureBox1 == null)
                    return;

                var img = new Bitmap(pictureBox1.Image).ToImage<Gray, byte>().ThresholdBinaryInv(new Gray(binaryThreshold), new Gray(255)); //в градации серого, где мы предполагаем
                                                                                                                                            //что фон белый и можем указать порог при конфертировании в белый
                                                                                                                                            //слова -белый, фон -темный

                //Находим линии таблиц (горизонтальны/вертикальные (нужно корректировать MorphThrehold ))
                int lenght = (int)(img.Width * MorphThrehold / 100);
                int Height = (int)(img.Height * 7f / 100);

                Mat vProfile = new Mat();
                Mat hProfile = new Mat();

                var kernelV = CvInvoke.GetStructuringElement(Emgu.CV.CvEnum.ElementShape.Rectangle, new System.Drawing.Size(1, Height), new System.Drawing.Point(-1,-1));
                var kernelH = CvInvoke.GetStructuringElement(Emgu.CV.CvEnum.ElementShape.Rectangle, new System.Drawing.Size(lenght, 1), new System.Drawing.Point(-1, -1));

                CvInvoke.Erode(img, vProfile, kernelV, new System.Drawing.Point(-1, -1), 1, Emgu.CV.CvEnum.BorderType.Default, new MCvScalar(255));
                CvInvoke.Dilate(vProfile, vProfile, kernelV, new System.Drawing.Point(-1, -1), 1, Emgu.CV.CvEnum.BorderType.Default, new MCvScalar(255));

                CvInvoke.Erode(img, hProfile, kernelH, new System.Drawing.Point(-1, -1), 1, Emgu.CV.CvEnum.BorderType.Default, new MCvScalar(255));
                CvInvoke.Dilate(hProfile, hProfile, kernelH, new System.Drawing.Point(-1, -1), 1, Emgu.CV.CvEnum.BorderType.Default, new MCvScalar(255));

                //соединение найденных линий в одно изображение
                var mergedImage = vProfile.ToImage<Gray, byte>().Or(hProfile.ToImage<Gray, byte>());
                mergedImage._ThresholdBinary(new Gray(1), new Gray(255));

                //Ищем внешние границы таблицы
                VectorOfVectorOfPoint contours = new VectorOfVectorOfPoint();
                Mat h = new Mat();

                //Поиск контуров
                CvInvoke.FindContours(mergedImage, contours, h, RetrType.External, ChainApproxMethod.ChainApproxSimple);
                int bigCID = GetBiggestContourID(contours);

                //Из контуров распознаем самый большой прямоугольник (внешние границы таблицы)
                var bbox = CvInvoke.BoundingRectangle(contours[bigCID]);

                mergedImage.ROI = bbox;
                //обрезаем изначальное изображение
                img.ROI = bbox;

                var temp = mergedImage.Copy();
                //Инвертируем цвет
                temp._Not();

                //Находим все прямоугольники (внутренние и внешние)
                var imgTable = img.Copy();
                contours.Clear();

                //Поиск контуров
                CvInvoke.FindContours(temp, contours, h, RetrType.External, ChainApproxMethod.ChainApproxSimple);

                var filtercontours = FilterContours(contours, 500);
                //Из контуров распознаем прямоугольники (внешние границы таблицы)
                var bboxList = Contours2BBox(filtercontours);

                //Упорядовачиваем
                var sortedBBoxes = bboxList.OrderBy(x => x.Y).ThenBy(x => x.X).ToList();

                //---------OCR

                Tesseract engine = new Tesseract(@"DataWords", "rus", OcrEngineMode.TesseractOnly);
                engine.PageSegMode = PageSegMode.SingleBlock;

                // write to excel 

                var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Recognition");

                int rowCounter = 1;
                char colCounter = 'A';

                for (int i = 0; i < sortedBBoxes.Count; i++)
                {
                    var rect = sortedBBoxes[i];
                    rect.X += offset;
                    rect.Y += offset;
                    rect.Width -= offset;
                    rect.Height -= offset;

                    imgTable.ROI = rect;
                    engine.SetImage(imgTable.Copy());

                    string text = engine.GetUTF8Text().Replace("\r\n", "");

                    if (i % NumberCols == 0)
                    {
                        if (i > 0)
                        {
                            rowCounter++;
                        }
                        colCounter = 'A';
                        worksheet.Cell(colCounter.ToString() + rowCounter.ToString()).Value = text;
                    }
                    else
                    {
                        colCounter++;
                        worksheet.Cell(colCounter + rowCounter.ToString()).Value = text;
                    }
                    imgTable.ROI = Rectangle.Empty;

                }
                string outputpath = @"C:\Users\SanzhievDB\Documents\output.xlsx";
                workbook.SaveAs(outputpath);
                MessageBox.Show("Обнаружение завершено\n" + outputpath);


                pictureBox2.Image = temp.ToBitmap();


            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private VectorOfVectorOfPoint FilterContours(VectorOfVectorOfPoint contours, double threshold = 50)
        {
            VectorOfVectorOfPoint filteredContours = new VectorOfVectorOfPoint();
            for (int i = 0; i < contours.Size; i++)
            {
                if (CvInvoke.ContourArea(contours[i]) >= threshold)
                {
                    filteredContours.Push(contours[i]);
                }
            }

            return filteredContours;
        }

        private List<Rectangle> Contours2BBox(VectorOfVectorOfPoint contours)
        {
            List<Rectangle> list = new List<Rectangle>();
            for(int i = 0; i<contours.Size; i++)
            {
                list.Add(CvInvoke.BoundingRectangle(contours[i]));
            }
            return list;
        }


        private int GetBiggestContourID(VectorOfVectorOfPoint contours)
        {
            double maxArea = double.MaxValue * (-1);
            int contourId = -1;
            for (int i = 0; i < contours.Size; i++)
            {
                double area = CvInvoke.ContourArea(contours[i]);
                if(area > maxArea)
                {
                    maxArea = area;
                    contourId = i;
                }
            }
            return contourId;
        }

        private void startToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                TableDetection();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private double angle(PointO p1, PointO p2, PointO p0)
        {
            double dx1 = p1.X - p0.X;
            double dy1 = p1.Y - p0.Y;
            double dx2 = p2.X - p0.X;
            double dy2 = p2.Y - p0.Y;
            return (dx1 * dx2 + dy1 * dy2) / Math.Sqrt((dx1 * dx1 + dy1 * dy1) * (dx2 * dx2 + dy2 * dy2) + 1e-10);
        }

        public Document TableRecognition(Bitmap selectionedPicture)
        {
            MatO image = BitmapConverter.ToMat(selectionedPicture);
            // Преобразуем изображение в черно-белый формат
            MatO grayImage = new MatO();
            Cv2.CvtColor(image, grayImage, ColorConversionCodes.BGR2GRAY);
            Cv2.AdaptiveThreshold(grayImage, grayImage, 255, AdaptiveThresholdTypes.GaussianC, ThresholdTypes.Binary, 75, 10);

            // Находим контуры на изображении
            OpenCvSharp.Point[][] contours;
            HierarchyIndex[] hierarchy;
            Cv2.FindContours(grayImage, out contours, out hierarchy, RetrievalModes.Tree, ContourApproximationModes.ApproxSimple);

            // Получаем прямоугольник, описывающий каждый контур
            List<RectO> rectangles = new List<RectO>();
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
                        RectO rectangle = Cv2.BoundingRect(contour);
                        rectangles.Add(rectangle);
                    }
                }
            }

            // Отображаем прямоугольники на изображении
            foreach (RectO rectangle in rectangles)
            {
                Cv2.Rectangle(image, rectangle, new Scalar(0, 255, 0), 2);
            }

            var doc = new Document();
            // Распознаем текст в каждом прямоугольнике
            //StringBuilder resultText = new StringBuilder();
            //using (TesseractEngine engine = new TesseractEngine(@"DataWords", "rus", EngineMode.Default))
            //{
            //    foreach (RectO rectangle in rectangles)
            //    {
            //        MatO roi = new MatO(image, rectangle);
            //        Bitmap bitmapImage = roi.ToBitmap();

            //        pictureBox4.Image = bitmapImage;

            //        // Конвертируем Mat в Pix
            //        using (Pix pix = PixConverter.ToPix(bitmapImage))
            //        {
            //            using (Page page = engine.Process(pix, PageSegMode.SingleBlock))
            //            {
            //                string text = page.GetText();
            //                resultText.AppendLine(text);
            //                doc.Accuracy = page.GetMeanConfidence();
            //            }
            //        }
            //    }
            //}
            // Выводим результат 
            //doc.Text = resultText.ToString();

            return doc;

        }

    }
}
