using OpenCvSharp;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tesseract;
using Tesseract.Interop;
using OpenCvSharp.Extensions;
using Point = OpenCvSharp.Point;
using Rect = OpenCvSharp.Rect;
using System.IO;
using System.Windows.Forms;

namespace DocumentRecognition.Recognition
{
    public class Recognition
    {
        public static Document Recognize (Bitmap bitmap)
        {
            var doc = new Document();
            var pixBytes = ImageToByte(bitmap);
            using (var engine = new TesseractEngine(@"DataWords", "rus", EngineMode.Default))
            {

                using (var pix = Pix.LoadFromMemory(pixBytes))
                {
                    using (var recognizedPage = engine.Process(pix, PageSegMode.SingleBlock))
                    {
                        doc.Accuracy = recognizedPage.GetMeanConfidence();
                        doc.Text = recognizedPage.GetText();
                        return doc;
                    }
                }      
   
            }
        }

        public static Document RecognizeTable(Bitmap bitmap)
        {
            Mat image = BitmapConverter.ToMat(bitmap);
            // Преобразуем изображение в черно-белый формат
            Mat grayImage = new Mat();
            Cv2.CvtColor(image, grayImage, ColorConversionCodes.BGR2GRAY);
            Cv2.AdaptiveThreshold(grayImage, grayImage, 255, AdaptiveThresholdTypes.GaussianC, ThresholdTypes.Binary, 75, 10);

            // Находим контуры на изображении
            Point[][] contours;
            HierarchyIndex[] hierarchy;
            Cv2.FindContours(grayImage, out contours, out hierarchy, RetrievalModes.Tree, ContourApproximationModes.ApproxSimple);

            // Получаем прямоугольник, описывающий каждый контур
            List<Rect> rectangles = new List<Rect>();
            foreach (Point[] contour in contours)
            {
                Rect rectangle = Cv2.BoundingRect(contour);
                rectangles.Add(rectangle);
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

                    var pixBytes = ImageToByte(bitmap);

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
            return doc;
        }


        public static byte[] ImageToByte(Image img)
        {
            ImageConverter converter = new ImageConverter();
            return (byte[])converter.ConvertTo(img, typeof(byte[]));
        }
    }
}
