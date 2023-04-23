using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentRecognition.Recognition
{
    public class SelectionDocument
    {
        public static Bitmap DocumentCropInfo(Bitmap image)
        {
            const double widthProcImage = 1000;
            const int sensitivity = 25;
            const int treshold = 50;
            const int widthQuantum = 100;


            var sourceImage = image;
            var sizeFactor = widthProcImage / sourceImage.Width;
            var procBtmp = new Bitmap(sourceImage, (int)Math.Round(sourceImage.Width * sizeFactor), (int)Math.Round(sourceImage.Height * sizeFactor));
            var bitmapData = procBtmp.LockBits(new Rectangle(0, 0, procBtmp.Width, procBtmp.Height), ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);
            var bytes = Math.Abs(bitmapData.Stride) * procBtmp.Height;
            var sourceBytes = new byte[bytes];

            System.Runtime.InteropServices.Marshal.Copy(bitmapData.Scan0, sourceBytes, 0, bytes);

            var x1 = procBtmp.Width;
            var y1 = procBtmp.Height;
            var x2 = 0;
            var y2 = 0;
            var maxV = 0;

            var pointList = new List<Point>();
            var regionSize = bitmapData.Width / widthQuantum;

            for (var y = 0; y < bitmapData.Height + regionSize; y += regionSize)
            { // x processing
                for (var x = 0; x < bitmapData.Width + regionSize; x += regionSize)
                { // y processing
                    var value = 0;

                    for (var yy = y; (yy < y + regionSize) && (yy < bitmapData.Height); yy++)
                    { // Horosontal counting
                        var pixel = Alignment.GetGrayPixel(sourceBytes, bitmapData.Width, x, yy);

                        for (var xx = x; (xx < x + regionSize) && (xx < bitmapData.Width); xx++)
                        {
                            var nextPixel = Alignment.GetGrayPixel(sourceBytes, bitmapData.Width, xx, yy);

                            if (Math.Abs(pixel - nextPixel) > sensitivity)
                            {
                                value++;
                            }

                            pixel = nextPixel;
                        }
                    }

                    for (var xx = x; (xx < x + regionSize) && (xx < bitmapData.Width); xx++)
                    { // Vertical counting
                        var pixel = Alignment.GetGrayPixel(sourceBytes, bitmapData.Width, xx, y);

                        for (var yy = y; (yy < y + regionSize) && (yy < bitmapData.Height); yy++)
                        {
                            var nextPixel = Alignment.GetGrayPixel(sourceBytes, bitmapData.Width, xx, yy);

                            if (Math.Abs(pixel - nextPixel) > sensitivity)
                            {
                                value++;
                            }

                            pixel = nextPixel;
                        }
                    }

                    pointList.Add(new Point() { V = value, X = x, Y = y });
                    maxV = Math.Max(maxV, value);
                }
            }

            var vFactor = 255.0 / maxV;

            foreach (var point in pointList)
            {
                var v = (byte)(point.V * vFactor);

                if (v > treshold)
                {
                    x1 = Math.Min(x1, point.X);
                    y1 = Math.Min(y1, point.Y);

                    x2 = Math.Max(x2, point.X + regionSize);
                    y2 = Math.Max(y2, point.Y + regionSize);
                }
            }

            procBtmp.UnlockBits(bitmapData);

            x1 = (int)Math.Round((x1 - regionSize) / sizeFactor);
            x2 = (int)Math.Round((x2 + regionSize) / sizeFactor);
            y1 = (int)Math.Round((y1 - regionSize) / sizeFactor);
            y2 = (int)Math.Round((y2 + regionSize) / sizeFactor);

            var bigRect = new Rectangle(x1, y1, x2 - x1, y2 - y1);
            var clippedImg = CropImage(sourceImage, bigRect);

            return clippedImg;
        }
        public static Bitmap CropImage(Bitmap source, Rectangle section)
        {
            section.X = Math.Max(0, section.X);
            section.Y = Math.Max(0, section.Y);
            section.Width = Math.Min(source.Width, section.Width);
            section.Height = Math.Min(source.Height, section.Height);

            var bmp = new Bitmap(section.Width, section.Height);

            var g = Graphics.FromImage(bmp);

            g.DrawImage(source, 0, 0, section, GraphicsUnit.Pixel);

            return bmp;
        }
        private class Point
        {
            public int X;
            public int Y;
            public int V;
        }


        /// <summary>
        /// Для удаления пустых страниц документа
        /// </summary>
        /// <param name="image"></param>
        /// <returns></returns>
        public static bool DocumentDetectInfo(Bitmap image)
        {
            const double widthProcImage = 200;
            const int sens = 15;
            const int treshold = 25;
            const int widthQuantum = 10;


            var sourceImage = image;
            var sizeFactor = widthProcImage / sourceImage.Width;
            var procBtmp = new Bitmap(sourceImage, (int)Math.Round(sourceImage.Width * sizeFactor), (int)Math.Round(sourceImage.Height * sizeFactor));
            var bd = procBtmp.LockBits(new Rectangle(0, 0, procBtmp.Width, procBtmp.Height), ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);
            var bytes = Math.Abs(bd.Stride) * procBtmp.Height;
            var source = new byte[bytes];

            System.Runtime.InteropServices.Marshal.Copy(bd.Scan0, source, 0, bytes);

            var maxV = 0;

            var size = bd.Width / widthQuantum;

            var hight = 0;
            var low = 0;

            for (var y = 0; y < bd.Height + size; y += size)
            { // x processing
                for (var x = 0; x < bd.Width + size; x += size)
                { // y processing
                    var value = 0;

                    for (var yy = y; (yy < y + size) && (yy < bd.Height); yy++)
                    { // Horosontal counting
                        var pixel = Alignment.GetGrayPixel(source, bd.Width, x, yy);

                        for (var xx = x; (xx < x + size) && (xx < bd.Width); xx++)
                        {
                            var point = Alignment.GetGrayPixel(source, bd.Width, xx, yy);

                            if (Math.Abs(pixel - point) > sens)
                            {
                                value++;
                            }

                            pixel = point;
                        }
                    }

                    for (var xx = x; (xx < x + size) && (xx < bd.Width); xx++)
                    { // Vertical counting
                        var pixel = Alignment.GetGrayPixel(source, bd.Width, xx, y);

                        for (var yy = y; (yy < y + size) && (yy < bd.Height); yy++)
                        {
                            var point = Alignment.GetGrayPixel(source, bd.Width, xx, yy);

                            if (Math.Abs(pixel - point) > sens)
                            {
                                value++;
                            }

                            pixel = point;
                        }
                    }

                    maxV = Math.Max(maxV, value);

                    if (value > treshold)
                    {
                        hight++;
                    }
                    else
                    {
                        low++;
                    }
                }
            }

            double cnt = hight + low;
            hight = (int)Math.Round(hight / cnt * 100);

            procBtmp.UnlockBits(bd);

            return (hight > treshold);
        }

    }
}
