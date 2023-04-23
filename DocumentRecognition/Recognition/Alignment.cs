namespace DocumentRecognition
{
    using AForge.Imaging;
    using AForge.Imaging.Filters;
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.Drawing.Imaging;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    public  class Alignment
    {
        public static Bitmap DocumentAngleCorrection(Bitmap image)
        {
            Bitmap newbmp = image.Clone(new Rectangle(0, 0, image.Width, image.Height), System.Drawing.Imaging.PixelFormat.Format24bppRgb);

            var grayImage = Grayscale.CommonAlgorithms.RMY.Apply(newbmp);
            var skewChecker = new DocumentSkewChecker();

            var angle = skewChecker.GetSkewAngle(grayImage);

            while (angle >= 90)
            {
                angle -= 90;
            }

            while (angle <= -90)
            {
                angle += 90;
            }

            var rotator = new RotateBilinear(-angle, false);
            //rotator.FillColor = GetAverageBorderColor(image);
            image = rotator.Apply(newbmp);

            return image;
        }
        /// <summary>
        /// Возвращение среднего значения интенсивности пикселя
        /// </summary>
        /// <param name="bitmap"></param>
        /// <returns></returns>
        private static Color GetAverageBorderColor(Bitmap bitmap)
        {
            var widthProcImage = (double)200;

            var sourceImage = bitmap;
            var sizeFactor = widthProcImage / sourceImage.Width;
            var procBtmp = new Bitmap(sourceImage, (int)Math.Round(sourceImage.Width * sizeFactor), (int)Math.Round(sourceImage.Height * sizeFactor));
            var bitmapData = procBtmp.LockBits(new Rectangle(0, 0, procBtmp.Width, procBtmp.Height), ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);

            var bytes = Math.Abs(bitmapData.Stride) * procBtmp.Height;
            var sourceBytes = new byte[bytes];
            System.Runtime.InteropServices.Marshal.Copy(bitmapData.Scan0, sourceBytes, 0, bytes);

            var channels = new Dictionary<char, int>();
            channels.Add('r', 0);
            channels.Add('g', 0);
            channels.Add('b', 0);

            var cnt = 0;

            for (var y = 0; y < bitmapData.Height; y++)
            { // vertical
                var c = GetColorPixel(sourceBytes, bitmapData.Width, 0, y);
                channels['r'] += c.R;
                channels['g'] += c.G;
                channels['b'] += c.B;
                cnt++;

                c = GetColorPixel(sourceBytes, bitmapData.Width, bitmapData.Width - 1, y);
                channels['r'] += c.R;
                channels['g'] += c.G;
                channels['b'] += c.B;
                cnt++;
            }

            for (var x = 0; x < bitmapData.Width; x++)
            { // horisontal
                var c = GetColorPixel(sourceBytes, bitmapData.Width, x, 0);
                channels['r'] += c.R;
                channels['g'] += c.G;
                channels['b'] += c.B;
                cnt++;

                c = GetColorPixel(sourceBytes, bitmapData.Width, x, bitmapData.Height - 1);
                channels['r'] += c.R;
                channels['g'] += c.G;
                channels['b'] += c.B;
                cnt++;
            }

            procBtmp.UnlockBits(bitmapData);

            var r = (int)Math.Round(((double)channels['r']) / cnt);
            var g = (int)Math.Round(((double)channels['g']) / cnt);
            var b = (int)Math.Round(((double)channels['b']) / cnt);

            var color = Color.FromArgb(r > 255 ? 255 : r, g > 255 ? 255 : g, b > 255 ? 255 : b);

            return color;
        }

        private static Color GetColorPixel(byte[] src, int w, int x, int y)
        {
            var s = GetShift(w, x, y);

            if ((s + 3 > src.Length) || (s < 0))
            {
                return Color.Gray;
            }

            byte r = src[s++];
            byte b = src[s++];
            byte g = src[s];

            var c = Color.FromArgb(r, g, b);

            return c;
        }
        private static int GetShift(int width, int x, int y)
        {
            return y * width * 3 + x * 3;
        }
        public static byte GetGrayPixel(byte[] src, int w, int x, int y)
        {
            var s = Alignment.GetShift(w, x, y);

            if ((s + 3 > src.Length) || (s < 0))
            {
                return 127;
            }

            int b = src[s++];
            b += src[s++];
            b += src[s];
            b = (int)(b / 3.0);
            return (byte)b;
        }
    }

}
