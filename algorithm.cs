using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using CcvLib;
using alg = CcvLib.Algorithm;
using LinqToDB;

namespace peilin
{
    class algorithm
    {
        public static Mat SafeROI(Mat src, Rect Roi, OpenCvSharp.Point2f Center)
        {
            int x = (int)Center.X - Roi.Width / 2 < 0 ? 0 : (int)Center.X - Roi.Width / 2;
            int y = (int)Center.Y - Roi.Height / 2 < 0 ? 0 : (int)Center.Y - Roi.Height / 2;
            int Width = (int)Center.X + Roi.Width / 2 > src.Width ? src.Width - (int)Center.X + Roi.Width / 2 : Roi.Width;
            int Height = (int)Center.Y + Roi.Height / 2 > src.Height ? src.Height - (int)Center.Y + Roi.Height / 2 : Roi.Height;

            return new Mat(src, new Rect(x, y, Width, Height));
        }
        public static Mat SafeROI(Mat src, Rect Roi)
        {
            if (Roi.X < 0)
                Roi.X = 0;
            if (Roi.Y < 0)
                Roi.Y = 0;
            if (Roi.X + Roi.Width > src.Width)
                Roi.Width = src.Width - Roi.X;
            if (Roi.Y + Roi.Height > src.Height)
                Roi.Height = src.Height - Roi.Y;
            if (Roi.X > src.Width)
                Roi.X = src.Width - 1;
            if (Roi.Y > src.Height)
                Roi.Y = src.Height - 1;
            if (Roi.Width < 0)
                Roi.Width = 1;
            if (Roi.Height < 0)
                Roi.Height = 1;
            return new Mat(src, Roi);
        }
    }
}
