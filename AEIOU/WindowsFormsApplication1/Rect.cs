using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace AEIOU
{
    // 範囲情報 クラス
    class Rect
    {
        public Rect()
        {
            X = 0;
            Y = 0;
            Width = 0;
            Height = 0;
        }

        public Rect(Point point, Size size)
        {
            X = point.X;
            Y = point.Y;
            Width = size.Width;
            Height = size.Height;
        }

        public Rect(int x, int y, int width, int height)
        {
            X = x;
            Y = y;
            Width = width;
            Height = height;
        }

        public int X;
        public int Y;
        public int Width;
        public int Height;

        public Rectangle Rectangle
        {
            get
            {
                return new Rectangle(X,Y,Width,Height);
            }
            set
            {
                X = value.X;
                Y = value.Y;
                Width = value.Width;
                Height = value.Height;
            }
        }

        public int Left
        {
            get
            {
                return X;
            }
        }
        public int Right
        {
            get
            {
                return X + (Width - 1);
            }
        }
        public int Top
        {
            get
            {
                return Y;
            }
        }
        public int Bottom
        {
            get
            {
                return Y + (Height - 1);
            }
        }
    }
}
