using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace AEIOU
{
    public struct Rect
    {
        public int X { get; set; }
        public int Y { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }

        public Rect(int x = 0, int y = 0, int width = 0, int height = 0)
        {
            X = x;
            Y = y;
            Width = width;
            Height = height;
        }

        public Rectangle Rectangle
        {
            get
            {
                return new Rectangle(X, Y, Width, Height);
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
            get { return X; }
        }
        public int Right
        {
            get { return X + (Width - 1); }
        }
        public int Top
        {
            get{ return Y; }
        }
        public int Bottom
        {
            get { return Y + (Height - 1); }
        }
    }
}
