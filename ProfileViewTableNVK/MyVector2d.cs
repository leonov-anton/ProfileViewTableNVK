using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace ProfileViewTableNVK
{
    class MyVector2d
    {
        private double x;
        private double y;

        public double X{
            get{
                return this.x;
            }
            set{
                this.x = value;
            }
        }

        public double Y{
            get{
                return this.y;
            }
            set{
                this.y = value;
            }
        }

        public MyVector2d(double x1, double y1, double x2, double y2)
        {
            this.x = x2 - x1;
            this.y = y2 - y1;
        }

        public static double operator *(MyVector2d vec1, MyVector2d vec2)
        {
            double result = vec1.X * vec2.X + vec1.Y * vec2.Y;
            return result;
        }
        public static double operator /(MyVector2d vec1, MyVector2d vec2)
        {
            double result = vec1.X * vec2.Y - vec1.Y * vec2.X;
            return result;
        }

        public static double Angle(MyVector2d vec1, MyVector2d vec2)
        {
            double result = Math.Asin((vec1 / vec2) / vec1.Lenht() / vec2.Lenht()) * 180 / Math.PI;
            return result;
        }

        public double Lenht()
        {
            double result = Math.Sqrt(Math.Pow(this.X, 2) + Math.Pow(this.Y, 2));
            return result;
        }

    }
}
