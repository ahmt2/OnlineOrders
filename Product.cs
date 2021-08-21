using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OnlineOrders
{
    class Product : IComparable<Product>
    {
        private readonly static List<string> sizePrecedence = new List<string> {
                                                "YS", "YM", "YL", "YXL",
                                                "2T", "3T", "4T", "5T",
                                                "3-6M", "6-12M", "12-18M", "18-24M",
                                                "NB", "6M", "12M", "18M", "24M",
                                                "XS", "S", "M", "Medium", "L", "XL", "2XL", "3XL", "4XL", "5XL"};

        public Product(string st, string col, string pCode, string sz)
        {
            this.Style = st;
            this.Color = col;
            this.ProductCode = pCode;
            this.Size = sz;
        }

        /// <summary>
        /// Assumes there are at least 4 string elements in the array
        /// </summary>
        /// <param name="objArray"></param>
        public Product(object[] objArray)
        {
            this.Style = (string)objArray[0];
            this.Color = (string)objArray[1];
            this.ProductCode = (string)objArray[2];
            this.Size = (string)objArray[3];
        }

        public string Style { get; set; }
        public string Color { get; set; }
        public string ProductCode { get; set; }
        public string Size { get; set; }

        public int CompareTo(Product other)
        {
            if (other == null)
            {
                return 1;
            }

            if (Style.CompareTo(other.Style) != 0)
            {
                return Style.CompareTo(other.Style);
            }
            else if (Size.CompareTo(other.Size) != 0)
            {
                if(sizePrecedence.IndexOf(Size) == -1 || sizePrecedence.IndexOf(other.Size) == -1)
                {
                    //if the sizes are not comparable, compare alphabetically
                    return Size.CompareTo(other.Size);
                }
                int result = 1;
                if (sizePrecedence.IndexOf(Size) < sizePrecedence.IndexOf(other.Size))
                {
                    result = -1; 
                }
                return result;
            }
            else if (Color.CompareTo(other.Color) != 0)
            {
                return Color.CompareTo(other.Color);
            }
            return 0;
        }

        public static bool operator >(Product operand1, Product operand2)
        {
            if (operand1.Style.CompareTo(operand2.Style) > 0)
            {
                return true;
            }
            else if (operand1.Style.CompareTo(operand2.Style) < 0)
            {
                return false;
            }
            else if (sizePrecedence.IndexOf(operand1.Size) > sizePrecedence.IndexOf(operand2.Size))
            {
                return true;
            }
            else if (sizePrecedence.IndexOf(operand1.Size) < sizePrecedence.IndexOf(operand2.Size))
            {
                return false;
            }
            else if (operand1.Color.CompareTo(operand2.Color) > 0)
            {
                return true;
            }
            else if (operand1.Color.CompareTo(operand2.Color) < 0)
            {
                return false;
            }
            return false;
        }

        public static bool operator <(Product operand1, Product operand2)
        {
            if (operand1.Style.CompareTo(operand2.Style) < 0)
            {
                return true;
            }
            else if (operand1.Style.CompareTo(operand2.Style) > 0)
            {
                return false;
            }
            else if (Product.sizePrecedence.IndexOf(operand1.Size) < Product.sizePrecedence.IndexOf(operand2.Size))
            {
                return true;
            }
            else if (Product.sizePrecedence.IndexOf(operand1.Size) > Product.sizePrecedence.IndexOf(operand2.Size))
            {
                return false;
            }
            else if (operand1.Color.CompareTo(operand2.Color) < 0)
            {
                return true;
            }
            else if (operand1.Color.CompareTo(operand2.Color) > 0)
            {
                return false;
            }
            return false;
        }

        public override string ToString()
        {
            string result = "";
            result += this.Style + " " + this.Color + " " + this.ProductCode + " " + this.Size;
            return result;
        }

        /// <summary>
        /// Assumes there cannot be different products with same product code
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (obj == null) return false;
            Product objAsProduct = obj as Product;
            return ProductCode.Equals(objAsProduct);
        }

        public override int GetHashCode()
        {
            int hashCode = -213559317;
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Style);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Color);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(ProductCode);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Size);
            return hashCode;
        }
    }
}
