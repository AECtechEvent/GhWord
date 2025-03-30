using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Sd = System.Drawing;

using WD = Microsoft.Office.Interop.Word;

namespace GhWord
{
    public class WdContent
    {

        #region members

        public enum Type { Text, Image, Table };
        protected Type type = Type.Text;

        string text = string.Empty;
        double fontSize = 12.0;
        string fontFamily = "Arial";
        Sd.Color fontColor = Sd.Color.Black;

        List<List<string>> values = new List<List<string>>();

        Sd.Bitmap image = new Sd.Bitmap(10, 10);

        #endregion

        #region constructors

        public WdContent(WdContent content)
        {
            this.type = content.type;
            this.text = content.text;
            this.fontSize = content.fontSize;
            this.fontFamily = content.fontFamily;
            this.fontColor = content.fontColor;
            this.image = new Sd.Bitmap(content.image);
            this.values = content.values.Duplicate();
        }

        public WdContent(string text, double size = 12, string family = "Arial")
        {
            this.type = Type.Text;
            this.text = text;
            this.fontSize = size;
            this.fontFamily = family;
        }

        public WdContent(string text, Sd.Color color, double size = 12, string family = "Arial")
        {
            this.type = Type.Text;
            this.text = text;
            this.fontSize = size;
            this.fontFamily = family;
            this.fontColor = color;
        }

        public WdContent(List<List<string>> values)
        {
            this.type = Type.Table;
            this.values = values.Duplicate();
        }

        public WdContent(Sd.Bitmap bitmap)
        {
            this.type = Type.Image;
            this.image = new Sd.Bitmap(bitmap);
        }

        public void AddToDocument(WdDocument document)
        {
                    WD.Paragraph paragraph = document.ComObj.Paragraphs.Add();
            switch (this.type)
            {
                default:
                    paragraph.Range.Text = text;
                    paragraph.Range.Font.Size = (float)this.fontSize;
                    paragraph.Range.Font.Name = this.fontFamily;
                    paragraph.Range.Font.Color = (WD.WdColor)this.fontColor.ToOle();
                    paragraph.Range.InsertParagraphAfter();
                    break;
                case Type.Image:
                    string path = Path.GetTempFileName() + ".png";
                    this.image.Save(path, Sd.Imaging.ImageFormat.Png);

                    WD.Range imageRange = paragraph.Range;

                    document.ComObj.InlineShapes.AddPicture(path, Range: imageRange);
                    File.Delete(path);
                    break;
                case Type.Table:

                    WD.Range tableRange = paragraph.Range;


                    int numCols= values.Count;
                    int numRows = values[0].Count;
                    WD.Table table = document.ComObj.Tables.Add(tableRange, numRows, numCols);

                    table.Borders.Enable = 1;

                    for (int i = 0; i < numRows; i++)
                    {
                        for (int j = 0; j < numCols; j++)
                        {
                            table.Cell(i + 1, j + 1).Range.Text = this.values[j][i];
                        }
                    }

                    break;
            }
        }

        #endregion

        #region properties

        public virtual Type ContentType
        {
            get { return this.type; }
        }

        public virtual string Text
        {
            get { return this.text; }
        }

        public virtual Sd.Bitmap Image
        {
            get
            {
                if (image != null) return new Sd.Bitmap(image);
                return null;
            }
        }

        public virtual List<List<string>> Values
        {
            get { return this.values.Duplicate(); }
        }

        #endregion

        #region methods



        #endregion

        #region overrides

        public override string ToString()
        {
            switch (this.type)
            {
                default:
                    if (this.text.Length < 16) return "Wd | " + this.type + " {" + this.text + "}";
                    return "Wd | " + this.type + " {" + this.text.Substring(0, 15) + "...}";
                case Type.Image:
                    return "Wd | " + this.type + " {" + this.image.Width + "w " + this.image.Height + "h}";
                case Type.Table:
                    return "Wd | " + this.type + " {" + this.values.Count + "c " + this.values[0].Count + "r}";
            }
        }

        #endregion

    }
}
