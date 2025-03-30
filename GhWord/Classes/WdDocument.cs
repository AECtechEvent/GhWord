using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using WD = Microsoft.Office.Interop.Word;

namespace GhWord
{
    public class WdDocument
    {

        #region members

        public WD.Document ComObj = null;

        #endregion

        #region constructors

        public WdDocument()
        {
        }

        public WdDocument(WD.Document comObj)
        {
            this.ComObj = comObj;
        }

        public WdDocument(WdDocument document)
        {
            this.ComObj = document.ComObj;
        }

        #endregion

        #region properties

        public virtual WdApp Application
        {
            get { return new WdApp(this.ComObj.Application); }
        }

        public virtual string Name
        {
            get { return System.IO.Path.GetFileNameWithoutExtension(this.ComObj.Name); }
        }

        #endregion

        #region methods

        public void Clear()
        {
            this.ComObj.Content.Delete();
        }

        public List<WdContent> GetContents()
        {
            List<WdContent> contents = new List<WdContent>();
            bool isTable = false;
            foreach (WD.Paragraph paragraph in this.ComObj.Paragraphs)
            {
                if (paragraph.Range.Tables.Count > 0)
                {
                    WD.Table table = paragraph.Range.Tables[1];
                    if (!isTable)
                    {
                        List<List<string>> values = new List<List<string>>();
                        for (int i = 0; i < table.Columns.Count; i++)
                        {
                            List<string> vals = new List<string>();
                            for (int j = 0; j < table.Rows.Count; j++)
                            {
                                vals.Add(table.Cell(j + 1, i + 1).Range.Paragraphs[1].Range.Text.Replace(((char)7).ToString(), ""));
                            }
                            values.Add(vals);
                        }
                        contents.Add(new WdContent(values));
                    }
                    isTable = true;
                }
                else
                {
                    isTable = false;
                    string text = paragraph.Range.Text.Trim();
                    if (!string.IsNullOrEmpty(text))
                    {
                        if (text != "/") contents.Add(new WdContent(text, paragraph.Range.Font.Size, paragraph.Range.Font.Name));
                    }
                }

                foreach (WD.InlineShape shape in paragraph.Range.InlineShapes)
                {
                    if (shape.Type == WD.WdInlineShapeType.wdInlineShapePicture)
                    {
                        if (shape.TryGetBitmap(out System.Drawing.Bitmap bitmap)) contents.Add(new WdContent(bitmap));
                    }
                }

                //Check if Paragraph contains image
            }

            return contents;
        }

        #endregion

        #region overrides

        public override string ToString()
        {
            return "Wd | Document {" + this.Name + "}";
        }

        #endregion

    }
}
