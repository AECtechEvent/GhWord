using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using Sd = System.Drawing;
using System.Windows.Forms;

using WD = Microsoft.Office.Interop.Word;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;

namespace GhWord
{
    public static class Extension
    {

        public static int ToOle(this Sd.Color input)
        {
            return (input.B << 16) | (input.G << 8) | input.R;
        }

        public static bool TryGetBitmap(this WD.InlineShape shape, out Sd.Bitmap bitmap)
        {
            shape.Select();
            shape.Range.CopyAsPicture();
            bitmap = null;

            if (Clipboard.ContainsImage())
            {
                using (Sd.Image img = Clipboard.GetImage())
                {
                    Sd.Bitmap bmp = new Sd.Bitmap(img);
                    if (bmp != null) bitmap = bmp;
                    return true;
                }
            }
            return false;
        }

        public static List<List<string>> Duplicate(this List<List<string>> input)
        {
            List<List<string>> output = new List<List<string>>();
            foreach (List<string> vals in input)
            {
                List<string> outputA = new List<string>();
                foreach (string val in vals)
                {
                    outputA.Add(val);
                }
                output.Add(outputA);
            }
            return output;
        }

        public static GH_Structure<GH_String> ToDataTree(this List<List<string>> input, GH_Path path)
        {
            GH_Structure<GH_String> ghData = new GH_Structure<GH_String>();
            for (int i = 0; i < input.Count; i++)
            {
                for (int j = 0; j < input[i].Count; j++)
                {
                    ghData.Append(new GH_String(input[i][j]), path.AppendElement(i));
                }
            }
            return ghData;
        }

    }
}
