using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using WD = Microsoft.Office.Interop.Word;

namespace GhWord
{
    public class WdApp
    {

        #region members

        public WD.Application ComObj = null;

        #endregion

        #region constructors

        public WdApp()
        {
            try
            {
                this.ComObj = (WD.Application)Marshal2.GetActiveObject("Word.Application");
            }
            catch (Exception e)
            {
                this.ComObj = new WD.Application();
            }
            if (!this.ComObj.Visible) this.ComObj.Visible = true;
            this.ComObj.Activate();
        }

        public WdApp(WdApp wdApp)
        {
            this.ComObj = wdApp.ComObj;
        }

        public WdApp(WD.Application comObj)
        {
            this.ComObj = comObj;
        }

        #endregion

        #region properties



        #endregion

        #region methods

        #region -documents

        public WdDocument LoadDocument(string filePath)
        {
            WdDocument document = new WdDocument(this.ComObj.Documents.Open(filePath));

            return document;
        }

        public WdDocument GetActiveDocument()
        {

            if (this.ComObj.Documents.Count < 1)
            {
                //Creates a new workbook if no document(s) are open
                return new WdDocument(this.ComObj.Documents.Add());
            }
            else
            {
                //Gets the topmost workbook if document(s) are open
                return new WdDocument(this.ComObj.ActiveDocument);
            }

        }

        public List<WdDocument> GetAllDocuments()
        {
            List<WdDocument> output = new List<WdDocument>();

            foreach (WD.Document document in this.ComObj.Documents)
            {
                output.Add(new WdDocument(document));
            }

            return output;
        }

        #endregion

        #endregion

        #region overrides

        public override string ToString()
        {
            return "WD | App";
        }

        #endregion

    }
}
