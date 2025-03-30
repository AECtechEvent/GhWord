using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Rhino.Geometry;

namespace GhWord.Components
{
    public class GH_Wd_Content_Text : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Wd_Content_Text class.
        /// </summary>
        public GH_Wd_Content_Text()
          : base("Create Text Content", "WD Txt",
              "Create a Text Content Object",
              "CORE", "Word")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Text", "T", "Text Content", GH_ParamAccess.item);
            pManager[0].Optional = false;
            pManager.AddNumberParameter("Size", "*S", "Font Size", GH_ParamAccess.item, 12);
            pManager[1].Optional = true;
            pManager.AddTextParameter("Family", "*F", "Font Family", GH_ParamAccess.item, "Arial");
            pManager[2].Optional = true;
            pManager.AddColourParameter("Color", "*C", "Font Color", GH_ParamAccess.item, System.Drawing.Color.Black);
            pManager[3].Optional = true;

        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Word Content", "Con", "A Word Content Object", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string text = string.Empty;
            if (!DA.GetData(0, ref text)) return;

            double size = 12;
            DA.GetData(1, ref size);

            string family = "Arial";
            DA.GetData(2, ref family);

            System.Drawing.Color color = System.Drawing.Color.Black;
            DA.GetData(3, ref color);

            WdContent content = new WdContent(text, color, size, family);
            DA.SetData(0, content);
        }

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                //You can add image files to your project resources and access them like this:
                // return Resources.IconForThisComponent;
                return Properties.Resources.Icons_Text;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("A31B2DE6-0830-4400-A111-27AE9F0BDF8B"); }
        }
    }
}