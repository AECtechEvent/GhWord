using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

using Sd = System.Drawing;

namespace GhWord.Components
{
    public class GH_Wd_Content_Image : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Wd_Content_Image class.
        /// </summary>
        public GH_Wd_Content_Image()
          : base("Create Image Content", "WD Img",
              "Create an Image Content Object",
              "CORE", "Word")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Image", "I", "A System Drawing Bitmap or full FilePath to an image file", GH_ParamAccess.item);
            pManager[0].Optional = false;
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
            IGH_Goo gooA = null;
            if (!DA.GetData(0, ref gooA)) return;

            bool isValid = false;
            if (gooA.CastTo<Sd.Bitmap>(out Sd.Bitmap image)) isValid = true;

            if (!isValid)
            {
                if (gooA.CastTo<string>(out string filepath))
                {
                    if (System.IO.File.Exists(filepath))
                    {
                        image = new Sd.Bitmap(filepath);
                        isValid = true;
                    }
                    else
                    {
                        this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "I input must be a System Drawing Bitmap or a full file path to an image file");
                        return;
                    }
                }
                else
                {
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "I input must be a System Drawing Bitmap or a full file path to an image file");
                    return;
                }
            }

            if (isValid)
            {
                WdContent content = new WdContent(image);
                DA.SetData(0, content);
            }
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
                return Properties.Resources.Icons_Image;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("18D736D2-DA7E-4CD2-A5C1-F956F86C117F"); }
        }
    }
}