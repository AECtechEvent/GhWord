using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace GhWord.Components
{
    public class GH_Wd_Deconstruct_Image : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Wd_Deconstruct_Set class.
        /// </summary>
        public GH_Wd_Deconstruct_Image()
          : base("Deconstruct Image Content", "WD DeImg",
              "Deconstruct a Image Content Object into its constituent parts",
              "CORE", "Word")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Image Content", "Con", "An Image Word Content Object", GH_ParamAccess.item);
            pManager[0].Optional = false;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Bitmap", "B", "A System Drawing Bitmap", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooA = null;
            if (!DA.GetData(0, ref gooA)) return;

            WdContent content = null;
            if (!gooA.CastTo<WdContent>(out content))
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Con input must be a Word Content Object");
                return;
            }
            else
            {
                if (content.ContentType != WdContent.Type.Image)
                {
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Provided Con input is " + content.ContentType.ToString() + " type and must be an Image type");
                    return;
                }
            }

            DA.SetData(0, content.Image);
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
                return Properties.Resources.Icons_DeImage;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("79B2AD4D-A4C0-44C7-9F3B-ED7C2AC2AD83"); }
        }
    }
}