using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace GhWord.Components
{
    public class GH_Wd_Deconstruct_Text : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Wd_Deconstruct_Text class.
        /// </summary>
        public GH_Wd_Deconstruct_Text()
          : base("Deconstruct Text Content", "WD DeTxt",
              "Deconstruct a Text Content Object into its constituent parts",
              "CORE", "Word")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Text Content", "Con", "A Text Word Content object", GH_ParamAccess.item);
            pManager[0].Optional = false;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Text", "T", "The text content as a string", GH_ParamAccess.item);
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
                if (content.ContentType != WdContent.Type.Text)
                {
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Provided Con input is " + content.ContentType.ToString() + " type and must be an Text type");
                    return;
                }
            }

            DA.SetData(0, content.Text);
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
                return Properties.Resources.Icons_DeText;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("1C2DCF54-7A49-45F8-955E-1AE30F726BAC"); }
        }
    }
}