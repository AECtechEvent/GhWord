using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace GhWord.Components
{
    public class GH_Wd_Content_Get : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Wd_Content_Get class.
        /// </summary>
        public GH_Wd_Content_Get()
          : base("Get Document Content", "WD Get",
              "Get the Content Objects from a Word Document",
              "CORE", "Word")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Word Document", "Doc", "A Word Document Object", GH_ParamAccess.item);
            pManager[0].Optional = false;
            pManager.AddBooleanParameter("Activate", "_A", "If true, the component will be activated.", GH_ParamAccess.item, false);
            pManager[1].Optional = false;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Content", "Con", "A list of Word Content objects from the Word Document", GH_ParamAccess.list);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool activate = false;
            if (DA.GetData(1, ref activate))
            {
                if (activate)
                {
                    IGH_Goo gooA = null;
                    if (!DA.GetData(0, ref gooA)) return;

                    if (!gooA.CastTo<WdDocument>(out WdDocument document))
                    {
                        this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Doc input must be a Word Document Object");
                        return;
                    }

                    DA.SetDataList(0, document.GetContents());
                }
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
                return Properties.Resources.Icons_Content_Get;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("0670991C-482B-461B-8824-5A13426BC0CA"); }
        }
    }
}