using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Rhino.Geometry;

namespace GhWord.Components
{
    public class GH_Wd_App : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Wd_App class.
        /// </summary>
        public GH_Wd_App()
          : base("Word Application", "WD App",
              "Gets or opens an active Microsoft Word Application Object",
              "CORE", "Word")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("Activate", "_A", "If true, the component will be activated.", GH_ParamAccess.item, false);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Word Application", "App", "Word Application Object", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool active = false;
            if (!DA.GetData(0, ref active)) return;

            if (active)
            {
                WdApp app = new WdApp();
                DA.SetData(0, app);
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
                return Properties.Resources.Icons_App;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("DD3E669E-65D8-4062-9F76-8B8A7F433A81"); }
        }
    }
}