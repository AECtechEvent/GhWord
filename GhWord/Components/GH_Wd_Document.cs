using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace GhWord.Components
{
    public class GH_Wd_Document : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Wd_Document class.
        /// </summary>
        public GH_Wd_Document()
          : base("Word Document", "WD Doc",
              "Gets or creates an active Microsoft Word Document Object from the Word Application",
              "CORE", "Word")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Word Application", "App", "An Word Application Object", GH_ParamAccess.item);
            pManager[0].Optional = false;
            pManager.AddGenericParameter("Filepath", "*F", "OPTIONAL: The full filepath to a Word Document (File)", GH_ParamAccess.item);
            pManager[1].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Word Document", "Doc", "A Word Document Object", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooA = null;
            if (!DA.GetData(0, ref gooA)) return;

            WdApp app = null;
            if (!gooA.CastTo<WdApp>(out app))
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "App input must be an Word Application Object");
                return;
            }

            string filepath = string.Empty;

            if (DA.GetData(1, ref filepath))
            {
                if (!System.IO.File.Exists(filepath))
                {
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "The specified filepath does not exist");
                    return;
                }
                DA.SetData(0, app.LoadDocument(filepath));
            }
            else
            {
                DA.SetData(0, app.GetActiveDocument());
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
                return Properties.Resources.Icons_Document;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("716E6D11-3FA7-478C-9C7C-9D8759AE529B"); }
        }
    }
}