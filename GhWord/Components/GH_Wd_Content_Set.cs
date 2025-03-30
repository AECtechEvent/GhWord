using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace GhWord.Components
{
    public class GH_Wd_Content_Set : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Wd_Content_Set class.
        /// </summary>
        public GH_Wd_Content_Set()
          : base("Set Document Content", "WD Set",
              "Sequentially add Contents to Word Document",
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
            pManager.AddGenericParameter("Content", "C", "A list of Word Content Objects to add to the Word Document", GH_ParamAccess.list);
            pManager[1].Optional = false;
            pManager.AddBooleanParameter("Clear", "_X", "If true, the documents current content will be deleted prior to writing the new content.", GH_ParamAccess.item, false);
            pManager[2].Optional = true;
            pManager.AddBooleanParameter("Activate", "_A", "If true, the component will be activated.", GH_ParamAccess.item, false);
            pManager[3].Optional = false;
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
            bool activate = false;
            if (DA.GetData(3, ref activate))
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

                    bool clear = false;
                    DA.GetData(2, ref clear);
                    if (clear) document.Clear();

                    List<IGH_Goo> gooB = new List<IGH_Goo>();
                    if (!DA.GetDataList(1, gooB)) return;

                    foreach (IGH_Goo goo in gooB)
                    {
                        if (goo.CastTo<WdContent>(out WdContent content))
                        {
                            content.AddToDocument(document);
                        }
                    }

                    DA.SetData(0, document);
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
                return Properties.Resources.Icons_Content_Set;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("43673281-EE01-4F97-AA90-45B83CC9D060"); }
        }
    }
}