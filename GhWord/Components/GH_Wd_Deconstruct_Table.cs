using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace GhWord.Components
{
    public class GH_Wd_Deconstruct_Table : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Wd_Deconstruct_Table class.
        /// </summary>
        public GH_Wd_Deconstruct_Table()
          : base("Deconstruct Table Content", "WD DeTbl",
              "Deconstruct a Table Content Object into its constituent parts",
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
            pManager.AddTextParameter("Data", "D", "A datatree of values. Each branch of the datatree will be used as a column and each item in the list will be populated to the rows in the column", GH_ParamAccess.tree);
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
                if (content.ContentType != WdContent.Type.Table)
                {
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Provided Con input is " + content.ContentType.ToString() + " type and must be an Table type");
                    return;
                }
            }

            GH_Path path = new GH_Path();

            if (this.Params.Input[0].VolatileData.PathCount > 1) path = this.Params.Input[0].VolatileData.get_Path(this.RunCount - 1);
            path = path.AppendElement(this.RunCount - 1);

            GH_Structure<GH_String> data = content.Values.ToDataTree(path);

            DA.SetDataTree(0, data);
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
                return Properties.Resources.Icons_DeTable;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("FFF80D13-07B3-4893-8430-77F9B15AB5B0"); }
        }
    }
}