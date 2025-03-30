using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace GhWord.Components
{
    public class GH_Wd_Content_Table : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Wd_Content_Table class.
        /// </summary>
        public GH_Wd_Content_Table()
          : base("Create Table Content", "WD Tbl",
              "Create a Table Content Object",
              "CORE", "Word")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Values", "V", "A datatree of text values", GH_ParamAccess.tree);
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
            List<List<string>> dataSet = new List<List<string>>();
            if (!DA.GetDataTree(0, out GH_Structure<GH_String> ghData)) return;

            foreach (List<GH_String> data in ghData.Branches)
            {
                List<string> values = new List<string>();
                foreach (GH_String value in data)
                {
                    values.Add(value.Value);
                }
                dataSet.Add(values);
            }

            WdContent content = new WdContent(dataSet);

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
                return Properties.Resources.Icons_Table;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("CD37DC83-713E-43E1-BE97-C876399C1E19"); }
        }
    }
}