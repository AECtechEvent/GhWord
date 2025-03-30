using Grasshopper;
using Grasshopper.Kernel;
using System;
using System.Drawing;

namespace GhWord
{
    public class GhWordInfo : GH_AssemblyInfo
    {
        public override string Name => "GhWord";

        //Return a 24x24 pixel bitmap to represent this GHA library.
        public override Bitmap Icon => null;

        //Return a short string describing the purpose of this GHA library.
        public override string Description => "Grasshopper Plugin for Microsoft Word interoperability";

        public override Guid Id => new Guid("3249A325-5D8C-4EAD-8456-C1805CFFDCA3");

        //Return a string identifying you or your company.
        public override string AuthorName => "Thornton Tomasetti | CORE studio / AECtech";

        //Return a string representing your preferred contact details.
        public override string AuthorContact => "corestudio@thorntontomasetti.com";
    }
}