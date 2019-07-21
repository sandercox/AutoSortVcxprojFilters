using System;
using System.Linq;
using System.Xml.Linq;

namespace AutoSortVcxprojFilters
{
    class VCXFilterSorter
    {
        public static void Sort(string FullFilePath)
        {
            try
            {
                if (System.IO.File.Exists(FullFilePath))
                {
                    System.Threading.Thread.Sleep(200); // Avoid load fail

                    XDocument xmlDocument = XDocument.Load(FullFilePath);

                    foreach (var itemgroup in xmlDocument.Root.Elements("{http://schemas.microsoft.com/developer/msbuild/2003}ItemGroup"))
                    {
                        itemgroup.ReplaceNodes(from elem in itemgroup.Elements()
                                               orderby elem.Attribute("Include")?.Value
                                               select elem);
                    }

                    xmlDocument.Save(FullFilePath);
                }
            }
            catch(Exception /*e*/)
            {
                // Make sure the user is not interrupted
            }
        }
    }
}
