using System;
using System.Linq;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Microsoft.VisualStudio.Tools.Applications;
// ReSharper disable RedundantNameQualifier

namespace SetExcelDocumentProperties
{
    class Program
    {
        static void Main(string[] args)
        {
            string assemblyLocation = "";
            Guid solutionId = new Guid();
            Uri deploymentManifestLocation = null;
            string documentLocation = "";

            for (int i = 0; i <= args.Count() - 1; i++)
            {
                Console.WriteLine(args[i]);
                string[] oArugment = args[i].Split('=');

                switch (oArugment[0])
                {
                    case "/assemblyLocation":
                        assemblyLocation = oArugment[1];
                        break;

                    case "/deploymentManifestLocation":
                        if (!Uri.TryCreate(oArugment[1], UriKind.Absolute, out deploymentManifestLocation))
                        {
                            Console.WriteLine("Error creating URI");
                        }
                        break;

                    case "/documentLocation":
                        documentLocation = oArugment[1];
                        break;

                    case "/solutionID":
                        solutionId = Guid.Parse(oArugment[1]);
                        break;
                }
            }
            try
            {
                ServerDocument.RemoveCustomization(documentLocation);
                string[] nonpublicCachedDataMembers;
                ServerDocument.AddCustomization(documentLocation, assemblyLocation,
                                            solutionId, deploymentManifestLocation,
                                            true, out nonpublicCachedDataMembers);

            }
            catch (System.IO.FileNotFoundException)
            {
                Console.WriteLine("The specified document does not exist.");
            }
            catch (System.IO.IOException)
            {
                Console.WriteLine("The specified document is read-only.");
            }
            catch (InvalidOperationException ex)
            {
                Console.WriteLine("The customization could not be removed.\n" +
                    ex.Message);
            }
            catch (DocumentNotCustomizedException ex)
            {
                Console.WriteLine("The document could not be customized.\n" +
                    ex.Message);
            }
        }
    }
}