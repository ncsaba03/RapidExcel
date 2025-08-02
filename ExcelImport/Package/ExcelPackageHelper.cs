using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.VariantTypes;

namespace ExcelImport.Package
{
    internal static class ExcelPackageHelper
    {
        internal static void GenerateWorkbook(SpreadsheetDocument spreadsheet)
        {
            CorePartsHelper.AddCoreParts(spreadsheet);

            var workbookPart = spreadsheet.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            StylePartHelper.AddStylesPart(spreadsheet);
        }

        internal static class CorePartsHelper
        {
            internal static void AddCoreParts(SpreadsheetDocument spreadsheet)
            {
                var extendedFilePropertiesPart = spreadsheet.ExtendedFilePropertiesPart ??
                                                            spreadsheet.AddNewPart<ExtendedFilePropertiesPart>();

                var properties1 = new Properties();
                properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
                Application application1 = new Application();
                application1.Text = "Microsoft Excel";
                DocumentSecurity documentSecurity1 = new DocumentSecurity();
                documentSecurity1.Text = "0";
                ScaleCrop scaleCrop1 = new ScaleCrop();
                scaleCrop1.Text = "false";

                HeadingPairs headingPairs1 = GenerateHeading();

                TitlesOfParts titlesOfParts1 = GenerateTitlesOfParts();

                Company company1 = new Company();
                company1.Text = "";
                LinksUpToDate linksUpToDate1 = new LinksUpToDate();
                linksUpToDate1.Text = "false";
                SharedDocument sharedDocument1 = new SharedDocument();
                sharedDocument1.Text = "false";
                HyperlinksChanged hyperlinksChanged1 = new HyperlinksChanged();
                hyperlinksChanged1.Text = "false";
                ApplicationVersion applicationVersion1 = new ApplicationVersion();
                applicationVersion1.Text = "16.0300";

                properties1.Append(application1);
                properties1.Append(documentSecurity1);
                properties1.Append(scaleCrop1);
                properties1.Append(headingPairs1);
                properties1.Append(titlesOfParts1);
                properties1.Append(company1);
                properties1.Append(linksUpToDate1);
                properties1.Append(sharedDocument1);
                properties1.Append(hyperlinksChanged1);
                properties1.Append(applicationVersion1);

                extendedFilePropertiesPart.Properties = properties1;
            }

            private static TitlesOfParts GenerateTitlesOfParts()
            {
                TitlesOfParts titlesOfParts1 = new TitlesOfParts();
                VTVector vTVector2 = new VTVector() { BaseType = VectorBaseValues.Lpstr, Size = (UInt32Value)1U };
                VTLPSTR vTLPSTR2 = new VTLPSTR();
                vTLPSTR2.Text = "Export";

                vTVector2.Append(vTLPSTR2);

                titlesOfParts1.Append(vTVector2);
                return titlesOfParts1;
            }

            private static HeadingPairs GenerateHeading()
            {
                HeadingPairs headingPairs1 = new HeadingPairs();

                VTVector vTVector1 = new VTVector() { BaseType = VectorBaseValues.Variant, Size = (UInt32Value)2U };

                Variant variant1 = new Variant();
                VTLPSTR vTLPSTR1 = new VTLPSTR();
                vTLPSTR1.Text = "Export";

                variant1.Append(vTLPSTR1);

                Variant variant2 = new Variant();
                VTInt32 vTInt321 = new VTInt32();
                vTInt321.Text = "1";

                variant2.Append(vTInt321);

                vTVector1.Append(variant1);
                vTVector1.Append(variant2);

                headingPairs1.Append(vTVector1);
                return headingPairs1;
            }
        }


    }
}

