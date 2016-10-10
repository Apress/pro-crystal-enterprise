The source code to this book is organized into functioning projects. This file is your guide to finding the code you need.

If you open in a project and search on the number of your listing, say, "2-1", it will bring you to a source code comment announcing the starting point for that code. 


Chapter 4 - Listing 4-2 is in the BOXIProxyDemo project


Chapter 5 - Most of the code is found in the BO XI Web Service project except for the exception listed below:
Listing 5-27 is in the BOXIProxyDemo project


Chapter 6 - Most of the code is found in the BO XI Web Service project except for the exceptions listed below:
Listing 6-17 is shown below:
   oReport.ReportPrinterOptions.Copies = 1;
   oReport.ReportPrinterOptions.PrinterName = @"\\servername\printername";
   oReport.ReportPrinterOptions.LandscapeMode = false;
   oReport.ReportPrinterOptions.PrintCollationType =
      CePrintCollateType.cePrintCollateTypeCollated;
   oReport.ReportPrinterOptions.PageSize = CePageSize.cePageSizeLegal;
   oReport.ReportPrinterOptions.FromPage = 1;
   oReport.ReportPrinterOptions.ToPage = 2;

Listings 6-25 and 6-26 are in the DataAccess project
Listings 6-27 and 6-35 are in the Web controls project
Listings 6-36 and 6-40 are in the .NET Providers project


Chapter 7 - Most of the code is in the CrystalReportsViewer project except for the exceptions listed below:

Listing 7-2 is shown below:
   crystalReportViewer1.ViewZoom += new CrystalDecisions.Windows.Forms.ZoomEventHandler(CrystalReportViewer1_ViewZoom);

   void CrystalReportViewer1_ViewZoom(object source,CrystalDecisions.Windows.Forms.ZoomEventArgs e)
   {
      MessageBox.Show("You're about to zoom from " +
         e.CurrentZoomFactor.ToString() +
         "% to " + e.NewZoomFactor.ToString() +"%");
   }

Listing 7-7 is shown below:
   ExportOptions oExportOptions;
   ExcelDataOnlyFormatOptions oExcelDataOnlyFormatOptions;
   DiskFileDestinationOptions oDiskFileDestinationOptions;

   oDiskFileDestinationOptions = new DiskFileDestinationOptions();
   oDiskFileDestinationOptions.DiskFileName = @"c:\temp\mydata.xls";

   oExcelDataOnlyFormatOptions = new ExcelDataOnlyFormatOptions();
   oExcelDataOnlyFormatOptions.ExcelConstantColumnWidth = 20;
   oExcelDataOnlyFormatOptions.ExcelUseConstantColumnWidth = true;
   oExcelDataOnlyFormatOptions.ExportImages = true;
   oExcelDataOnlyFormatOptions.ExportPageHeaderAndPageFooter = true;
   oExcelDataOnlyFormatOptions.MaintainColumnAlignment = true;
   oExcelDataOnlyFormatOptions.MaintainRelativeObjectPosition = true;
   oExcelDataOnlyFormatOptions.ShowGroupOutlines = true;
   oExcelDataOnlyFormatOptions.SimplifyPageHeaders = false;
   oExcelDataOnlyFormatOptions.UseWorksheetFunctionsForSummaries = true;

   oExportOptions = new ExportOptions();
   oExportOptions.ExportFormatOptions = oExcelDataOnlyFormatOptions;
   oExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
   oExportOptions.ExportDestinationOptions = oDiskFileDestinationOptions;
   oExportOptions.ExportFormatType = ExportFormatType.ExcelRecord;

   oOrders.Export(oExportOptions);


Listing 7-8 is shown below:
   ExportOptions oExportOptions;
   PdfFormatOptions oPdfFormatOptions;
   MicrosoftMailDestinationOptions oMicrosoftMailDestinationOptions;

   oMicrosoftMailDestinationOptions = new MicrosoftMailDestinationOptions();
   oMicrosoftMailDestinationOptions.UserName = "myname";
   oMicrosoftMailDestinationOptions.Password = "mypass";
   oMicrosoftMailDestinationOptions.MailToList = "seton.software@verizon.net";
   oMicrosoftMailDestinationOptions.MailCCList = "wendy.p.ganz@verizon.net";
   oMicrosoftMailDestinationOptions.MailSubject = "Here's your report";
   oMicrosoftMailDestinationOptions.MailMessage = "Please review this report and get back to me ASAP";

   oPdfFormatOptions = new PdfFormatOptions();
   oPdfFormatOptions.UsePageRange = false;

   oExportOptions = new ExportOptions();
   oExportOptions.ExportFormatOptions = oPdfFormatOptions;
   oExportOptions.ExportDestinationType = ExportDestinationType.MicrosoftMail;
   oExportOptions.ExportDestinationOptions = oMicrosoftMailDestinationOptions;
   oExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;

   oOrders.Export(oExportOptions);


Listing 7-9 is shown below:
   FieldObject oFieldObject;
   
   //oOrders is the generated class wrapper object for the Orders report
   foreach (ReportObject oReportObject in oOrders.ReportDefinition.ReportObjects)
   {
      if (oReportObject.Kind == ReportObjectKind.FieldObject)
      {
         oFieldObject = ((FieldObject) oReportObject);
         oFieldObject.Color = Color.Aqua;
      }
   }



Listing 7-10 is shown below:
   string szName;
   FieldValueType sFieldValueType;

   foreach (DatabaseFieldDefinition oDatabaseFieldDefinition in oOrders.Database.Tables[0].Fields)
   {
      szName = oDatabaseFieldDefinition.Name;
      sFieldValueType = oDatabaseFieldDefinition.ValueType;
   }




Listing 7-11 is shown below:
   public void SectionSettings(SectionFormat oSectionFormat)
   {
      oSectionFormat.EnableKeepTogether = true;
      oSectionFormat.EnableNewPageAfter = true;
      oSectionFormat.EnablePrintAtBottomOfPage = true;
      oSectionFormat.EnableResetPageNumberAfter = true;
      oSectionFormat.EnableSuppress = true;
      oSectionFormat.EnableSuppressIfBlank = true;
      oSectionFormat.EnableUnderlaySection = true;
   }


Listing 7-14 is shown below:
   if (oOrders.Database.Tables[0].TestConnectivity())
      MessageBox.Show("We're in");
   else
      MessageBox.Show("We're not in");


Listing 7-15 is shown below:
   oOrders.PrintOptions.PaperOrientation = PaperOrientation.Landscape;
   oOrders.PrintOptions.PaperSize = PaperSize.PaperLetter;
   oOrders.PrintOptions.PaperSource = PaperSource.Upper;
   oOrders.PrintOptions.PrinterDuplex = PrinterDuplex.Default;

   Console.Write(oOrders.PrintOptions.PageMargins.leftMargin);
   Console.Write(oOrders.PrintOptions.PageMargins.rightMargin);
   Console.Write(oOrders.PrintOptions.PageMargins.topMargin);
   Console.Write(oOrders.PrintOptions.PageMargins.bottomMargin);


Listing 7-16 is shown below:
   oOrders.SummaryInfo.ReportAuthor = "Carl Ganz, III";
   oOrders.SummaryInfo.ReportComments = "These are comments";
   oOrders.SummaryInfo.ReportSubject = "This is the subject";
   oOrders.SummaryInfo.ReportTitle = "My title";



Chapter 8 - Most of the code is found in the BOXIProxyDemo project except for the exceptions listed below:

Listing 8-6 is shown below:
   ReportDefinition oReportDefinition;
   oReportDefinition = oReportClientDocument.ReportDefinition;

   oSection = oReportDefinition.ReportHeaderArea.Sections[0];   
   oSection = oReportDefinition.ReportFooterArea.Sections[0];
   oSection = oReportDefinition.PageHeaderArea.Sections[0];
   oSection = oReportDefinition.PageFooterArea.Sections[0];

Listing 8-7 is shown below:
   ReportDefinition oReportDefinition;
   oReportDefinition = oReportClientDocument.ReportDefinition;
   oSection = oReportDefinition.get_GroupHeaderArea(0).Sections[0];
   oSection = oReportDefinition.get_GroupFooterArea(0).Sections[0];


Listing 8-8 is shown below:
   oSection = oReportClientDocument.ReportDefinition.DetailArea.Sections[0];
   
   foreach (ReportObject oReportObject in oSection.ReportObjects)
   {
      if (oReportObject.Kind == CrReportObjectKindEnum.crReportObjectKindField)
      oReportObject.Border.BackgroundColor = uint.Parse(ColorTranslator.ToWin32(Color.Gold).ToString());
   }


Listing 8-17 is shown below:
   FieldFormat oFieldFormat;
   oFieldFormat = new FieldFormatClass();
   oFieldFormat.CommonFormat.EnableSystemDefault = false;
   oFieldFormat.CommonFormat.EnableSuppressIfDuplicated = false;


Listings 8-27 though 8-30 are found in the BO XI Web Service project


Listing 8-31 is shown below:
   oReportClientDocument = oReportAppFactory.OpenDocument(int.Parse("1266"), 0);

   ParameterFieldController oParameterFieldController;

   oParameterFieldController = oReportClientDocument.DataDefController.ParameterFieldController;
   oParameterFieldController.SetCurrentValue("", "@Country", "UK");
   oParameterFieldController.SetCurrentValue("", "@ReportsTo", 2);
   
   crystalReportViewer1.ReportSource = oReportClientDocument;


Listing 8-32 is shown below:
   FilterController oFilterController;
   string szFilter;

   oFilterController = oReportClientDocument.DataDefController.GroupFilterController;
   szFilter = "{spc_SalesOrders;1.CompanyName} = 'Around the Horn'";
   oFilterController.SetFormulaText(szFilter);


Listing 8-33 is shown below:
   oFilterController = oReportClientDocument.DataDefController.RecordFilterController;
   szFilter = "{spc_SalesOrders;1.Country} = 'USA'";
   oFilterController.SetFormulaText(szFilter);


Listing 8-36 is shown below:
   ISCRField oSummaryField;
   TopNSort oTopNSort;

   oSummaryField = oReportClientDocument.DataDefController.DataDefinition.SummaryFields[0];

   oTopNSort = new TopNSort();
   oTopNSort.Direction = CrSortDirectionEnum.crSortDirectionTopNOrder;
   oTopNSort.DiscardOthers = true;
   oTopNSort.NIndividualGroups = 3;
   oTopNSort.SortField = oSummaryField;
   
   oReportClientDocument.DataDefController.SortController.Add(-1, oTopNSort);


Listing 8-39 is shown below:
   ReportOptions oReportOptions;

   oReportOptions = new ReportOptions();
   oReportOptions.DisplayGroupContentView = true;
   oReportOptions.EnableAsyncQuery = true;
   oReportOptions.EnablePushDownGroupBy = true;
   oReportOptions.EnableSaveDataWithReport = true;
   oReportOptions.EnableSaveSummariesWithReport = true;
   oReportOptions.EnableSelectDistinctRecords = true;
   oReportOptions.EnableTranslateDOSMemos = true;
   oReportOptions.EnableTranslateDOSStrings = true;
   oReportOptions.EnableUseCaseInsensitiveSQLData = true;
   oReportOptions.EnableUseDummyData = true;
   oReportOptions.EnableUseIndexForSpeed = true;
   oReportOptions.EnableVerifyOnEveryPrint = true;
   oReportOptions.ErrorOnMaxNumOfRecords = true;
   oReportOptions.InitialDataContext;
   oReportOptions.MaxNumOfRecords = 10000;
   oReportOptions.NumOfBrowsingRecords = 10000;
   oReportOptions.NumOfCachedBatches = 10000;
   oReportOptions.PreferredView = CrReportDocumentViewEnum.crReportDocumentReportView;
   oReportOptions.RefreshCEProperties = true;
   oReportOptions.ReportStyle = CrReportStyleEnum.crReportStyleExecutiveLeadingBreak;
   oReportOptions.RowsetBatchSize = 10000;
   oReportOptions.ConvertDateTimeType = CrConvertDateTimeTypeEnum.crConvertDateTimeTypeToDate;

   oReportClientDocument.ModifyReportOptions(oReportOptions);

Listing 8-40 is shown below:
   SummaryInfo oSummaryInfo;

   oSummaryInfo = new SummaryInfo();
   oSummaryInfo.Author = "Carl Ganz, Jr.";
   oSummaryInfo.Keywords = "Financial;Orders";
   oSummaryInfo.Comments = "This is the report which lists " +
      "the order details for each customer";
   oSummaryInfo.Title = "Order Details Report";
   oSummaryInfo.Subject = "Order Details";
   oSummaryInfo.IsSavingWithPreview = true; //save a thumbnail image?
   
   oReportClientDocument.ModifySummaryInfo(oSummaryInfo);


Listing 8-41 is shown below:
   PrintOptions oPrintOptions;

   oPrintOptions = new PrintOptionsClass();
   oPrintOptions.PaperSize = CrPaperSizeEnum.crPaperSizePaperLetter;
   oPrintOptions.PaperSource = CrPaperSourceEnum.crPaperSourceAuto;
   oPrintOptions.PrinterDuplex = CrPrinterDuplexEnum.crPrinterDuplexDefault;


Listing 8-42 is shown below:
   PageMargins oPageMargins;

   oPageMargins = new PageMarginsClass();
   oPageMargins.Left = 770;
   oPageMargins.Right = 770;
   oPageMargins.Top = 1440;
   oPageMargins.Bottom = 1440;
   
   oPrintOptions.PageMargins = oPageMargins;


Listing 8-44 is shown below:
   oReportClientDocument = oReportAppFactory.OpenDocument(int.Parse("1266"), 0);
   oSection = oReportClientDocument.ReportDefinition.PageHeaderArea.Sections[0];
   AddBox(oReportClientDocument, oSection, Color.Black, Color.Yellow,
      CrLineStyleEnum.crLineStyleSingle, 3, 9000, 720, 9500, 1100);
   oReportClientDocument.Save();




Chapter 9 - The code is found in different projects as listed below:

Listing 9-1 is in the BOXIProxyDemo project
Listing 9-2 is in the BO XI Web Service project
Listing 9-3 is in the BOXIProxyDemo project
Listings 9-4 through 9-11 are in the BO XI Web Service project
Listings 9-12 and 9-13 are in the ScheduledReports project
Listings 9-14 through 9-21 are in the BOXIProxyDemo project

Listing 9-22 is shown below:
   DECLARE @Data Varchar(1000)
   
   SET @Data = '1,2'

   SELECT * FROM Employees WHERE EmployeeID IN @Data

Listing 9-23 is shown below:
   DECLARE @Data Varchar(1000)
   SET @Data = '1,2'

   SELECT *
   FROM Employees
   WHERE EmployeeID IN
      (SELECT data
       FROM dbo.fnc_NumericCodes(@Data, ','))

Listing 9-24 is shown below:
   DECLARE @Data Varchar(1000)
   SET @Data = '1,2'

   SELECT data
   FROM dbo.fnc_NumericCodes(@Data, ',')

Listing 9-25 is shown below:
   CREATE FUNCTION dbo.fnc_NumericCodes
   (
   @Items varchar(4000),
   @Delimiter varchar(1)
   )
   RETURNS @DataTable TABLE (data int) AS

   BEGIN

   DECLARE @Pos int
   DECLARE @DataPos int
   DECLARE @DataLen smallint
   DECLARE @Temp varchar(4000)
   DECLARE @DataRemain varchar(4000)
   DECLARE @OneItem varchar(4000)

   SET @DataPos = 1
   SET @DataRemain = ''

   WHILE @DataPos <= DATALENGTH(@Items) / 2
      BEGIN
         SET @DataLen = 4000 - DATALENGTH(@DataRemain) / 2
         SET @Temp = @DataRemain + SUBSTRING(@Items, @DataPos, @DataLen)
         SET @DataPos = @DataPos + @DataLen
         SET @Pos = CHARINDEX(@Delimiter, @Temp)
         WHILE @Pos > 0
            BEGIN
               SET @OneItem = LTRIM(RTRIM(LEFT(@Temp, @Pos - 1)))
               INSERT @DataTable (data) VALUES(@OneItem)
               SET @Temp = SUBSTRING(@Temp, @Pos + 1, LEN(@Temp))
               SET @Pos = CHARINDEX(@Delimiter, @Temp)
            END
            
            SET @DataRemain = @Temp
      END
      
      IF LEN(@Items) = 1
         SET @DataRemain = @Items
         INSERT @DataTable(data) VALUES (LTRIM(RTRIM(@DataRemain)))
         RETURN
   END

Listing 9-26 is shown below:
   'SQL Server
   DECLARE @Data varchar(1000)
   SET @Data = ',1,2,3,'
   SELECT EmployeeID
   FROM Employees
      WHERE CHARINDEX(',' + CONVERT(varchar(10), EmployeeID)
      + ',', @Data) <> 0

   'Oracle
   SELECT EmployeeID
   FROM Employees
   WHERE INSTR(',1,2,3,',','||EmployeeID||',') <> 0

Listings 9-27 through 9-34 are in the BOXICriteriaPage project
Listing 9-35 is in the BOXIServiceMonitor project

Listing 9-50 is shown below:
   oInfoObjects = oInfoStore.NewInfoObjectCollection();
   oPluginManager = oInfoStore.PluginManager;
   oPluginInfo = oPluginManager.GetPluginInfo("Report");
   oInfoObject = oInfoObjects.Add(oPluginInfo);


Listing 9-52 is shown below:
   szSQL = "SELECT TOP 10000 SI_ID, MY_COMMENTS " +
      "FROM CI_INFOOBJECTS " +
      "WHERE SI_KIND = 'CrystalReport'";

   oInfoObjects = oInfoStore.Query(szSQL);
   oStringBuilder = new System.Text.StringBuilder();
   oStringBuilder.Append("<MyData>");

   foreach (InfoObject oInfoObject in oInfoObjects)
   {
      string szData = "<Report SI_ID='{0}'>{2}</Report>";
      oStringBuilder.AppendFormat(szData,
      oInfoObject.Properties["SI_ID"],
      oInfoObject.Properties["MY_COMMENTS"]);
   }

   oStringBuilder.Append("</MyData>");
   return oStringBuilder.ToString();


The remainder of Chapter 9's code is in the BOXIProxyDemo and BO XI Web Service projects

Chapter 10 - Listing 4-2 is in the BOXIProxyDemo and BO XI Web Service projects
Chapter 11 - The code is in the BOUnifiedWebServices project
Chapter 12 - The code is in the ThirdPartyDemo project

