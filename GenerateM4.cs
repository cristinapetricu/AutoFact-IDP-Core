using System;
using System.IO;
using System.Text;

class Program
{
    static void Main()
    {
        StringBuilder sb = new StringBuilder();
        sb.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
        sb.AppendLine("<Activity mc:Ignorable=\"sap sap2010\" x:Class=\"M4_GenerateInvoice\" xmlns=\"http://schemas.microsoft.com/netfx/2009/xaml/activities\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:s=\"clr-namespace:System;assembly=System.Private.CoreLib\" xmlns:sap=\"http://schemas.microsoft.com/netfx/2009/xaml/activities/presentation\" xmlns:sap2010=\"http://schemas.microsoft.com/netfx/2010/xaml/activities/presentation\" xmlns:scg=\"clr-namespace:System.Collections.Generic;assembly=System.Private.CoreLib\" xmlns:sco=\"clr-namespace:System.Collections.ObjectModel;assembly=System.Private.CoreLib\" xmlns:ue=\"clr-namespace:UiPath.Excel;assembly=UiPath.Excel.Activities\" xmlns:ueab=\"clr-namespace:UiPath.Excel.Activities.Business;assembly=UiPath.Excel.Activities\" xmlns:ui=\"http://schemas.uipath.com/workflow/activities\" xmlns:x=\"http://schemas.microsoft.com/winfx/2006/xaml\">");
        sb.AppendLine("  <x:Members>");
        sb.AppendLine("    <x:Property Name=\"in_Config\" Type=\"InArgument(scg:Dictionary(x:String, x:String))\" />");
        sb.AppendLine("    <x:Property Name=\"out_PdfPath\" Type=\"OutArgument(x:String)\" />");
        sb.AppendLine("    <x:Property Name=\"out_InvoiceNum\" Type=\"OutArgument(x:String)\" />");
        sb.AppendLine("    <x:Property Name=\"out_Success\" Type=\"OutArgument(x:Boolean)\" />");
        sb.AppendLine("  </x:Members>");
        sb.AppendLine("  <Sequence>");
        sb.AppendLine("    <Sequence.Variables>");
        
        string[] vars = {
            "invoiceNumber,String", "dateEmitere,String", "dateScadenta,String",
            "parsedEmitere,s:DateTime", "parsedScadenta,s:DateTime",
            "supplierName,String", "supplierCUI,String", "supplierAddress,String",
            "clientName,String", "clientCUI,String", "clientAddress,String",
            "productDesc,String", "productQtyStr,String", "productPriceStr,String", "productUnit,String",
            "productQty,Double", "productPrice,Double", "tvaRate,Double", "subtotal,Double", "tvaAmount,Double", "totalRON,Double",
            "templatePath,String", "outputExcelPath,String", "outputPdfPath,String", "templateExists,Boolean"
        };
        foreach(var v in vars) {
            var parts = v.Split(',');
            sb.AppendLine($"      <Variable x:TypeArguments=\"{parts[1].Replace("String", "x:String").Replace("Double", "x:Double").Replace("Boolean", "x:Boolean")}\" Name=\"{parts[0]}\" />");
        }
        
        string[] boolVars = {
            "isInvoiceValid", "isDateEmitereValid", "isDateScadentaValid",
            "isSupplierValid", "isSupplierCUIValid", "isSupplierAddressValid",
            "isClientValid", "isClientCUIValid", "isClientAddressValid",
            "isProductValid", "isQtyValid", "isPriceValid", "isUnitValid"
        };
        foreach(var v in boolVars) {
            sb.AppendLine($"      <Variable x:TypeArguments=\"x:Boolean\" Name=\"{v}\" Default=\"False\" />");
        }
        
        string[] intVars = {
            "retryInvoice", "retryEmitere", "retryScadenta",
            "retrySupplier", "retrySupplierCUI", "retrySupplierAddress",
            "retryClient", "retryClientCUI", "retryClientAddress",
            "retryProduct", "retryQty", "retryPrice", "retryUnit"
        };
        foreach(var v in intVars) {
            sb.AppendLine($"      <Variable x:TypeArguments=\"x:Int32\" Name=\"{v}\" Default=\"0\" />");
        }
        
        sb.AppendLine("    </Sequence.Variables>");
        sb.AppendLine("    <TryCatch>");
        sb.AppendLine("      <TryCatch.Try>");
        sb.AppendLine("        <Sequence>");

        void AddBlock(string fieldVar, string isValidVar, string retryVar, string type, string defaultVal, string prompt) {
            sb.AppendLine($"          <DoWhile Condition=\"[Not {isValidVar} And {retryVar} &lt; 5]\">");
            sb.AppendLine("            <Sequence>");
            sb.AppendLine($"              <ui:InputDialog IsPassword=\"False\" Label=\"[{prompt}]\" Title=\"[{fieldVar}]\">");
            sb.AppendLine("                <ui:InputDialog.Result>");
            sb.AppendLine($"                  <OutArgument x:TypeArguments=\"x:String\">[{fieldVar}]</OutArgument>");
            sb.AppendLine("                </ui:InputDialog.Result>");
            sb.AppendLine("              </ui:InputDialog>");
            sb.AppendLine("              <Assign>");
            sb.AppendLine($"                <Assign.To><OutArgument x:TypeArguments=\"x:Int32\">[{retryVar}]</OutArgument></Assign.To>");
            sb.AppendLine($"                <Assign.Value><InArgument x:TypeArguments=\"x:Int32\">[{retryVar} + 1]</InArgument></Assign.Value>");
            sb.AppendLine("              </Assign>");
            
            if (type == "String") {
                sb.AppendLine($"              <If Condition=\"[String.IsNullOrWhiteSpace({fieldVar})]\">");
                sb.AppendLine("                <If.Then>");
                sb.AppendLine("                  <Sequence>");
                sb.AppendLine($"                    <ui:MessageBox Caption=\"Eroare\" Text=\"[&quot;❌ Câmpul nu poate fi gol! Încercare &quot; &amp; {retryVar}.ToString() &amp; &quot; din 5&quot;]\" />");
                sb.AppendLine($"                    <Assign><Assign.To><OutArgument x:TypeArguments=\"x:Boolean\">[{isValidVar}]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=\"x:Boolean\">False</InArgument></Assign.Value></Assign>");
                sb.AppendLine("                  </Sequence>");
                sb.AppendLine("                </If.Then>");
                sb.AppendLine("                <If.Else>");
                sb.AppendLine("                  <Sequence>");
                sb.AppendLine($"                    <Assign><Assign.To><OutArgument x:TypeArguments=\"x:String\">[{fieldVar}]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=\"x:String\">[{fieldVar}.Trim()]</InArgument></Assign.Value></Assign>");
                sb.AppendLine($"                    <Assign><Assign.To><OutArgument x:TypeArguments=\"x:Boolean\">[{isValidVar}]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=\"x:Boolean\">True</InArgument></Assign.Value></Assign>");
                sb.AppendLine("                  </Sequence>");
                sb.AppendLine("                </If.Else>");
                sb.AppendLine("              </If>");
            } else if (type == "Date") {
                string parsedVar = fieldVar == "dateEmitere" ? "parsedEmitere" : "parsedScadenta";
                sb.AppendLine($"              <If Condition=\"[DateTime.TryParseExact(If({fieldVar} Is Nothing, &quot;&quot;, {fieldVar}.Trim()), &quot;dd.MM.yyyy&quot;, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, {parsedVar})]\">");
                sb.AppendLine("                <If.Then>");
                sb.AppendLine("                  <Sequence>");
                sb.AppendLine($"                    <Assign><Assign.To><OutArgument x:TypeArguments=\"x:Boolean\">[{isValidVar}]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=\"x:Boolean\">True</InArgument></Assign.Value></Assign>");
                sb.AppendLine("                  </Sequence>");
                sb.AppendLine("                </If.Then>");
                sb.AppendLine("                <If.Else>");
                sb.AppendLine("                  <Sequence>");
                sb.AppendLine($"                    <ui:MessageBox Caption=\"Eroare\" Text=\"[&quot;❌ Format dată invalid! Folosiți dd.MM.yyyy. Încercare &quot; &amp; {retryVar}.ToString() &amp; &quot; din 5&quot;]\" />");
                sb.AppendLine($"                    <Assign><Assign.To><OutArgument x:TypeArguments=\"x:Boolean\">[{isValidVar}]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=\"x:Boolean\">False</InArgument></Assign.Value></Assign>");
                sb.AppendLine("                  </Sequence>");
                sb.AppendLine("                </If.Else>");
                sb.AppendLine("              </If>");
            } else if (type == "Number") {
                string parsedVar = fieldVar == "productQtyStr" ? "productQty" : "productPrice";
                sb.AppendLine($"              <If Condition=\"[Double.TryParse(If({fieldVar} Is Nothing, &quot;&quot;, {fieldVar}.Trim().Replace(&quot;,&quot;, &quot;.&quot;)), {parsedVar})]\">");
                sb.AppendLine("                <If.Then>");
                sb.AppendLine("                  <Sequence>");
                sb.AppendLine($"                    <Assign><Assign.To><OutArgument x:TypeArguments=\"x:Boolean\">[{isValidVar}]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=\"x:Boolean\">True</InArgument></Assign.Value></Assign>");
                sb.AppendLine("                  </Sequence>");
                sb.AppendLine("                </If.Then>");
                sb.AppendLine("                <If.Else>");
                sb.AppendLine("                  <Sequence>");
                sb.AppendLine($"                    <ui:MessageBox Caption=\"Eroare\" Text=\"[&quot;❌ Valoare numerică invalidă! Încercare &quot; &amp; {retryVar}.ToString() &amp; &quot; din 5&quot;]\" />");
                sb.AppendLine($"                    <Assign><Assign.To><OutArgument x:TypeArguments=\"x:Boolean\">[{isValidVar}]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=\"x:Boolean\">False</InArgument></Assign.Value></Assign>");
                sb.AppendLine("                  </Sequence>");
                sb.AppendLine("                </If.Else>");
                sb.AppendLine("              </If>");
            }
            
            sb.AppendLine("            </Sequence>");
            sb.AppendLine("          </DoWhile>");
            
            sb.AppendLine($"          <If Condition=\"[{retryVar} &gt;= 5 And Not {isValidVar}]\">");
            sb.AppendLine("            <If.Then>");
            sb.AppendLine("              <Sequence>");
            sb.AppendLine("                <Assign>");
            sb.AppendLine($"                  <Assign.To><OutArgument x:TypeArguments=\"x:String\">[{fieldVar}]</OutArgument></Assign.To>");
            sb.AppendLine($"                  <Assign.Value><InArgument x:TypeArguments=\"x:String\">[{defaultVal}]</InArgument></Assign.Value>");
            sb.AppendLine("                </Assign>");
            sb.AppendLine("                <Assign>");
            sb.AppendLine($"                  <Assign.To><OutArgument x:TypeArguments=\"x:Boolean\">[{isValidVar}]</OutArgument></Assign.To>");
            sb.AppendLine("                  <Assign.Value><InArgument x:TypeArguments=\"x:Boolean\">True</InArgument></Assign.Value>");
            sb.AppendLine("                </Assign>");
            if (type == "Number") {
                string parsedVar = fieldVar == "productQtyStr" ? "productQty" : "productPrice";
                sb.AppendLine($"                <Assign><Assign.To><OutArgument x:TypeArguments=\"x:Double\">[{parsedVar}]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=\"x:Double\">[Double.Parse({defaultVal}.Replace(&quot;,&quot;, &quot;.&quot;))]</InArgument></Assign.Value></Assign>");
            }
            sb.AppendLine($"                <ui:MessageBox Caption=\"Atenționare\" Text=\"[&quot;⚠ S-a setat valoarea default pentru {fieldVar}: &quot; &amp; {fieldVar}]\" />");
            sb.AppendLine("              </Sequence>");
            sb.AppendLine("            </If.Then>");
            sb.AppendLine("          </If>");
        }

        AddBlock("invoiceNumber", "isInvoiceValid", "retryInvoice", "String", "&quot;AUTO-&quot; &amp; DateTime.Now.ToString(&quot;yyyyMMddHHmm&quot;)", "&quot;Introduceți numărul facturii&quot;");
        AddBlock("dateEmitere", "isDateEmitereValid", "retryEmitere", "Date", "DateTime.Now.ToString(&quot;dd.MM.yyyy&quot;)", "&quot;Introduceți data emiterii (dd.MM.yyyy)&quot;");
        AddBlock("dateScadenta", "isDateScadentaValid", "retryScadenta", "Date", "DateTime.Now.AddDays(30).ToString(&quot;dd.MM.yyyy&quot;)", "&quot;Introduceți data scadenței (dd.MM.yyyy)&quot;");
        AddBlock("supplierName", "isSupplierValid", "retrySupplier", "String", "&quot;Furnizor Default SRL&quot;", "&quot;Introduceți numele furnizorului&quot;");
        AddBlock("supplierCUI", "isSupplierCUIValid", "retrySupplierCUI", "String", "&quot;RO12345678&quot;", "&quot;Introduceți CUI furnizor&quot;");
        AddBlock("supplierAddress", "isSupplierAddressValid", "retrySupplierAddress", "String", "&quot;Adresa Default Furnizor&quot;", "&quot;Introduceți adresa furnizorului&quot;");
        AddBlock("clientName", "isClientValid", "retryClient", "String", "&quot;Client Default SRL&quot;", "&quot;Introduceți numele clientului&quot;");
        AddBlock("clientCUI", "isClientCUIValid", "retryClientCUI", "String", "&quot;RO87654321&quot;", "&quot;Introduceți CUI client&quot;");
        AddBlock("clientAddress", "isClientAddressValid", "retryClientAddress", "String", "&quot;Adresa Default Client&quot;", "&quot;Introduceți adresa clientului&quot;");
        AddBlock("productDesc", "isProductValid", "retryProduct", "String", "&quot;Produs/Serviciu Default&quot;", "&quot;Introduceți descrierea produsului&quot;");
        AddBlock("productQtyStr", "isQtyValid", "retryQty", "Number", "&quot;1&quot;", "&quot;Introduceți cantitatea&quot;");
        AddBlock("productPriceStr", "isPriceValid", "retryPrice", "Number", "&quot;100&quot;", "&quot;Introduceți prețul unitar&quot;");
        AddBlock("productUnit", "isUnitValid", "retryUnit", "String", "&quot;buc&quot;", "&quot;Introduceți unitatea de măsură&quot;");

        sb.AppendLine("          <Assign><Assign.To><OutArgument x:TypeArguments=\"x:Double\">[tvaRate]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=\"x:Double\">0.19</InArgument></Assign.Value></Assign>");
        sb.AppendLine("          <Assign><Assign.To><OutArgument x:TypeArguments=\"x:Double\">[subtotal]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=\"x:Double\">[productQty * productPrice]</InArgument></Assign.Value></Assign>");
        sb.AppendLine("          <Assign><Assign.To><OutArgument x:TypeArguments=\"x:Double\">[tvaAmount]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=\"x:Double\">[subtotal * tvaRate]</InArgument></Assign.Value></Assign>");
        sb.AppendLine("          <Assign><Assign.To><OutArgument x:TypeArguments=\"x:Double\">[totalRON]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=\"x:Double\">[subtotal + tvaAmount]</InArgument></Assign.Value></Assign>");

        sb.AppendLine("          <Assign><Assign.To><OutArgument x:TypeArguments=\"x:String\">[templatePath]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=\"x:String\">[&quot;Data\\Templates\\InvoiceTemplate.xlsx&quot;]</InArgument></Assign.Value></Assign>");
        sb.AppendLine("          <Assign><Assign.To><OutArgument x:TypeArguments=\"x:String\">[outputExcelPath]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=\"x:String\">[&quot;Data\\Output\\Invoice_&quot; &amp; invoiceNumber &amp; &quot;.xlsx&quot;]</InArgument></Assign.Value></Assign>");
        sb.AppendLine("          <Assign><Assign.To><OutArgument x:TypeArguments=\"x:String\">[outputPdfPath]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=\"x:String\">[&quot;Data\\Output\\Invoice_&quot; &amp; invoiceNumber &amp; &quot;.pdf&quot;]</InArgument></Assign.Value></Assign>");
        
        sb.AppendLine("          <ui:CopyFile ContinueOnError=\"False\" Destination=\"[outputExcelPath]\" Overwrite=\"True\" Path=\"[templatePath]\" />");
        
        sb.AppendLine("          <ueab:ExcelProcessScopeX DisplayName=\"Excel Process Scope\">");
        sb.AppendLine("            <ueab:ExcelProcessScopeX.Body>");
        sb.AppendLine("              <ActivityAction x:TypeArguments=\"ui:IExcelProcess\">");
        sb.AppendLine("                <ActivityAction.Argument>");
        sb.AppendLine("                  <DelegateInArgument x:TypeArguments=\"ui:IExcelProcess\" Name=\"ExcelProcessScopeTag\" />");
        sb.AppendLine("                </ActivityAction.Argument>");
        sb.AppendLine("                <ueab:ExcelApplicationCard Password=\"{x:Null}\" ReadFormatting=\"{x:Null}\" WorkbookPath=\"[outputExcelPath]\">");
        sb.AppendLine("                  <ueab:ExcelApplicationCard.Body>");
        sb.AppendLine("                    <ActivityAction x:TypeArguments=\"ue:IWorkbookQuickHandle\">");
        sb.AppendLine("                      <ActivityAction.Argument>");
        sb.AppendLine("                        <DelegateInArgument x:TypeArguments=\"ue:IWorkbookQuickHandle\" Name=\"Excel\" />");
        sb.AppendLine("                      </ActivityAction.Argument>");
        sb.AppendLine("                      <Sequence>");
        sb.AppendLine("                        <ueab:WriteCellX Cell=\"[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;F3&quot;)]\" Value=\"[invoiceNumber]\" />");
        sb.AppendLine("                        <ueab:WriteCellX Cell=\"[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;F4&quot;)]\" Value=\"[dateEmitere]\" />");
        sb.AppendLine("                        <ueab:WriteCellX Cell=\"[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;F5&quot;)]\" Value=\"[dateScadenta]\" />");
        sb.AppendLine("                        <ueab:WriteCellX Cell=\"[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;B9&quot;)]\" Value=\"[supplierName]\" />");
        sb.AppendLine("                        <ueab:WriteCellX Cell=\"[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;B10&quot;)]\" Value=\"[supplierCUI]\" />");
        sb.AppendLine("                        <ueab:WriteCellX Cell=\"[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;B11&quot;)]\" Value=\"[supplierAddress]\" />");
        sb.AppendLine("                        <ueab:WriteCellX Cell=\"[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;E9&quot;)]\" Value=\"[clientName]\" />");
        sb.AppendLine("                        <ueab:WriteCellX Cell=\"[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;E10&quot;)]\" Value=\"[clientCUI]\" />");
        sb.AppendLine("                        <ueab:WriteCellX Cell=\"[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;E11&quot;)]\" Value=\"[clientAddress]\" />");
        sb.AppendLine("                        <ueab:WriteCellX Cell=\"[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;B16&quot;)]\" Value=\"[productDesc]\" />");
        sb.AppendLine("                        <ueab:WriteCellX Cell=\"[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;D16&quot;)]\" Value=\"[productQty.ToString()]\" />");
        sb.AppendLine("                        <ueab:WriteCellX Cell=\"[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;E16&quot;)]\" Value=\"[productUnit]\" />");
        sb.AppendLine("                        <ueab:WriteCellX Cell=\"[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;F16&quot;)]\" Value=\"[productPrice.ToString(&quot;N2&quot;)]\" />");
        sb.AppendLine("                        <ueab:WriteCellX Cell=\"[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;G16&quot;)]\" Value=\"[subtotal.ToString(&quot;N2&quot;)]\" />");
        sb.AppendLine("                        <ueab:WriteCellX Cell=\"[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;G27&quot;)]\" Value=\"[subtotal.ToString(&quot;N2&quot;)]\" />");
        sb.AppendLine("                        <ueab:WriteCellX Cell=\"[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;G28&quot;)]\" Value=\"[tvaAmount.ToString(&quot;N2&quot;)]\" />");
        sb.AppendLine("                        <ueab:WriteCellX Cell=\"[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;G29&quot;)]\" Value=\"[totalRON.ToString(&quot;N2&quot;)]\" />");
        sb.AppendLine("                        <ueab:SaveExcelFileX Workbook=\"[Excel]\" />");
        sb.AppendLine("                        <ueab:SaveAsPdfX DestinationPdfPath=\"[outputPdfPath]\" SaveQuality=\"StandardQuality\" Workbook=\"[Excel]\" />");
        sb.AppendLine("                      </Sequence>");
        sb.AppendLine("                    </ActivityAction>");
        sb.AppendLine("                  </ueab:ExcelApplicationCard.Body>");
        sb.AppendLine("                </ueab:ExcelApplicationCard>");
        sb.AppendLine("              </ActivityAction>");
        sb.AppendLine("            </ueab:ExcelProcessScopeX.Body>");
        sb.AppendLine("          </ueab:ExcelProcessScopeX>");

        sb.AppendLine("          <Assign><Assign.To><OutArgument x:TypeArguments=\"x:String\">[out_PdfPath]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=\"x:String\">[outputPdfPath]</InArgument></Assign.Value></Assign>");
        sb.AppendLine("          <Assign><Assign.To><OutArgument x:TypeArguments=\"x:String\">[out_InvoiceNum]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=\"x:String\">[invoiceNumber]</InArgument></Assign.Value></Assign>");
        sb.AppendLine("          <Assign><Assign.To><OutArgument x:TypeArguments=\"x:Boolean\">[out_Success]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=\"x:Boolean\">True</InArgument></Assign.Value></Assign>");

        sb.AppendLine("        </Sequence>");
        sb.AppendLine("      </TryCatch.Try>");
        sb.AppendLine("      <TryCatch.Catches>");
        sb.AppendLine("        <Catch x:TypeArguments=\"s:Exception\">");
        sb.AppendLine("          <ActivityAction x:TypeArguments=\"s:Exception\">");
        sb.AppendLine("            <ActivityAction.Argument>");
        sb.AppendLine("              <DelegateInArgument x:TypeArguments=\"s:Exception\" Name=\"exception\" />");
        sb.AppendLine("            </ActivityAction.Argument>");
        sb.AppendLine("            <Sequence>");
        sb.AppendLine("              <Assign><Assign.To><OutArgument x:TypeArguments=\"x:Boolean\">[out_Success]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=\"x:Boolean\">False</InArgument></Assign.Value></Assign>");
        sb.AppendLine("              <ui:LogMessage Level=\"Error\" Message=\"[&quot;EROARE M4 | &quot; &amp; exception.Message]\" />");
        sb.AppendLine("              <ui:MessageBox Caption=\"Eroare neașteptată\" Text=\"[&quot;⚠ Eroare neașteptată: &quot; &amp; exception.Message]\" />");
        sb.AppendLine("            </Sequence>");
        sb.AppendLine("          </ActivityAction>");
        sb.AppendLine("        </Catch>");
        sb.AppendLine("      </TryCatch.Catches>");
        sb.AppendLine("    </TryCatch>");
        sb.AppendLine("  </Sequence>");
        sb.AppendLine("</Activity>");

        File.WriteAllText(@"C:\Users\COPY4ALL\Documents\UiPath\AutoFact_IDP_Core\M4_GenerateInvoice.xaml", sb.ToString());
    }
}