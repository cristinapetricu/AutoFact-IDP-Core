$fields = @(
    @{ var="invoiceNumber"; valid="isInvoiceValid"; retry="retryInvoice"; type="String"; def='"AUTO-" & DateTime.Now.ToString("yyyyMMddHHmm")'; prompt="Introduceți numărul facturii" },
    @{ var="dateEmitere"; valid="isDateEmitereValid"; retry="retryEmitere"; type="Date"; def='DateTime.Now.ToString("dd.MM.yyyy")'; prompt="Introduceți data emiterii (dd.MM.yyyy)" },
    @{ var="dateScadenta"; valid="isDateScadentaValid"; retry="retryScadenta"; type="Date"; def='DateTime.Now.AddDays(30).ToString("dd.MM.yyyy")'; prompt="Introduceți data scadenței (dd.MM.yyyy)" },
    @{ var="supplierName"; valid="isSupplierValid"; retry="retrySupplier"; type="String"; def='"Furnizor Default SRL"'; prompt="Introduceți numele furnizorului" },
    @{ var="supplierCUI"; valid="isSupplierCUIValid"; retry="retrySupplierCUI"; type="String"; def='"RO12345678"'; prompt="Introduceți CUI furnizor" },
    @{ var="supplierAddress"; valid="isSupplierAddressValid"; retry="retrySupplierAddress"; type="String"; def='"Adresa Default Furnizor"'; prompt="Introduceți adresa furnizorului" },
    @{ var="clientName"; valid="isClientValid"; retry="retryClient"; type="String"; def='"Client Default SRL"'; prompt="Introduceți numele clientului" },
    @{ var="clientCUI"; valid="isClientCUIValid"; retry="retryClientCUI"; type="String"; def='"RO87654321"'; prompt="Introduceți CUI client" },
    @{ var="clientAddress"; valid="isClientAddressValid"; retry="retryClientAddress"; type="String"; def='"Adresa Default Client"'; prompt="Introduceți adresa clientului" },
    @{ var="productDesc"; valid="isProductValid"; retry="retryProduct"; type="String"; def='"Produs/Serviciu Default"'; prompt="Introduceți descrierea produsului" },
    @{ var="productQtyStr"; valid="isQtyValid"; retry="retryQty"; type="Number"; def='"1"'; prompt="Introduceți cantitatea" },
    @{ var="productPriceStr"; valid="isPriceValid"; retry="retryPrice"; type="Number"; def='"100"'; prompt="Introduceți prețul unitar" },
    @{ var="productUnit"; valid="isUnitValid"; retry="retryUnit"; type="String"; def='"buc"'; prompt="Introduceți unitatea de măsură" }
)

$sb = New-Object System.Text.StringBuilder
$sb.AppendLine('<?xml version="1.0" encoding="utf-8"?>') | Out-Null
$sb.AppendLine('<Activity mc:Ignorable="sap sap2010" x:Class="M4_GenerateInvoice" xmlns="http://schemas.microsoft.com/netfx/2009/xaml/activities" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:s="clr-namespace:System;assembly=System.Private.CoreLib" xmlns:sap="http://schemas.microsoft.com/netfx/2009/xaml/activities/presentation" xmlns:sap2010="http://schemas.microsoft.com/netfx/2010/xaml/activities/presentation" xmlns:scg="clr-namespace:System.Collections.Generic;assembly=System.Private.CoreLib" xmlns:sco="clr-namespace:System.Collections.ObjectModel;assembly=System.Private.CoreLib" xmlns:ue="clr-namespace:UiPath.Excel;assembly=UiPath.Excel.Activities" xmlns:ueab="clr-namespace:UiPath.Excel.Activities.Business;assembly=UiPath.Excel.Activities" xmlns:ui="http://schemas.uipath.com/workflow/activities" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">') | Out-Null
$sb.AppendLine('  <x:Members>') | Out-Null
$sb.AppendLine('    <x:Property Name="in_Config" Type="InArgument(scg:Dictionary(x:String, x:String))" />') | Out-Null
$sb.AppendLine('    <x:Property Name="out_PdfPath" Type="OutArgument(x:String)" />') | Out-Null
$sb.AppendLine('    <x:Property Name="out_InvoiceNum" Type="OutArgument(x:String)" />') | Out-Null
$sb.AppendLine('    <x:Property Name="out_Success" Type="OutArgument(x:Boolean)" />') | Out-Null
$sb.AppendLine('  </x:Members>') | Out-Null
$sb.AppendLine('  <Sequence>') | Out-Null
$sb.AppendLine('    <Sequence.Variables>') | Out-Null

$vars = @(
    "invoiceNumber,String", "dateEmitere,String", "dateScadenta,String",
    "parsedEmitere,s:DateTime", "parsedScadenta,s:DateTime",
    "supplierName,String", "supplierCUI,String", "supplierAddress,String",
    "clientName,String", "clientCUI,String", "clientAddress,String",
    "productDesc,String", "productQtyStr,String", "productPriceStr,String", "productUnit,String",
    "productQty,Double", "productPrice,Double", "tvaRate,Double", "subtotal,Double", "tvaAmount,Double", "totalRON,Double",
    "templatePath,String", "outputExcelPath,String", "outputPdfPath,String", "templateExists,Boolean"
)
foreach ($v in $vars) {
    $parts = $v.Split(',')
    $type = $parts[1].Replace("String", "x:String").Replace("Double", "x:Double").Replace("Boolean", "x:Boolean")
    $sb.AppendLine("      <Variable x:TypeArguments=`"$type`" Name=`"$($parts[0])`" />") | Out-Null
}

foreach ($f in $fields) {
    $sb.AppendLine("      <Variable x:TypeArguments=`"x:Boolean`" Name=`"$($f.valid)`" Default=`"False`" />") | Out-Null
    $sb.AppendLine("      <Variable x:TypeArguments=`"x:Int32`" Name=`"$($f.retry)`" Default=`"0`" />") | Out-Null
}

$sb.AppendLine('    </Sequence.Variables>') | Out-Null
$sb.AppendLine('    <TryCatch>') | Out-Null
$sb.AppendLine('      <TryCatch.Try>') | Out-Null
$sb.AppendLine('        <Sequence>') | Out-Null

foreach ($f in $fields) {
    $sb.AppendLine("          <DoWhile Condition=`"[Not $($f.valid) And $($f.retry) &lt; 5]`">") | Out-Null
    $sb.AppendLine("            <Sequence>") | Out-Null
    $sb.AppendLine("              <ui:InputDialog IsPassword=`"False`" Label=`"[$($f.prompt)]`" Title=`"[$($f.var)]`">") | Out-Null
    $sb.AppendLine("                <ui:InputDialog.Result>") | Out-Null
    $sb.AppendLine("                  <OutArgument x:TypeArguments=`"x:String`">[$($f.var)]</OutArgument>") | Out-Null
    $sb.AppendLine("                </ui:InputDialog.Result>") | Out-Null
    $sb.AppendLine("              </ui:InputDialog>") | Out-Null
    $sb.AppendLine("              <Assign>") | Out-Null
    $sb.AppendLine("                <Assign.To><OutArgument x:TypeArguments=`"x:Int32`">[$($f.retry)]</OutArgument></Assign.To>") | Out-Null
    $sb.AppendLine("                <Assign.Value><InArgument x:TypeArguments=`"x:Int32`">[$($f.retry) + 1]</InArgument></Assign.Value>") | Out-Null
    $sb.AppendLine("              </Assign>") | Out-Null
    
    if ($f.type -eq "String") {
        $sb.AppendLine("              <If Condition=`"[String.IsNullOrWhiteSpace($($f.var))]`">") | Out-Null
        $sb.AppendLine("                <If.Then>") | Out-Null
        $sb.AppendLine("                  <Sequence>") | Out-Null
        $sb.AppendLine("                    <ui:MessageBox Caption=`"Eroare`" Text=`"[&quot;❌ Câmpul nu poate fi gol! Încercare &quot; &amp; $($f.retry).ToString() &amp; &quot; din 5&quot;]`" />") | Out-Null
        $sb.AppendLine("                    <Assign><Assign.To><OutArgument x:TypeArguments=`"x:Boolean`">[$($f.valid)]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=`"x:Boolean`">False</InArgument></Assign.Value></Assign>") | Out-Null
        $sb.AppendLine("                  </Sequence>") | Out-Null
        $sb.AppendLine("                </If.Then>") | Out-Null
        $sb.AppendLine("                <If.Else>") | Out-Null
        $sb.AppendLine("                  <Sequence>") | Out-Null
        $sb.AppendLine("                    <Assign><Assign.To><OutArgument x:TypeArguments=`"x:String`">[$($f.var)]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=`"x:String`">[$($f.var).Trim()]</InArgument></Assign.Value></Assign>") | Out-Null
        $sb.AppendLine("                    <Assign><Assign.To><OutArgument x:TypeArguments=`"x:Boolean`">[$($f.valid)]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=`"x:Boolean`">True</InArgument></Assign.Value></Assign>") | Out-Null
        $sb.AppendLine("                  </Sequence>") | Out-Null
        $sb.AppendLine("                </If.Else>") | Out-Null
        $sb.AppendLine("              </If>") | Out-Null
    } elseif ($f.type -eq "Date") {
        $parsedVar = if ($f.var -eq "dateEmitere") { "parsedEmitere" } else { "parsedScadenta" }
        $sb.AppendLine("              <If Condition=`"[DateTime.TryParseExact(If($($f.var) Is Nothing, &quot;&quot;, $($f.var).Trim()), &quot;dd.MM.yyyy&quot;, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, $parsedVar)]`">") | Out-Null
        $sb.AppendLine("                <If.Then>") | Out-Null
        $sb.AppendLine("                  <Sequence>") | Out-Null
        $sb.AppendLine("                    <Assign><Assign.To><OutArgument x:TypeArguments=`"x:Boolean`">[$($f.valid)]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=`"x:Boolean`">True</InArgument></Assign.Value></Assign>") | Out-Null
        $sb.AppendLine("                  </Sequence>") | Out-Null
        $sb.AppendLine("                </If.Then>") | Out-Null
        $sb.AppendLine("                <If.Else>") | Out-Null
        $sb.AppendLine("                  <Sequence>") | Out-Null
        $sb.AppendLine("                    <ui:MessageBox Caption=`"Eroare`" Text=`"[&quot;❌ Format dată invalid! Folosiți dd.MM.yyyy. Încercare &quot; &amp; $($f.retry).ToString() &amp; &quot; din 5&quot;]`" />") | Out-Null
        $sb.AppendLine("                    <Assign><Assign.To><OutArgument x:TypeArguments=`"x:Boolean`">[$($f.valid)]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=`"x:Boolean`">False</InArgument></Assign.Value></Assign>") | Out-Null
        $sb.AppendLine("                  </Sequence>") | Out-Null
        $sb.AppendLine("                </If.Else>") | Out-Null
        $sb.AppendLine("              </If>") | Out-Null
    } elseif ($f.type -eq "Number") {
        $parsedVar = if ($f.var -eq "productQtyStr") { "productQty" } else { "productPrice" }
        $sb.AppendLine("              <If Condition=`"[Double.TryParse(If($($f.var) Is Nothing, &quot;&quot;, $($f.var).Trim().Replace(&quot;,&quot;, &quot;.&quot;)), $parsedVar)]`">") | Out-Null
        $sb.AppendLine("                <If.Then>") | Out-Null
        $sb.AppendLine("                  <Sequence>") | Out-Null
        $sb.AppendLine("                    <Assign><Assign.To><OutArgument x:TypeArguments=`"x:Boolean`">[$($f.valid)]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=`"x:Boolean`">True</InArgument></Assign.Value></Assign>") | Out-Null
        $sb.AppendLine("                  </Sequence>") | Out-Null
        $sb.AppendLine("                </If.Then>") | Out-Null
        $sb.AppendLine("                <If.Else>") | Out-Null
        $sb.AppendLine("                  <Sequence>") | Out-Null
        $sb.AppendLine("                    <ui:MessageBox Caption=`"Eroare`" Text=`"[&quot;❌ Valoare numerică invalidă! Încercare &quot; &amp; $($f.retry).ToString() &amp; &quot; din 5&quot;]`" />") | Out-Null
        $sb.AppendLine("                    <Assign><Assign.To><OutArgument x:TypeArguments=`"x:Boolean`">[$($f.valid)]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=`"x:Boolean`">False</InArgument></Assign.Value></Assign>") | Out-Null
        $sb.AppendLine("                  </Sequence>") | Out-Null
        $sb.AppendLine("                </If.Else>") | Out-Null
        $sb.AppendLine("              </If>") | Out-Null
    }
    
    $sb.AppendLine("            </Sequence>") | Out-Null
    $sb.AppendLine("          </DoWhile>") | Out-Null
    
    $sb.AppendLine("          <If Condition=`"[$($f.retry) &gt;= 5 And Not $($f.valid)]`">") | Out-Null
    $sb.AppendLine("            <If.Then>") | Out-Null
    $sb.AppendLine("              <Sequence>") | Out-Null
    $sb.AppendLine("                <Assign>") | Out-Null
    $sb.AppendLine("                  <Assign.To><OutArgument x:TypeArguments=`"x:String`">[$($f.var)]</OutArgument></Assign.To>") | Out-Null
    $sb.AppendLine("                  <Assign.Value><InArgument x:TypeArguments=`"x:String`">[$($f.def)]</InArgument></Assign.Value>") | Out-Null
    $sb.AppendLine("                </Assign>") | Out-Null
    $sb.AppendLine("                <Assign>") | Out-Null
    $sb.AppendLine("                  <Assign.To><OutArgument x:TypeArguments=`"x:Boolean`">[$($f.valid)]</OutArgument></Assign.To>") | Out-Null
    $sb.AppendLine("                  <Assign.Value><InArgument x:TypeArguments=`"x:Boolean`">True</InArgument></Assign.Value>") | Out-Null
    $sb.AppendLine("                </Assign>") | Out-Null
    if ($f.type -eq "Number") {
        $parsedVar = if ($f.var -eq "productQtyStr") { "productQty" } else { "productPrice" }
        $sb.AppendLine("                <Assign><Assign.To><OutArgument x:TypeArguments=`"x:Double`">[$parsedVar]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=`"x:Double`">[Double.Parse($($f.def).Replace(&quot;,&quot;, &quot;.&quot;))]</InArgument></Assign.Value></Assign>") | Out-Null
    }
    $sb.AppendLine("                <ui:MessageBox Caption=`"Atenționare`" Text=`"[&quot;⚠ S-a setat valoarea default pentru $($f.var): &quot; &amp; $($f.var)]`" />") | Out-Null
    $sb.AppendLine("              </Sequence>") | Out-Null
    $sb.AppendLine("            </If.Then>") | Out-Null
    $sb.AppendLine("            <If.Else><Sequence/></If.Else>") | Out-Null
    $sb.AppendLine("          </If>") | Out-Null
}

$sb.AppendLine('          <Assign><Assign.To><OutArgument x:TypeArguments="x:Double">[tvaRate]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments="x:Double">0.19</InArgument></Assign.Value></Assign>') | Out-Null
$sb.AppendLine('          <Assign><Assign.To><OutArgument x:TypeArguments="x:Double">[subtotal]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments="x:Double">[productQty * productPrice]</InArgument></Assign.Value></Assign>') | Out-Null
$sb.AppendLine('          <Assign><Assign.To><OutArgument x:TypeArguments="x:Double">[tvaAmount]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments="x:Double">[subtotal * tvaRate]</InArgument></Assign.Value></Assign>') | Out-Null
$sb.AppendLine('          <Assign><Assign.To><OutArgument x:TypeArguments="x:Double">[totalRON]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments="x:Double">[subtotal + tvaAmount]</InArgument></Assign.Value></Assign>') | Out-Null

$sb.AppendLine('          <Assign><Assign.To><OutArgument x:TypeArguments="x:String">[templatePath]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments="x:String">["Data\Templates\InvoiceTemplate.xlsx"]</InArgument></Assign.Value></Assign>') | Out-Null
$sb.AppendLine('          <Assign><Assign.To><OutArgument x:TypeArguments="x:String">[outputExcelPath]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments="x:String">["Data\Output\Invoice_" &amp; invoiceNumber &amp; ".xlsx"]</InArgument></Assign.Value></Assign>') | Out-Null
$sb.AppendLine('          <Assign><Assign.To><OutArgument x:TypeArguments="x:String">[outputPdfPath]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments="x:String">["Data\Output\Invoice_" &amp; invoiceNumber &amp; ".pdf"]</InArgument></Assign.Value></Assign>') | Out-Null

$sb.AppendLine('          <ui:CopyFile ContinueOnError="False" Destination="[outputExcelPath]" Overwrite="True" Path="[templatePath]" />') | Out-Null

$sb.AppendLine('          <ueab:ExcelProcessScopeX DisplayName="Excel Process Scope">') | Out-Null
$sb.AppendLine('            <ueab:ExcelProcessScopeX.Body>') | Out-Null
$sb.AppendLine('              <ActivityAction x:TypeArguments="ui:IExcelProcess">') | Out-Null
$sb.AppendLine('                <ActivityAction.Argument>') | Out-Null
$sb.AppendLine('                  <DelegateInArgument x:TypeArguments="ui:IExcelProcess" Name="ExcelProcessScopeTag" />') | Out-Null
$sb.AppendLine('                </ActivityAction.Argument>') | Out-Null
$sb.AppendLine('                <ueab:ExcelApplicationCard Password="{x:Null}" ReadFormatting="{x:Null}" WorkbookPath="[outputExcelPath]">') | Out-Null
$sb.AppendLine('                  <ueab:ExcelApplicationCard.Body>') | Out-Null
$sb.AppendLine('                    <ActivityAction x:TypeArguments="ue:IWorkbookQuickHandle">') | Out-Null
$sb.AppendLine('                      <ActivityAction.Argument>') | Out-Null
$sb.AppendLine('                        <DelegateInArgument x:TypeArguments="ue:IWorkbookQuickHandle" Name="Excel" />') | Out-Null
$sb.AppendLine('                      </ActivityAction.Argument>') | Out-Null
$sb.AppendLine('                      <Sequence>') | Out-Null
$sb.AppendLine('                        <ueab:WriteCellX Cell="[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;F3&quot;)]" Value="[invoiceNumber]" />') | Out-Null
$sb.AppendLine('                        <ueab:WriteCellX Cell="[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;F4&quot;)]" Value="[dateEmitere]" />') | Out-Null
$sb.AppendLine('                        <ueab:WriteCellX Cell="[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;F5&quot;)]" Value="[dateScadenta]" />') | Out-Null
$sb.AppendLine('                        <ueab:WriteCellX Cell="[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;B9&quot;)]" Value="[supplierName]" />') | Out-Null
$sb.AppendLine('                        <ueab:WriteCellX Cell="[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;B10&quot;)]" Value="[supplierCUI]" />') | Out-Null
$sb.AppendLine('                        <ueab:WriteCellX Cell="[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;B11&quot;)]" Value="[supplierAddress]" />') | Out-Null
$sb.AppendLine('                        <ueab:WriteCellX Cell="[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;E9&quot;)]" Value="[clientName]" />') | Out-Null
$sb.AppendLine('                        <ueab:WriteCellX Cell="[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;E10&quot;)]" Value="[clientCUI]" />') | Out-Null
$sb.AppendLine('                        <ueab:WriteCellX Cell="[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;E11&quot;)]" Value="[clientAddress]" />') | Out-Null
$sb.AppendLine('                        <ueab:WriteCellX Cell="[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;B16&quot;)]" Value="[productDesc]" />') | Out-Null
$sb.AppendLine('                        <ueab:WriteCellX Cell="[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;D16&quot;)]" Value="[productQty.ToString()]" />') | Out-Null
$sb.AppendLine('                        <ueab:WriteCellX Cell="[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;E16&quot;)]" Value="[productUnit]" />') | Out-Null
$sb.AppendLine('                        <ueab:WriteCellX Cell="[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;F16&quot;)]" Value="[productPrice.ToString(&quot;N2&quot;)]" />') | Out-Null
$sb.AppendLine('                        <ueab:WriteCellX Cell="[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;G16&quot;)]" Value="[subtotal.ToString(&quot;N2&quot;)]" />') | Out-Null
$sb.AppendLine('                        <ueab:WriteCellX Cell="[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;G27&quot;)]" Value="[subtotal.ToString(&quot;N2&quot;)]" />') | Out-Null
$sb.AppendLine('                        <ueab:WriteCellX Cell="[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;G28&quot;)]" Value="[tvaAmount.ToString(&quot;N2&quot;)]" />') | Out-Null
$sb.AppendLine('                        <ueab:WriteCellX Cell="[Excel.Sheet(&quot;Invoice&quot;).Cell(&quot;G29&quot;)]" Value="[totalRON.ToString(&quot;N2&quot;)]" />') | Out-Null
$sb.AppendLine('                        <ueab:SaveExcelFileX Workbook="[Excel]" />') | Out-Null
$sb.AppendLine('                        <ueab:SaveAsPdfX DestinationPdfPath="[outputPdfPath]" SaveQuality="StandardQuality" Workbook="[Excel]" />') | Out-Null
$sb.AppendLine('                      </Sequence>') | Out-Null
$sb.AppendLine('                    </ActivityAction>') | Out-Null
$sb.AppendLine('                  </ueab:ExcelApplicationCard.Body>') | Out-Null
$sb.AppendLine('                </ueab:ExcelApplicationCard>') | Out-Null
$sb.AppendLine('              </ActivityAction>') | Out-Null
$sb.AppendLine('            </ueab:ExcelProcessScopeX.Body>') | Out-Null
$sb.AppendLine('          </ueab:ExcelProcessScopeX>') | Out-Null

$sb.AppendLine('          <Assign><Assign.To><OutArgument x:TypeArguments="x:String">[out_PdfPath]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments="x:String">[outputPdfPath]</InArgument></Assign.Value></Assign>') | Out-Null
$sb.AppendLine('          <Assign><Assign.To><OutArgument x:TypeArguments="x:String">[out_InvoiceNum]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments="x:String">[invoiceNumber]</InArgument></Assign.Value></Assign>') | Out-Null
$sb.AppendLine('          <Assign><Assign.To><OutArgument x:TypeArguments="x:Boolean">[out_Success]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments="x:Boolean">True</InArgument></Assign.Value></Assign>') | Out-Null

$sb.AppendLine('        </Sequence>') | Out-Null
$sb.AppendLine('      </TryCatch.Try>') | Out-Null
$sb.AppendLine('      <TryCatch.Catches>') | Out-Null
$sb.AppendLine('        <Catch x:TypeArguments="s:Exception">') | Out-Null
$sb.AppendLine('          <ActivityAction x:TypeArguments="s:Exception">') | Out-Null
$sb.AppendLine('            <ActivityAction.Argument>') | Out-Null
$sb.AppendLine('              <DelegateInArgument x:TypeArguments="s:Exception" Name="exception" />') | Out-Null
$sb.AppendLine('            </ActivityAction.Argument>') | Out-Null
$sb.AppendLine('            <Sequence>') | Out-Null
$sb.AppendLine('              <Assign><Assign.To><OutArgument x:TypeArguments="x:Boolean">[out_Success]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments="x:Boolean">False</InArgument></Assign.Value></Assign>') | Out-Null
$sb.AppendLine('              <ui:LogMessage Level="Error" Message="[&quot;EROARE M4 | &quot; &amp; exception.Message]" />') | Out-Null
$sb.AppendLine('              <ui:MessageBox Caption="Eroare neașteptată" Text="[&quot;⚠ Eroare neașteptată: &quot; &amp; exception.Message]" />') | Out-Null
$sb.AppendLine('            </Sequence>') | Out-Null
$sb.AppendLine('          </ActivityAction>') | Out-Null
$sb.AppendLine('        </Catch>') | Out-Null
$sb.AppendLine('      </TryCatch.Catches>') | Out-Null
$sb.AppendLine('    </TryCatch>') | Out-Null
$sb.AppendLine('  </Sequence>') | Out-Null
$sb.AppendLine('</Activity>') | Out-Null

Set-Content -Path "C:\Users\COPY4ALL\Documents\UiPath\AutoFact_IDP_Core\M4_GenerateInvoice.xaml" -Value $sb.ToString() -Encoding UTF8