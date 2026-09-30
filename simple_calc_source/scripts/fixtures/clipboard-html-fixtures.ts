/**
 * Hand-written clipboard HTML fixtures mirroring what Excel 365, Google Sheets, LibreOffice Calc,
 * Word and an ordinary web page put on the Windows clipboard. Side-effect free so the Node QA
 * (scripts/qa-clipboard-html.ts) and a renderer DOMParser check can share them.
 */

export const EXCEL_365 = `<html xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 15">
<link id=Main-File rel=Main-File
href="file:///C:/Users/test/AppData/Local/Temp/msohtmlclip1/01/clip.htm">
<link rel=File-List
href="file:///C:/Users/test/AppData/Local/Temp/msohtmlclip1/01/clip_filelist.xml">
<style>
<!--table
	{mso-displayed-decimal-separator:"\\.";
	mso-displayed-thousand-separator:"\\,";}
@page
	{margin:.75in .7in .75in .7in;
	mso-header-margin:.3in;
	mso-footer-margin:.3in;}
tr
	{mso-height-source:auto;}
col
	{mso-width-source:auto;}
br
	{mso-data-placement:same-cell;}
td
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:bottom;
	border:none;
	mso-background-source:auto;
	mso-pattern:auto;
	mso-protection:locked visible;
	white-space:nowrap;
	mso-rotate:0;}
.xl65
	{font-weight:700;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	text-align:center;
	vertical-align:middle;
	border:.5pt solid windowtext;
	background:#FFFF00;
	mso-pattern:black none;}
.xl66
	{mso-number-format:"\\0022$\\0022\\#\\,\\#\\#0\\.00";}
.xl67
	{mso-number-format:Percent;}
.xl68
	{mso-number-format:"Short Date";}
.xl69
	{white-space:normal;}
.xl70
	{color:red;
	font-size:14.0pt;
	font-style:italic;
	text-decoration:underline;
	text-underline-style:single;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:"\\@";}
.xl71
	{mso-number-format:"0\\.00_\\)\\;\\[Red\\]\\\\\\(0\\.00\\\\\\)";
	border-top:none;
	border-right:none;
	border-bottom:2.0pt double windowtext;
	border-left:none;}
.font5
	{color:#0070C0;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;}
-->
</style>
</head>

<body link="#0563C1" vlink="#954F72">

<table border=0 cellpadding=0 cellspacing=0 width=320 style='border-collapse:
 collapse;width:240pt'>
<!--StartFragment-->
 <col width=64 style='width:48pt'>
 <col width=128 style='mso-width-source:userset;mso-width-alt:4681;width:96pt'>
 <col width=64 span=2 style='width:48pt'>
 <tr height=20 style='height:15.0pt'>
  <td height=20 class=xl65 width=64 style='height:15.0pt;width:48pt'>Name</td>
  <td class=xl65 width=128 style='border-left:none;width:96pt'>Amount</td>
  <td class=xl65 width=64 style='border-left:none;width:48pt'>Rate</td>
  <td class=xl65 width=64 style='border-left:none;width:48pt'>Due</td>
 </tr>
 <tr height=40 style='height:30.0pt'>
  <td height=40 class=xl69 width=64 style='height:30.0pt;width:48pt'>Alpha<br>
  Beta</td>
  <td class=xl66 align=right>$1,234.50</td>
  <td class=xl67 align=right x:num="0.125">12.50%</td>
  <td class=xl68 align=right>1/15/2024</td>
 </tr>
 <tr height=20 style='height:15.0pt'>
  <td height=20 class=xl70 style='height:15.0pt'>007</td>
  <td class=xl71 align=right x:num="-42.5">(42.50)</td>
  <td colspan=2 style='mso-ignore:colspan'>Mixed <font class="font5">bold blue</font></td>
 </tr>
 <tr height=20 style='height:15.0pt'>
  <td height=20 align=center style='height:15.0pt'>TRUE</td>
  <td align=right x:num="2469" x:fmla="=B2*2">2469</td>
  <td colspan=2 class=xl65 style='border-left:none'>Merged total</td>
 </tr>
<!--EndFragment-->
</table>

</body>

</html>`

const q = (value: unknown) => JSON.stringify(value).replace(/"/g, '&quot;')
export const GOOGLE_SHEETS = `<meta charset='utf-8'><google-sheets-html-origin><style type="text/css"><!--td {border: 1px solid #cccccc;}br {mso-data-placement:same-cell;}--></style><table xmlns="http://www.w3.org/1999/xhtml" cellspacing="0" cellpadding="0" dir="ltr" border="1" style="table-layout:fixed;font-size:10pt;font-family:Arial;width:0px;border-collapse:collapse;border:none" data-sheets-root="1"><colgroup><col width="100"/><col width="120"/><col width="100"/></colgroup><tbody><tr style="height:21px;"><td style="overflow:hidden;padding:2px 3px 2px 3px;vertical-align:bottom;font-weight:bold;background-color:#d9ead3;" data-sheets-value="${q({ 1: 2, 2: 'Item' })}">Item</td><td style="overflow:hidden;padding:2px 3px 2px 3px;vertical-align:bottom;font-weight:bold;text-align:center;" data-sheets-value="${q({ 1: 2, 2: 'Price' })}">Price</td><td style="overflow:hidden;padding:2px 3px 2px 3px;vertical-align:bottom;font-weight:bold;" data-sheets-value="${q({ 1: 2, 2: 'Total' })}">Total</td></tr><tr style="height:21px;"><td style="overflow:hidden;padding:2px 3px 2px 3px;vertical-align:bottom;font-style:italic;color:#ff0000;" data-sheets-value="${q({ 1: 2, 2: 'Widget' })}" data-sheets-note="Check supplier">Widget</td><td style="overflow:hidden;padding:2px 3px 2px 3px;vertical-align:bottom;text-align:right;border-bottom:2px solid #000000;" data-sheets-value="${q({ 1: 3, 3: 19.99 })}" data-sheets-numberformat="${q({ 1: 4, 2: '"$"#,##0.00', 3: 1 })}">$19.99</td><td style="overflow:hidden;padding:2px 3px 2px 3px;vertical-align:bottom;text-align:right;" data-sheets-value="${q({ 1: 3, 3: 39.98 })}" data-sheets-formula="=R[0]C[-1]*2">39.98</td></tr><tr style="height:21px;"><td rowspan="2" style="overflow:hidden;padding:2px 3px 2px 3px;vertical-align:middle;" data-sheets-value="${q({ 1: 2, 2: 'Merged' })}">Merged</td><td style="overflow:hidden;padding:2px 3px 2px 3px;vertical-align:bottom;text-align:right;" data-sheets-value="${q({ 1: 3, 3: 45306 })}" data-sheets-numberformat="${q({ 1: 5, 2: 'yyyy-mm-dd', 3: 1 })}">2024-01-15</td><td style="overflow:hidden;padding:2px 3px 2px 3px;vertical-align:bottom;text-align:center;" data-sheets-value="${q({ 1: 4, 4: 1 })}">TRUE</td></tr><tr style="height:21px;"><td style="overflow:hidden;padding:2px 3px 2px 3px;vertical-align:bottom;text-align:right;" data-sheets-value="${q({ 1: 3, 3: 0.25 })}" data-sheets-numberformat="${q({ 1: 3, 2: '0%', 3: 1 })}" data-sheets-formula="=SUM(R[-2]C[1]:R[-1]C[1])/R2C3">25%</td><td style="overflow:hidden;padding:2px 3px 2px 3px;vertical-align:bottom;" data-sheets-value="${q({ 1: 2, 2: 'Docs' })}" data-sheets-hyperlink="https://example.com/docs"><a class="in-cell-link" href="https://example.com/docs" target="_blank">Docs</a></td></tr></tbody></table>`

export const LIBREOFFICE = `<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
	<META HTTP-EQUIV="CONTENT-TYPE" CONTENT="text/html; charset=utf-8"/>
	<TITLE></TITLE>
	<META NAME="GENERATOR" CONTENT="LibreOffice 7.6.4.1 (Windows)"/>
	<STYLE>
		<!--
		BODY,DIV,TABLE,THEAD,TBODY,TFOOT,TR,TH,TD,P { font-family:"Liberation Sans"; font-size:x-small }
		a.comment-indicator:hover + comment { background:#ffd; position:absolute; display:block; border:1px solid black; padding:0.5em;  }
		a.comment-indicator { background:red; display:inline-block; border:1px solid black; width:0.5em; height:0.5em;  }
		comment { display:none;  }
		-->
	</STYLE>
</HEAD>
<BODY TEXT="#000000">
<TABLE FRAME=VOID CELLSPACING=0 COLS=3 RULES=NONE BORDER=0>
	<COLGROUP><COL WIDTH=86><COL WIDTH=86><COL WIDTH=113></COLGROUP>
	<TBODY>
		<TR>
			<TD WIDTH=86 HEIGHT=17 ALIGN=LEFT BGCOLOR="#FFFF00"><B>Region</B></TD>
			<TD WIDTH=86 ALIGN=CENTER><B><I>Sales</I></B></TD>
			<TD WIDTH=113 ALIGN=LEFT><FONT COLOR="#C9211E">Notes</FONT></TD>
		</TR>
		<TR>
			<TD HEIGHT=17 ALIGN=LEFT>North</TD>
			<TD ALIGN=RIGHT SDVAL="1234.5" SDNUM="1033;0;[$$-409]#,##0.00;[RED]-[$$-409]#,##0.00">$1,234.50</TD>
			<TD ALIGN=LEFT>Line one<BR>line two</TD>
		</TR>
		<TR>
			<TD HEIGHT=17 ALIGN=LEFT>South</TD>
			<TD ALIGN=RIGHT SDVAL="45306" SDNUM="1033;1033;DD/MM/YYYY">15/01/2024</TD>
			<TD ALIGN=CENTER SDVAL="1" SDNUM="1033;0;BOOLEAN">TRUE</TD>
		</TR>
		<TR>
			<TD COLSPAN=2 HEIGHT=17 ALIGN=CENTER VALIGN=MIDDLE><FONT FACE="DejaVu Serif" SIZE=4>Wide</FONT></TD>
			<TD ALIGN=RIGHT SDVAL="0.5" SDNUM="1033;0;0%">50%</TD>
		</TR>
	</TBODY>
</TABLE>
</BODY>
</HTML>`

export const WEB_PAGE = `<html><head><title>Report</title><script>alert('x')</script><style>td { color: green }</style></head><body>
<h2 onclick="steal()">Quarterly report</h2>
<table class="data" onmouseover="evil()">
<thead><tr><th>Product<th>Revenue<th>Growth<th>Launch<th>Link</tr></thead>
<tbody>
<tr><td>Caf&eacute; &amp; Bar<img src="x.png" onerror="evil()"><td>$12,345.67<td>12.5%<td>Mar 5, 2024<td><a href="javascript:alert(1)">bad</a>
<tr><td style="background: rgb(255, 0, 0); color: white; font-weight: bold" onclick="evil()">Hot<script>document.cookie</script><td>(1,500)<td>-3%<td>2024-02-29<td><a href="https://ok.example/x?a=1&amp;b=2">ok</a>
<tr><td>&nbsp;<td>007<td>1.5E3<td>13:45<td align="center">centered
</tbody></table>
<p>Source: internal</p>
</body></html>`

export const WORD = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
<head><meta name=ProgId content=Word.Document><meta name=Generator content="Microsoft Word 15">
<style><!-- p.MsoNormal, li.MsoNormal, div.MsoNormal {margin:0in; font-size:11.0pt; font-family:"Calibri",sans-serif;} --></style></head>
<body lang=EN-US><!--StartFragment-->
<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;border:none'>
 <tr>
  <td width=156 valign=top style='width:117.0pt;border:solid windowtext 1.0pt;padding:0in 5.4pt 0in 5.4pt'>
  <p class=MsoNormal><b><span style='font-size:9.0pt;font-family:"Arial",sans-serif'>Header<o:p></o:p></span></b></p>
  </td>
  <td width=156 valign=top style='width:117.0pt;border:solid windowtext 1.0pt;border-left:none;background:#D9E2F3;padding:0in 5.4pt 0in 5.4pt'>
  <p class=MsoNormal align=right style='text-align:right'>42<o:p></o:p></p>
  </td>
 </tr>
</table>
<p class=MsoNormal><o:p>&nbsp;</o:p></p>
<!--EndFragment--></body></html>`

/** Excel's "XML Spreadsheet" clipboard flavour (SpreadsheetML 2003) for the same kind of copy. */
export const EXCEL_XML_SPREADSHEET = `<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
  <Style ss:ID="s62">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2" ss:Color="#FF0000"/>
   </Borders>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000" ss:Bold="1"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s63">
   <NumberFormat ss:Format="&quot;$&quot;#,##0.00"/>
  </Style>
  <Style ss:ID="s64">
   <NumberFormat ss:Format="Short Date"/>
  </Style>
  <Style ss:ID="s65" ss:Parent="s63">
   <Font ss:FontName="Times New Roman" x:Family="Roman" ss:Size="12" ss:Color="#0070C0" ss:Italic="1" ss:Underline="Double"/>
  </Style>
  <Style ss:ID="s66">
   <Alignment ss:Vertical="Bottom" ss:WrapText="1"/>
  </Style>
 </Styles>
 <Worksheet ss:Name="Budget 2024">
  <Table ss:ExpandedColumnCount="4" ss:ExpandedRowCount="4" x:FullColumns="1"
   x:FullRows="1" ss:DefaultRowHeight="15">
   <Column ss:Width="48"/>
   <Column ss:AutoFitWidth="0" ss:Width="96" ss:Span="1"/>
   <Row ss:AutoFitHeight="0">
    <Cell ss:MergeAcross="1" ss:StyleID="s62"><Data ss:Type="String">Header</Data></Cell>
    <Cell ss:StyleID="s63"><Data ss:Type="Number">1234.5</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="DateTime">2024-01-15T00:00:00.000</Data></Cell>
   </Row>
   <Row ss:Height="30">
    <Cell ss:StyleID="s66"><Data ss:Type="String">Line 1&#10;Line 2</Data></Cell>
    <Cell ss:Formula="=R[-1]C[1]*2"><Data ss:Type="Number">2469</Data></Cell>
    <Cell ss:StyleID="s65" ss:Formula="=SUM(R1C3:R[-1]C)"><Data ss:Type="Number">1234.5</Data></Cell>
    <Cell ss:Formula="=1/0"><Data ss:Type="Error">#DIV/0!</Data></Cell>
   </Row>
   <Row>
    <Cell><Data ss:Type="Boolean">1</Data></Cell>
    <Cell ss:HRef="https://example.com/"><Data ss:Type="String">Link</Data><Comment ss:Author="Ana"><ss:Data xmlns="http://www.w3.org/TR/REC-html40"><B><Font html:Face="Tahoma" x:Family="Swiss" html:Size="9" html:Color="#000000">Ana:</Font></B><Font html:Face="Tahoma" x:Family="Swiss" html:Size="9" html:Color="#000000">&#10;Check this</Font></ss:Data></Comment></Cell>
    <Cell ss:Index="4"><ss:Data ss:Type="String" xmlns="http://www.w3.org/TR/REC-html40"><B>Bold</B><Font html:Color="#FF0000"> red</Font></ss:Data></Cell>
   </Row>
   <Row ss:Index="4">
    <Cell><Data ss:Type="DateTime">1899-12-31T13:30:00.000</Data></Cell>
   </Row>
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <Selected/>
  </WorksheetOptions>
 </Worksheet>
</Workbook>
`
