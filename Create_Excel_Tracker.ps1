Add-Type -Assembly "System.IO.Compression"
Add-Type -Assembly "System.IO.Compression.FileSystem"

$baseDir = "C:\Users\chris\Documents\Claude kurs"
$outFile = Join-Path $baseDir "Barcelona_Pavilion_2026_Tracker.xlsx"
$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("xlsx_" + [System.Guid]::NewGuid().ToString("N"))

function Write-Xml($path, $content) {
    $enc = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($path, $content, $enc)
}

foreach ($d in @("","_rels","xl","xl\_rels","xl\worksheets")) {
    New-Item -ItemType Directory -Path (Join-Path $tempDir $d) -Force | Out-Null
}

Write-Xml (Join-Path $tempDir "[Content_Types].xml") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>
'@

Write-Xml (Join-Path $tempDir "_rels\.rels") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'@

Write-Xml (Join-Path $tempDir "xl\workbook.xml") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews><workbookView activeTab="0"/></bookViews>
  <sheets>
    <sheet name="Company Pipeline" sheetId="1" r:id="rId1"/>
    <sheet name="Budget Dashboard" sheetId="2" r:id="rId2"/>
    <sheet name="Follow-Up Log" sheetId="3" r:id="rId3"/>
  </sheets>
</workbook>
'@

Write-Xml (Join-Path $tempDir "xl\_rels\workbook.xml.rels") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>
'@

Write-Xml (Join-Path $tempDir "xl\sharedStrings.xml") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>
'@

# --- STYLES ---
# Fonts: 0=default, 1=white bold (header), 2=navy bold (section heading), 3=bold black (budget label)
# Fills: 0=none, 1=gray125, 2=navy (header), 3=light blue-gray (section heading)
# CellXfs: 0=default, 1=header, 2=date, 3=number/currency, 4=percent, 5=section heading, 6=bold label, 7=bold currency, 8=centered, 9=status centered+border
# dxfs: 0=To contact(gray), 1=Contacted no reply(yellow), 2=In dialogue(blue), 3=Meeting booked(orange),
#        4=Proposal sent(light green), 5=Contract sent(medium blue), 6=SIGNED & PAID(green+white),
#        7=Declined / No(red+white)
Write-Xml (Join-Path $tempDir "xl\styles.xml") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="2">
    <numFmt numFmtId="164" formatCode="#,##0"/>
    <numFmt numFmtId="165" formatCode="YYYY\-MM\-DD"/>
  </numFmts>
  <fonts count="4">
    <font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FF1F3864"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFB8B8B8"/></left>
      <right style="thin"><color rgb="FFB8B8B8"/></right>
      <top style="thin"><color rgb="FFB8B8B8"/></top>
      <bottom style="thin"><color rgb="FFB8B8B8"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="10">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1">
      <alignment horizontal="right"/>
    </xf>
    <xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1">
      <alignment horizontal="right"/>
    </xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="164" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1" applyAlignment="1">
      <alignment horizontal="right"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
      <alignment horizontal="center"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center"/>
    </xf>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="8">
    <dxf><fill><patternFill patternType="solid"><fgColor rgb="FFE8E8E8"/></patternFill></fill></dxf>
    <dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/></patternFill></fill></dxf>
    <dxf><fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/></patternFill></fill></dxf>
    <dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFCE4D6"/></patternFill></fill></dxf>
    <dxf><fill><patternFill patternType="solid"><fgColor rgb="FFE2EFDA"/></patternFill></fill></dxf>
    <dxf><fill><patternFill patternType="solid"><fgColor rgb="FFBDD7EE"/></patternFill></fill></dxf>
    <dxf>
      <font><b/><color rgb="FFFFFFFF"/></font>
      <fill><patternFill patternType="solid"><fgColor rgb="FF70AD47"/></patternFill></fill>
    </dxf>
    <dxf>
      <font><b/><color rgb="FFFFFFFF"/></font>
      <fill><patternFill patternType="solid"><fgColor rgb="FFC00000"/></patternFill></fill>
    </dxf>
  </dxfs>
</styleSheet>
'@

# --- SHEET 1: Company Pipeline ---
# Columns: A=Company, B=Contact, C=Email, D=Phone, E=Participant Type, F=Size/Tier,
#           G=Price (NOK), H=Status, I=Date First Contacted, J=Last Contact, K=Next Follow-Up,
#           L=Notes/Objections, M=Contract Sent, N=Contract Signed, O=Invoice Sent, P=Payment Received
Write-Xml (Join-Path $tempDir "xl\worksheets\sheet1.xml") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
      <selection pane="bottomLeft" activeCell="A2" sqref="A2"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>
    <col min="1" max="1" width="26" customWidth="1"/>
    <col min="2" max="2" width="20" customWidth="1"/>
    <col min="3" max="3" width="28" customWidth="1"/>
    <col min="4" max="4" width="16" customWidth="1"/>
    <col min="5" max="5" width="18" customWidth="1"/>
    <col min="6" max="6" width="12" customWidth="1"/>
    <col min="7" max="7" width="16" customWidth="1"/>
    <col min="8" max="8" width="22" customWidth="1"/>
    <col min="9" max="9" width="20" customWidth="1"/>
    <col min="10" max="10" width="18" customWidth="1"/>
    <col min="11" max="11" width="20" customWidth="1"/>
    <col min="12" max="12" width="42" customWidth="1"/>
    <col min="13" max="16" width="15" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1" ht="36" customHeight="1">
      <c r="A1" t="inlineStr" s="1"><is><t>Company Name</t></is></c>
      <c r="B1" t="inlineStr" s="1"><is><t>Contact Name</t></is></c>
      <c r="C1" t="inlineStr" s="1"><is><t>Email</t></is></c>
      <c r="D1" t="inlineStr" s="1"><is><t>Phone</t></is></c>
      <c r="E1" t="inlineStr" s="1"><is><t>Participant Type</t></is></c>
      <c r="F1" t="inlineStr" s="1"><is><t>Booth Size (S/M/L)</t></is></c>
      <c r="G1" t="inlineStr" s="1"><is><t>Price Offered (NOK)</t></is></c>
      <c r="H1" t="inlineStr" s="1"><is><t>Status</t></is></c>
      <c r="I1" t="inlineStr" s="1"><is><t>Date First Contacted</t></is></c>
      <c r="J1" t="inlineStr" s="1"><is><t>Last Contact Date</t></is></c>
      <c r="K1" t="inlineStr" s="1"><is><t>Next Follow-Up Date</t></is></c>
      <c r="L1" t="inlineStr" s="1"><is><t>Notes / Objections</t></is></c>
      <c r="M1" t="inlineStr" s="1"><is><t>Contract Sent?</t></is></c>
      <c r="N1" t="inlineStr" s="1"><is><t>Contract Signed?</t></is></c>
      <c r="O1" t="inlineStr" s="1"><is><t>Invoice Sent?</t></is></c>
      <c r="P1" t="inlineStr" s="1"><is><t>Payment Received?</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>Havfisk AS</t></is></c>
      <c r="B2" t="inlineStr"><is><t>Erik Nilsen</t></is></c>
      <c r="C2" t="inlineStr"><is><t>erik@havfisk.no</t></is></c>
      <c r="D2" t="inlineStr"><is><t>+47 900 12345</t></is></c>
      <c r="E2" t="inlineStr" s="8"><is><t>Exhibitor</t></is></c>
      <c r="F2" t="inlineStr" s="8"><is><t>L</t></is></c>
      <c r="G2" s="3"><v>75000</v></c>
      <c r="H2" t="inlineStr" s="9"><is><t>SIGNED &amp; PAID</t></is></c>
      <c r="I2" t="inlineStr"><is><t>2026-01-15</t></is></c>
      <c r="J2" t="inlineStr"><is><t>2026-02-20</t></is></c>
      <c r="K2" t="inlineStr"><is><t></t></is></c>
      <c r="L2" t="inlineStr"><is><t>Requested extra branding space - resolved</t></is></c>
      <c r="M2" t="inlineStr" s="8"><is><t>Y</t></is></c>
      <c r="N2" t="inlineStr" s="8"><is><t>Y</t></is></c>
      <c r="O2" t="inlineStr" s="8"><is><t>Y</t></is></c>
      <c r="P2" t="inlineStr" s="8"><is><t>Y</t></is></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>Nordic Seafood Solutions</t></is></c>
      <c r="B3" t="inlineStr"><is><t>Ingrid Larsen</t></is></c>
      <c r="C3" t="inlineStr"><is><t>ingrid@nordicsea.no</t></is></c>
      <c r="D3" t="inlineStr"><is><t>+47 930 56789</t></is></c>
      <c r="E3" t="inlineStr" s="8"><is><t>Exhibitor</t></is></c>
      <c r="F3" t="inlineStr" s="8"><is><t>M</t></is></c>
      <c r="G3" s="3"><v>55000</v></c>
      <c r="H3" t="inlineStr" s="9"><is><t>In dialogue</t></is></c>
      <c r="I3" t="inlineStr"><is><t>2026-02-01</t></is></c>
      <c r="J3" t="inlineStr"><is><t>2026-03-04</t></is></c>
      <c r="K3" t="inlineStr"><is><t>2026-03-10</t></is></c>
      <c r="L3" t="inlineStr"><is><t>Wants M booth at L price - needs approval</t></is></c>
      <c r="M3" t="inlineStr" s="8"><is><t>N</t></is></c>
      <c r="N3" t="inlineStr" s="8"><is><t>N</t></is></c>
      <c r="O3" t="inlineStr" s="8"><is><t>N</t></is></c>
      <c r="P3" t="inlineStr" s="8"><is><t>N</t></is></c>
    </row>
    <row r="4">
      <c r="A4" t="inlineStr"><is><t>AquaNor Export AS</t></is></c>
      <c r="B4" t="inlineStr"><is><t>Tor Bergstrom</t></is></c>
      <c r="C4" t="inlineStr"><is><t>tor.b@aquanor.no</t></is></c>
      <c r="D4" t="inlineStr"><is><t>+47 952 34567</t></is></c>
      <c r="E4" t="inlineStr" s="8"><is><t>Exhibitor</t></is></c>
      <c r="F4" t="inlineStr" s="8"><is><t>S</t></is></c>
      <c r="G4" s="3"><v>35000</v></c>
      <c r="H4" t="inlineStr" s="9"><is><t>Proposal sent</t></is></c>
      <c r="I4" t="inlineStr"><is><t>2026-02-10</t></is></c>
      <c r="J4" t="inlineStr"><is><t>2026-03-01</t></is></c>
      <c r="K4" t="inlineStr"><is><t>2026-03-05</t></is></c>
      <c r="L4" t="inlineStr"><is><t>Price-sensitive - mention early bird deadline</t></is></c>
      <c r="M4" t="inlineStr" s="8"><is><t>N</t></is></c>
      <c r="N4" t="inlineStr" s="8"><is><t>N</t></is></c>
      <c r="O4" t="inlineStr" s="8"><is><t>N</t></is></c>
      <c r="P4" t="inlineStr" s="8"><is><t>N</t></is></c>
    </row>
    <row r="5">
      <c r="A5" t="inlineStr"><is><t>Fjord Tech Marine</t></is></c>
      <c r="B5" t="inlineStr"><is><t>Silje Andersen</t></is></c>
      <c r="C5" t="inlineStr"><is><t>silje@fjordtech.no</t></is></c>
      <c r="D5" t="inlineStr"><is><t>+47 910 99001</t></is></c>
      <c r="E5" t="inlineStr" s="8"><is><t>Exhibitor</t></is></c>
      <c r="F5" t="inlineStr" s="8"><is><t>M</t></is></c>
      <c r="G5" s="3"><v>55000</v></c>
      <c r="H5" t="inlineStr" s="9"><is><t>Contacted - no reply</t></is></c>
      <c r="I5" t="inlineStr"><is><t>2026-02-25</t></is></c>
      <c r="J5" t="inlineStr"><is><t>2026-02-25</t></is></c>
      <c r="K5" t="inlineStr"><is><t>2026-03-02</t></is></c>
      <c r="L5" t="inlineStr"><is><t>Send follow-up email</t></is></c>
      <c r="M5" t="inlineStr" s="8"><is><t>N</t></is></c>
      <c r="N5" t="inlineStr" s="8"><is><t>N</t></is></c>
      <c r="O5" t="inlineStr" s="8"><is><t>N</t></is></c>
      <c r="P5" t="inlineStr" s="8"><is><t>N</t></is></c>
    </row>
    <row r="6">
      <c r="A6" t="inlineStr"><is><t>Norsk Laks Eksport</t></is></c>
      <c r="B6" t="inlineStr"><is><t>Bjorn Haugen</t></is></c>
      <c r="C6" t="inlineStr"><is><t>bjorn@norsk-laks.no</t></is></c>
      <c r="D6" t="inlineStr"><is><t>+47 940 11223</t></is></c>
      <c r="E6" t="inlineStr" s="8"><is><t>Network Partner</t></is></c>
      <c r="F6" t="inlineStr" s="8"><is><t>-</t></is></c>
      <c r="G6" s="3"><v>15000</v></c>
      <c r="H6" t="inlineStr" s="9"><is><t>Contract sent</t></is></c>
      <c r="I6" t="inlineStr"><is><t>2026-01-20</t></is></c>
      <c r="J6" t="inlineStr"><is><t>2026-03-03</t></is></c>
      <c r="K6" t="inlineStr"><is><t>2026-03-06</t></is></c>
      <c r="L6" t="inlineStr"><is><t>Awaiting signature - called today, signing this week</t></is></c>
      <c r="M6" t="inlineStr" s="8"><is><t>Y</t></is></c>
      <c r="N6" t="inlineStr" s="8"><is><t>N</t></is></c>
      <c r="O6" t="inlineStr" s="8"><is><t>N</t></is></c>
      <c r="P6" t="inlineStr" s="8"><is><t>N</t></is></c>
    </row>
    <row r="7">
      <c r="A7" t="inlineStr"><is><t>Bergen Pelagic AS</t></is></c>
      <c r="B7" t="inlineStr"><is><t>Lars Svensson</t></is></c>
      <c r="C7" t="inlineStr"><is><t>lars@bergenpelagic.no</t></is></c>
      <c r="D7" t="inlineStr"><is><t>+47 958 44321</t></is></c>
      <c r="E7" t="inlineStr" s="8"><is><t>Exhibitor</t></is></c>
      <c r="F7" t="inlineStr" s="8"><is><t>S</t></is></c>
      <c r="G7" s="3"><v>35000</v></c>
      <c r="H7" t="inlineStr" s="9"><is><t>Declined / No</t></is></c>
      <c r="I7" t="inlineStr"><is><t>2026-01-28</t></is></c>
      <c r="J7" t="inlineStr"><is><t>2026-02-15</t></is></c>
      <c r="K7" t="inlineStr"><is><t></t></is></c>
      <c r="L7" t="inlineStr"><is><t>Budget frozen this year - try again 2027</t></is></c>
      <c r="M7" t="inlineStr" s="8"><is><t>N</t></is></c>
      <c r="N7" t="inlineStr" s="8"><is><t>N</t></is></c>
      <c r="O7" t="inlineStr" s="8"><is><t>N</t></is></c>
      <c r="P7" t="inlineStr" s="8"><is><t>N</t></is></c>
    </row>
  </sheetData>
  <!-- Status colour coding — 8 stages including Declined -->
  <conditionalFormatting sqref="H2:H1000">
    <cfRule type="cellIs" operator="equal" dxfId="0" priority="8">
      <formula>"To contact"</formula>
    </cfRule>
    <cfRule type="cellIs" operator="equal" dxfId="1" priority="7">
      <formula>"Contacted - no reply"</formula>
    </cfRule>
    <cfRule type="cellIs" operator="equal" dxfId="2" priority="6">
      <formula>"In dialogue"</formula>
    </cfRule>
    <cfRule type="cellIs" operator="equal" dxfId="3" priority="5">
      <formula>"Meeting booked"</formula>
    </cfRule>
    <cfRule type="cellIs" operator="equal" dxfId="4" priority="4">
      <formula>"Proposal sent"</formula>
    </cfRule>
    <cfRule type="cellIs" operator="equal" dxfId="5" priority="3">
      <formula>"Contract sent"</formula>
    </cfRule>
    <cfRule type="cellIs" operator="equal" dxfId="6" priority="2">
      <formula>"SIGNED &amp; PAID"</formula>
    </cfRule>
    <cfRule type="cellIs" operator="equal" dxfId="7" priority="1">
      <formula>"Declined / No"</formula>
    </cfRule>
  </conditionalFormatting>
  <pageSetup orientation="landscape" paperSize="9"/>
</worksheet>
'@

# --- SHEET 2: Budget Dashboard ---
# Now includes: Costs, Capacity (with formula from Pipeline), Revenue, IN Funding, Net Position
Write-Xml (Join-Path $tempDir "xl\worksheets\sheet2.xml") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>
    <col min="1" max="1" width="40" customWidth="1"/>
    <col min="2" max="2" width="20" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1" ht="32" customHeight="1">
      <c r="A1" t="inlineStr" s="1"><is><t>BARCELONA SEAFOOD EXPO 2026 — BUDGET DASHBOARD</t></is></c>
    </row>

    <row r="3" ht="20" customHeight="1">
      <c r="A3" t="inlineStr" s="5"><is><t>COSTS</t></is></c>
    </row>
    <row r="4">
      <c r="A4" t="inlineStr"><is><t>Floor Space Rental (NOK)</t></is></c>
      <c r="B4" s="7"><v>280000</v></c>
    </row>
    <row r="5">
      <c r="A5" t="inlineStr"><is><t>Stand Build / Construction (NOK)</t></is></c>
      <c r="B5" s="7"><v>170000</v></c>
    </row>
    <row r="6">
      <c r="A6" t="inlineStr"><is><t>Catering (NOK)</t></is></c>
      <c r="B6" s="7"><v>50000</v></c>
    </row>
    <row r="7">
      <c r="A7" t="inlineStr"><is><t>Travel / Accommodation (NOK)</t></is></c>
      <c r="B7" s="7"><v>50000</v></c>
    </row>
    <row r="8">
      <c r="A8" t="inlineStr"><is><t>Other / Contingency (NOK)</t></is></c>
      <c r="B8" s="7"><v>30000</v></c>
    </row>
    <row r="9">
      <c r="A9" t="inlineStr" s="6"><is><t>TOTAL GROSS COST (NOK)</t></is></c>
      <c r="B9" s="7"><f>B4+B5+B6+B7+B8</f><v>580000</v></c>
    </row>

    <row r="11" ht="20" customHeight="1">
      <c r="A11" t="inlineStr" s="5"><is><t>INCOME FROM PARTICIPANTS</t></is></c>
    </row>
    <row r="12">
      <c r="A12" t="inlineStr"><is><t>Revenue — Signed &amp; Paid (NOK)</t></is></c>
      <c r="B12" s="7"><f>SUMIF('Company Pipeline'!H:H,"SIGNED &amp; PAID",'Company Pipeline'!G:G)</f><v>0</v></c>
    </row>
    <row r="13">
      <c r="A13" t="inlineStr"><is><t>Revenue — Contracts Sent (not yet paid, NOK)</t></is></c>
      <c r="B13" s="3"><f>SUMIF('Company Pipeline'!H:H,"Contract sent",'Company Pipeline'!G:G)</f><v>0</v></c>
    </row>
    <row r="14">
      <c r="A14" t="inlineStr"><is><t>Pipeline Value — In Progress (NOK)</t></is></c>
      <c r="B14" s="3"><f>SUMIF('Company Pipeline'!H:H,"Proposal sent",'Company Pipeline'!G:G)+SUMIF('Company Pipeline'!H:H,"In dialogue",'Company Pipeline'!G:G)</f><v>0</v></c>
    </row>

    <row r="16" ht="20" customHeight="1">
      <c r="A16" t="inlineStr" s="5"><is><t>INNOVATION NORWAY FUNDING</t></is></c>
    </row>
    <row r="17">
      <c r="A17" t="inlineStr"><is><t>IN / IBE Fund Grant (NOK) — update if confirmed</t></is></c>
      <c r="B17" s="7"><v>0</v></c>
    </row>
    <row r="18">
      <c r="A18" t="inlineStr"><is><t>IN Project Management Hours Covered (NOK)</t></is></c>
      <c r="B18" s="7"><v>0</v></c>
    </row>
    <row r="19">
      <c r="A19" t="inlineStr" s="6"><is><t>TOTAL IN FUNDING (NOK)</t></is></c>
      <c r="B19" s="7"><f>B17+B18</f><v>0</v></c>
    </row>

    <row r="21" ht="20" customHeight="1">
      <c r="A21" t="inlineStr" s="5"><is><t>NET POSITION</t></is></c>
    </row>
    <row r="22">
      <c r="A22" t="inlineStr" s="6"><is><t>Net Cost (after IN funding, NOK)</t></is></c>
      <c r="B22" s="7"><f>B9-B19</f><v>0</v></c>
    </row>
    <row r="23">
      <c r="A23" t="inlineStr" s="6"><is><t>NET PROFIT / LOSS (Revenue minus Net Cost, NOK)</t></is></c>
      <c r="B23" s="7"><f>B12-B22</f><v>0</v></c>
    </row>

    <row r="25" ht="20" customHeight="1">
      <c r="A25" t="inlineStr" s="5"><is><t>CAPACITY &amp; PROGRESS</t></is></c>
    </row>
    <row r="26">
      <c r="A26" t="inlineStr"><is><t>Total Spots Available</t></is></c>
      <c r="B26" s="3"><v>12</v></c>
    </row>
    <row r="27">
      <c r="A27" t="inlineStr"><is><t>Confirmed (Signed &amp; Paid)</t></is></c>
      <c r="B27" s="3"><f>COUNTIF('Company Pipeline'!H:H,"SIGNED &amp; PAID")</f><v>0</v></c>
    </row>
    <row r="28">
      <c r="A28" t="inlineStr"><is><t>Contract Sent (awaiting signature)</t></is></c>
      <c r="B28" s="3"><f>COUNTIF('Company Pipeline'!H:H,"Contract sent")</f><v>0</v></c>
    </row>
    <row r="29">
      <c r="A29" t="inlineStr"><is><t>In Active Pipeline</t></is></c>
      <c r="B29" s="3"><f>COUNTIF('Company Pipeline'!H:H,"In dialogue")+COUNTIF('Company Pipeline'!H:H,"Meeting booked")+COUNTIF('Company Pipeline'!H:H,"Proposal sent")</f><v>0</v></c>
    </row>
    <row r="30">
      <c r="A30" t="inlineStr"><is><t>Declined / No</t></is></c>
      <c r="B30" s="3"><f>COUNTIF('Company Pipeline'!H:H,"Declined / No")</f><v>0</v></c>
    </row>
    <row r="31">
      <c r="A31" t="inlineStr"><is><t>% of Spots Confirmed</t></is></c>
      <c r="B31" s="4"><f>IF(B26&gt;0,B27/B26,0)</f><v>0</v></c>
    </row>

    <row r="33">
      <c r="A33" t="inlineStr"><is><t>NOTE: Update costs (B4-B8) and IN funding (B17-B18) with actual figures. Set Total Spots (B26) to your pavilion capacity.</t></is></c>
    </row>
  </sheetData>
  <pageSetup orientation="portrait" paperSize="9"/>
</worksheet>
'@

# --- SHEET 3: Follow-Up Log ---
Write-Xml (Join-Path $tempDir "xl\worksheets\sheet3.xml") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
      <selection pane="bottomLeft" activeCell="A2" sqref="A2"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>
    <col min="1" max="1" width="14" customWidth="1"/>
    <col min="2" max="2" width="26" customWidth="1"/>
    <col min="3" max="3" width="38" customWidth="1"/>
    <col min="4" max="4" width="38" customWidth="1"/>
    <col min="5" max="5" width="18" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1" ht="36" customHeight="1">
      <c r="A1" t="inlineStr" s="1"><is><t>Date</t></is></c>
      <c r="B1" t="inlineStr" s="1"><is><t>Company</t></is></c>
      <c r="C1" t="inlineStr" s="1"><is><t>Action Taken</t></is></c>
      <c r="D1" t="inlineStr" s="1"><is><t>Next Step</t></is></c>
      <c r="E1" t="inlineStr" s="1"><is><t>Responsible</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>2026-01-15</t></is></c>
      <c r="B2" t="inlineStr"><is><t>Havfisk AS</t></is></c>
      <c r="C2" t="inlineStr"><is><t>Sent initial outreach email</t></is></c>
      <c r="D2" t="inlineStr"><is><t>Follow up if no reply by 20 Jan</t></is></c>
      <c r="E2" t="inlineStr"><is><t>Karin Haugen</t></is></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>2026-01-20</t></is></c>
      <c r="B3" t="inlineStr"><is><t>Havfisk AS</t></is></c>
      <c r="C3" t="inlineStr"><is><t>Follow-up call - very positive, meeting booked</t></is></c>
      <c r="D3" t="inlineStr"><is><t>Send proposal after meeting 22 Jan</t></is></c>
      <c r="E3" t="inlineStr"><is><t>Karin Haugen</t></is></c>
    </row>
    <row r="4">
      <c r="A4" t="inlineStr"><is><t>2026-02-01</t></is></c>
      <c r="B4" t="inlineStr"><is><t>Nordic Seafood Solutions</t></is></c>
      <c r="C4" t="inlineStr"><is><t>Initial outreach email sent</t></is></c>
      <c r="D4" t="inlineStr"><is><t>Follow up by 6 Feb if no reply</t></is></c>
      <c r="E4" t="inlineStr"><is><t>Karin Haugen</t></is></c>
    </row>
    <row r="5">
      <c r="A5" t="inlineStr"><is><t>2026-01-28</t></is></c>
      <c r="B5" t="inlineStr"><is><t>Bergen Pelagic AS</t></is></c>
      <c r="C5" t="inlineStr"><is><t>Called Lars - budget frozen this year, said no</t></is></c>
      <c r="D5" t="inlineStr"><is><t>Note in CRM - re-contact for 2027 show</t></is></c>
      <c r="E5" t="inlineStr"><is><t>Karin Haugen</t></is></c>
    </row>
  </sheetData>
  <pageSetup orientation="landscape" paperSize="9"/>
</worksheet>
'@

if (Test-Path $outFile) { Remove-Item $outFile -Force }
[System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $outFile)
Remove-Item -Recurse -Force $tempDir

Write-Host "Created: $outFile"
