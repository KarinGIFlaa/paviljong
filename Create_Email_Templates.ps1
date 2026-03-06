Add-Type -Assembly "System.IO.Compression"
Add-Type -Assembly "System.IO.Compression.FileSystem"

$baseDir = "C:\Users\chris\Documents\Claude kurs"
$outFile = Join-Path $baseDir "Email_Templates_Barcelona.docx"
$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("docx_" + [System.Guid]::NewGuid().ToString("N"))

function Write-Xml($path, $content) {
    $enc = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($path, $content, $enc)
}

foreach ($d in @("","_rels","word","word\_rels")) {
    New-Item -ItemType Directory -Path (Join-Path $tempDir $d) -Force | Out-Null
}

Write-Xml (Join-Path $tempDir "[Content_Types].xml") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>
'@

Write-Xml (Join-Path $tempDir "_rels\.rels") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
'@

Write-Xml (Join-Path $tempDir "word\_rels\document.xml.rels") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>
'@

Write-Xml (Join-Path $tempDir "word\settings.xml") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
</w:settings>
'@

Write-Xml (Join-Path $tempDir "word\styles.xml") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr>
      <w:pBdr>
        <w:bottom w:val="single" w:sz="8" w:space="4" w:color="1F3864"/>
      </w:pBdr>
      <w:spacing w:before="480" w:after="120"/>
      <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:color w:val="1F3864"/>
      <w:sz w:val="36"/>
      <w:szCs w:val="36"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:pPr>
      <w:spacing w:before="320" w:after="80"/>
      <w:outlineLvl w:val="1"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:color w:val="2E74B5"/>
      <w:sz w:val="26"/>
      <w:szCs w:val="26"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="SubjectLine">
    <w:name w:val="SubjectLine"/>
    <w:pPr>
      <w:spacing w:before="80" w:after="80"/>
      <w:shd w:val="clear" w:color="auto" w:fill="EEF3FA"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:color w:val="1F3864"/>
      <w:sz w:val="22"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="TemplateBody">
    <w:name w:val="TemplateBody"/>
    <w:pPr>
      <w:spacing w:before="80" w:after="120" w:line="276" w:lineRule="auto"/>
      <w:ind w:left="360"/>
    </w:pPr>
  </w:style>
</w:styles>
'@

Write-Xml (Join-Path $tempDir "word\document.xml") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>

    <!-- TITLE -->
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
        <w:jc w:val="center"/>
      </w:pPr>
      <w:r><w:t>Email Template Library</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r><w:rPr><w:color w:val="595959"/></w:rPr>
        <w:t>Norwegian Pavilion — Barcelona Seafood Expo 2026</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r><w:rPr><w:color w:val="595959"/><w:i/></w:rPr>
        <w:t>Instructions: Replace all [bracketed placeholders] before sending. Use BCC for bulk outreach.</w:t>
      </w:r>
    </w:p>

    <!-- ============================================================ -->
    <!-- TEMPLATE 1 -->
    <!-- ============================================================ -->
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Template 1: Initial Outreach (Day 0)</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="SubjectLine"/></w:pPr>
      <w:r><w:t>Subject: Norwegian Pavilion — Barcelona Seafood Expo 2026</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t xml:space="preserve">Hi [First Name],</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>My name is Karin Haugen, and I manage the Norwegian pavilion at international seafood trade shows. We are currently building our delegation for Barcelona Seafood Expo 2026 (taking place [Date, e.g. 21–23 April 2026]), and based on [Company Name]'s work in [specific product/market area], I think this is a strong fit.</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>The expo attracts buyers, distributors and decision-makers from across Europe and Asia. We offer a turnkey package — you show up, we handle the stand, furniture, branding, and logistics.</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Spots are limited. Would you have 15 minutes this week or next to hear more?</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t xml:space="preserve">Best regards,</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Karin Haugen</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>[Phone] | [Email]</w:t></w:r>
    </w:p>

    <!-- ============================================================ -->
    <!-- TEMPLATE 2 -->
    <!-- ============================================================ -->
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Template 2: Follow-Up 1 (Day 5 — no reply)</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="SubjectLine"/></w:pPr>
      <w:r><w:t>Subject: Re: Norwegian Pavilion — Barcelona Seafood Expo 2026</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Hi [First Name],</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>I sent a note last week about our Norwegian pavilion at Barcelona Seafood Expo 2026 — just wanted to make sure it didn't get buried.</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Is this something worth a quick 15-minute conversation? Happy to keep it short and give you an honest picture of whether it makes sense for [Company Name] this year.</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Best, Karin</w:t></w:r>
    </w:p>

    <!-- ============================================================ -->
    <!-- TEMPLATE 3 -->
    <!-- ============================================================ -->
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Template 3: Follow-Up 2 / Last Attempt (Day 10)</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="SubjectLine"/></w:pPr>
      <w:r><w:t>Subject: Last note — Barcelona Seafood Expo 2026</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Hi [First Name],</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>This will be my last reach-out on this. We have one [S/M/L] spot remaining in the Norwegian pavilion at Barcelona 2026, and I wanted to offer it to [Company Name] before we close the list.</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>If the timing isn't right, no problem at all — feel free to reach out for future shows. If you'd like to explore it, I can send the details in 5 minutes.</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Karin</w:t></w:r>
    </w:p>

    <!-- ============================================================ -->
    <!-- TEMPLATE 4 -->
    <!-- ============================================================ -->
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Template 4: Meeting Confirmation</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="SubjectLine"/></w:pPr>
      <w:r><w:t>Subject: Confirmed — Barcelona Seafood Expo Briefing, [Date] at [Time]</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Hi [First Name],</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Great — looking forward to our call on [Date] at [Time] ([Timezone]).</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>We will cover:</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/><w:ind w:left="720"/></w:pPr>
      <w:r><w:t>- What is included in your pavilion package (space, stand, branding, wifi, catering)</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/><w:ind w:left="720"/></w:pPr>
      <w:r><w:t>- Booth size options and pricing (S / M / L)</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/><w:ind w:left="720"/></w:pPr>
      <w:r><w:t>- Key dates and confirmation deadline</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/><w:ind w:left="720"/></w:pPr>
      <w:r><w:t>- Any questions you have</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Dial-in: [Teams / Zoom / Phone link]</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>See you then, Karin</w:t></w:r>
    </w:p>

    <!-- ============================================================ -->
    <!-- TEMPLATE 5 -->
    <!-- ============================================================ -->
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Template 5: Proposal Email</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="SubjectLine"/></w:pPr>
      <w:r><w:t>Subject: Barcelona Seafood Expo 2026 — Pavilion Proposal for [Company Name]</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Hi [First Name],</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Thank you for our conversation earlier. As discussed, please find attached the pricing sheet for [Company Name]'s participation in the Norwegian Pavilion at Barcelona Seafood Expo 2026.</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Summary:</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/><w:ind w:left="720"/></w:pPr>
      <w:r><w:t>- Booth size: [S / M / L]</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/><w:ind w:left="720"/></w:pPr>
      <w:r><w:t>- Investment: NOK [Amount] + VAT</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/><w:ind w:left="720"/></w:pPr>
      <w:r><w:t>- Includes: Stand space, furniture, company signage, wifi, [other inclusions]</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/><w:ind w:left="720"/></w:pPr>
      <w:r><w:t>- Early bird deadline: [Date] (10% discount applies)</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>I'm happy to answer any questions or adjust the package. If you'd like to move forward, I can have a contract to you by end of this week.</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Best regards, Karin</w:t></w:r>
    </w:p>

    <!-- ============================================================ -->
    <!-- TEMPLATE 6 -->
    <!-- ============================================================ -->
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Template 6: Contract Email</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="SubjectLine"/></w:pPr>
      <w:r><w:t>Subject: Contract — [Company Name] / Norwegian Pavilion Barcelona 2026</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Hi [First Name],</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Please find below your participation contract for the Norwegian Pavilion at Barcelona Seafood Expo 2026.</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Sign digitally here (takes under 2 minutes):</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/><w:ind w:left="720"/></w:pPr>
      <w:r><w:t>[DocuSign / Docuseal signing link]</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Once signed, I will send your invoice immediately and your spot is confirmed.</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Deadline to confirm your spot: [Date]</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Any questions, just call or reply here. Looking forward to having [Company Name] in the pavilion.</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Best, Karin</w:t></w:r>
    </w:p>

    <!-- ============================================================ -->
    <!-- TEMPLATE 7 -->
    <!-- ============================================================ -->
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Template 7: Payment Reminder</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t>7a — Friendly first reminder</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="SubjectLine"/></w:pPr>
      <w:r><w:t>Subject: Reminder: Invoice [#] Due [Date] — Norwegian Pavilion Barcelona 2026</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Hi [First Name],</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Just a friendly reminder that invoice [#] for NOK [Amount] (50% deposit) is due on [Date].</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Payment details:</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/><w:ind w:left="720"/></w:pPr>
      <w:r><w:t>Bank: [Bank Name] | Account: [Account No.] | IBAN: [IBAN] | Ref: Invoice [#]</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Let me know if you have any questions or need a different format.</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Thanks, Karin</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t>7b — Firm second reminder (overdue)</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="SubjectLine"/></w:pPr>
      <w:r><w:t>Subject: OVERDUE: Invoice [#] — Norwegian Pavilion Barcelona 2026</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Hi [First Name],</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Invoice [#] for NOK [Amount] is now [X] days overdue. Your spot in the pavilion is reserved pending receipt of payment.</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Please process payment by [Hard Deadline Date].</w:t></w:r>
      <w:r><w:t xml:space="preserve"> After this date we may need to release your spot to other companies on the waiting list.</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Please confirm receipt of this message.</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Karin Haugen | [Phone]</w:t></w:r>
    </w:p>

    <!-- ============================================================ -->
    <!-- TEMPLATE 8 -->
    <!-- ============================================================ -->
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Template 8: Welcome / Confirmed Participant</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="SubjectLine"/></w:pPr>
      <w:r><w:t>Subject: You're confirmed! Norwegian Pavilion — Barcelona Seafood Expo 2026</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Hi [First Name],</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Fantastic — [Company Name] is officially confirmed for the Norwegian Pavilion at Barcelona Seafood Expo 2026!</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Here is what happens next:</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/><w:ind w:left="720"/></w:pPr>
      <w:r><w:t>1. Exhibitor information pack — I will send this within 5 business days</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/><w:ind w:left="720"/></w:pPr>
      <w:r><w:t>2. Branding deadline: [Date] — please send your logo (EPS/PNG) and 50-word company description</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/><w:ind w:left="720"/></w:pPr>
      <w:r><w:t>3. Build day: [Date] — access from [Time]. Arrival instructions to follow.</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/><w:ind w:left="720"/></w:pPr>
      <w:r><w:t>4. Show opens: [Date] at [Time]</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>Welcome to the team. Let's make it a great show!</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Karin Haugen</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="TemplateBody"/></w:pPr>
      <w:r><w:t>[Phone] | [Email]</w:t></w:r>
    </w:p>

    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>
'@

if (Test-Path $outFile) { Remove-Item $outFile -Force }
[System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $outFile)
Remove-Item -Recurse -Force $tempDir

Write-Host "Created: $outFile"
