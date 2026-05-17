# =============================================================================
#  RESUME BUILDER (Word .docx + .pdf)
#  Replicates the formatting of Sydney_Gray_Resume_Redcloud.docx exactly.
# =============================================================================
#
#  HOW TO USE
#    1. Edit the $Resume hashtable below with your content.
#    2. Edit $DocxOut and $PdfOut paths if you want them somewhere else.
#    3. From PowerShell:
#         powershell -ExecutionPolicy Bypass -File C:\Users\yendy\resume-template.ps1
#
#  REQUIREMENTS
#    - Microsoft Word installed (Word COM automation; tested with Office 2016+).
#    - Run on Windows.
#
#  WHAT YOU CAN CHANGE FREELY
#    - Anything inside $Resume (name, contact, jobs, bullets, certs, etc.).
#    - Output paths.
#
#  WHAT NOT TO TOUCH (unless you want to change the look)
#    - Build-Resume function below. The font sizes, colors, margins, spacing,
#      border weights, and tab stops are all tuned to match the Redcloud doc.
#
#  FORMATTING SPEC (locked into Build-Resume; here for reference)
#    Page:        US Letter, margins T/B 0.4", L/R 0.6"
#    Font:        Calibri throughout
#    Name:        18pt bold, color #1F2937, centered
#    Contact:     9.5pt centered, separator "  •  "
#    Section
#      headings:  11pt bold, color #1F2937, single 0.75pt bottom border
#                 in #404040, 2pt space before / 1pt after
#    Body text:   9.5pt
#    Job line:    role in bold, right-aligned tab to italic period at 6.5"
#    Company:     italic 9.5pt, " – " between company and city/desc
#    Bullets:     9.5pt, standard Word bullet list (round bullet, hanging indent)
#    Education:   degree/field/school bold; right-tab to italic year at 6.5"
#    Skills:      bold "Label: " prefix, then regular text on same line
#    Areas:       single line, items joined with "  •  "
#
#  UNICODE HELPERS (use these in your text where needed)
#    [char]0x2022 = •  (bullet, used in Contact and Areas)
#    [char]0x2013 = –  (en dash, used in Period and Company lines)
# =============================================================================


# ============== CONTENT — EDIT THIS HASHTABLE ==============
$BULLET = [char]0x2022   # •
$ENDASH = [char]0x2013   # –

$Resume = @{
    Name    = "SYDNEY GRAY"
    Contact = "SydneyGray808@outlook.com  $BULLET  linkedin.com/in/SydneyGray808"

    SummaryHeading = "PROFESSIONAL SUMMARY"
    Summary        = "Operations and administrative professional supporting consulting and professional-services environments. Experienced in internal operations, vendor management, scheduling, systems building, facilities coordination, and cross-functional support. Built 10+ operational mechanisms with 100% adoption at a CEO advisory firm, improving turnaround by 50$($ENDASH)67%. Microsoft 365 fluent, confidentiality-trained, and reliable in fast-paced, multi-stakeholder settings."

    ExperienceHeading = "PROFESSIONAL EXPERIENCE"
    Jobs = @(
        @{
            Role    = "Operations & Executive Support (Contract)"
            Period  = "Nov 2025 $ENDASH Present"
            Company = "C-Suite Success $ENDASH CEO Advisory (Remote)"
            Bullets = @(
                "Built 10+ SOPs and operational systems (vendor coordination, workflow handoffs, prioritization) with 100% team adoption and 50$($ENDASH)67% faster turnaround.",
                "Manage day-to-day operational infrastructure for the CEO: scheduling support, status tracking, vendor outreach, and cross-functional follow-through.",
                "Triage 150+ emails weekly with zero errors; maintain internal documentation and process guides.",
                "Coordinated offsite networking event end-to-end: venue sourcing, catering, logistics, attendee outreach, and on-site hosting.",
                "Operate within a professional-services environment requiring confidentiality, precision, and executive-level communication."
            )
        },
        @{
            Role    = "Assistant Manager, Operations & Front Office"
            Period  = "Sep 2025 $ENDASH Present"
            Company = "Weldon Barber"
            Bullets = @(
                "First point of contact for ~250 weekly guests in a 13-person professional-services environment; oversee check-in, visitor management, and front-office readiness.",
                "Maintain accurate inventory of ~300 retail, service, and operational supplies using Booker POS and Excel-based tracking systems; reconcile discrepancies and ensure availability.",
                "Coordinate vendor relationships and manage office/employee supply operations.",
                "Oversee facilities operations for all technological, employee, and guest usage; diagnose root causes, document incidents, and escalate tickets to IT/HR through resolution.",
                "Support onboarding: system access, Day 1 setup, training schedules, and completion of operational checklists."
            )
        },
        @{
            Role    = "Operations & Events Associate"
            Period  = "Aug 2025 $ENDASH Present"
            Company = "Pure Barre"
            Bullets = @(
                "Serve as primary front-office contact for 300+ weekly clients; manage check-in systems, visitor support, and client experience.",
                "Manage operational readiness across supply inventory, equipment rooms, and client-facing spaces.",
                "Planned and executed 16 community and partnership events (15$($ENDASH)200 attendees): vendor coordination, logistics, materials, supply management, and post-event follow-up.",
                "Support onboarding and day-to-day coaching for 10 staff members."
            )
        },
        @{
            Role    = "Behavioral Technician (Operations & Coordination)"
            Period  = "Feb 2024 $ENDASH Mar 2025"
            Company = "Amergis"
            Bullets = @(
                "Coordinated 30 interdisciplinary meetings end-to-end: agendas, documentation, supply prep, and stakeholder communication.",
                "Built organizational systems saving ~400 hours annually across multi-site teams.",
                "Maintained confidential documentation with HIPAA-aligned handling.",
                "Supported 300 students through structured, data-driven systems and cross-functional collaboration with educators, clinicians, and administrators."
            )
        }
    )

    EducationHeading = "EDUCATION & CERTIFICATIONS"
    # Education shows as: "Degree  Field  •  School" left-side, "Year" right-side italic.
    Degree   = "B.A."
    Field    = "Psychology"
    School   = "University of Hawaii"
    GradYear = "2023"
    Certifications = @(
        "Public Value Leadership $ENDASH Harvard",
        "Foundational Leadership $ENDASH Stanford",
        "Introduction to AI $ENDASH Stanford"
    )

    SkillsHeading = "TECHNICAL SKILLS"
    SkillGroups = @(
        @{ Label = "Microsoft 365";              Items = "Outlook, Excel, Word, PowerPoint, Teams, OneDrive" },
        @{ Label = "Operations & Productivity";  Items = "Notion, Google Workspace (advanced Sheets), SOP authoring, asset/inventory tracking" },
        @{ Label = "Front-of-House & POS";       Items = "Booker, ClubReady, MindBody" },
        @{ Label = "Event Coordination";         Items = "LinkedIn, Luma, Eventbrite, vendor/venue sourcing, catering logistics" }
    )

    AreasHeading = "AREAS OF EXPERTISE"
    Areas = @(
        "Office Operations",
        "Internal Coordination",
        "Vendor Management",
        "Facilities Support",
        "IT Troubleshooting",
        "Onboarding",
        "Event Execution",
        "Professional Services Support"
    )
}

# ============== OUTPUT PATHS ==============
$DocxOut = "C:\Users\yendy\my-resume.docx"
$PdfOut  = "C:\Users\yendy\my-resume.pdf"   # set to $null to skip PDF export


# ============== BUILDER (don't touch unless changing the look) ==============

function Build-Resume {
    param(
        [Parameter(Mandatory)] [hashtable]$R,
        [Parameter(Mandatory)] [string]$DocxPath,
        [string]$PdfPath
    )

    # --- Color helpers (Word COM uses BGR ints, NOT RGB) ---
    function HexToBgr([string]$hex) {
        $r = [Convert]::ToInt32($hex.Substring(0,2), 16)
        $g = [Convert]::ToInt32($hex.Substring(2,2), 16)
        $b = [Convert]::ToInt32($hex.Substring(4,2), 16)
        return ($b -shl 16) -bor ($g -shl 8) -bor $r
    }
    $C_DARK   = HexToBgr "1F2937"   # name + section heading text
    $C_BORDER = HexToBgr "404040"   # section heading bottom border
    $C_AUTO   = -16777216           # wdColorAutomatic (true black/auto)

    # --- Word COM constants (literal values to avoid loading the assembly) ---
    $wdAlignLeft        = 0
    $wdAlignCenter      = 1
    $wdAlignRight       = 2
    $wdBorderBottom     = -3
    $wdLineStyleNone    = 0
    $wdLineStyleSingle  = 1
    $wdFormatXMLDocument = 16
    $wdExportFormatPDF   = 17
    # Margins/PageSetup take points. 1 inch = 72 pt.
    function Inches([double]$i) { return $i * 72 }

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0   # wdAlertsNone
    try {
        $doc = $word.Documents.Add()
        $sel = $word.Selection

        # --- Page setup ---
        $ps = $doc.PageSetup
        $ps.PageWidth     = Inches 8.5
        $ps.PageHeight    = Inches 11
        $ps.TopMargin     = Inches 0.4
        $ps.BottomMargin  = Inches 0.4
        $ps.LeftMargin    = Inches 0.6
        $ps.RightMargin   = Inches 0.6

        # --- Override Normal-style defaults (CRITICAL for single-page fit) ---
        # Word 2016+ Normal template injects 8pt SpaceAfter and 1.15 line spacing
        # on every paragraph by default. The original docx had neither. Without
        # this override, every paragraph gets ~30% more vertical space than it
        # should, and the resume overflows to page 2.
        $wdStyleNormal      = -1
        $wdLineSpaceSingle  = 0
        $normal = $doc.Styles.Item($wdStyleNormal)
        $normal.ParagraphFormat.SpaceAfter      = 0
        $normal.ParagraphFormat.SpaceBefore     = 0
        $normal.ParagraphFormat.LineSpacingRule = $wdLineSpaceSingle

        # --- Default paragraph + font baseline ---
        $sel.Font.Name = "Calibri"
        $sel.Font.Size = 9.5
        $sel.Font.Color = $C_AUTO
        $sel.Font.Bold = $false
        $sel.Font.Italic = $false
        $sel.ParagraphFormat.SpaceAfter = 0
        $sel.ParagraphFormat.SpaceBefore = 0
        $sel.ParagraphFormat.LineSpacingRule = $wdLineSpaceSingle
        $sel.ParagraphFormat.Alignment = $wdAlignLeft

        # --- Helper: type plain text run with optional bold/italic ---
        function Write-Run([string]$text, [bool]$bold = $false, [bool]$italic = $false) {
            $sel.Font.Bold   = $bold
            $sel.Font.Italic = $italic
            $sel.TypeText($text)
            $sel.Font.Bold   = $false
            $sel.Font.Italic = $false
        }

        # --- Helper: section heading (11pt bold + bottom border) ---
        # Two subtle Word-COM rules at play here:
        #   1. Apply the border directly to the heading paragraph object — NOT
        #      via Selection.ParagraphFormat. The Selection-based path is
        #      unreliable: a later "clear" call on the Selection nukes the
        #      heading's border retroactively.
        #   2. Word INHERITS paragraph borders through Enter/TypeParagraph. So
        #      after we move to the next paragraph, the new paragraph carries
        #      the border too. We must explicitly clear the border on the new
        #      paragraph or every paragraph in the doc ends up with a line.
        function Write-SectionHeading([string]$title) {
            $sel.Font.Bold  = $true
            $sel.Font.Size  = 11
            $sel.Font.Color = $C_DARK
            $sel.ParagraphFormat.SpaceBefore = 2
            $sel.ParagraphFormat.SpaceAfter  = 1
            $sel.TypeText($title)

            # (1) Apply border to the just-typed heading paragraph.
            $p = $sel.Paragraphs.Item(1)
            $bottom = $p.Borders.Item($wdBorderBottom)
            $bottom.LineStyle = $wdLineStyleSingle
            $bottom.LineWidth = 6           # 6 eighth-points = 0.75pt
            $bottom.Color     = $C_BORDER

            $sel.TypeParagraph()

            # (2) Clear the border on the NEW paragraph the cursor is now in,
            # so it doesn't propagate to body content below.
            $pNew = $sel.Paragraphs.Item(1)
            $pNew.Borders.Item($wdBorderBottom).LineStyle = $wdLineStyleNone

            # Reset run-level formatting for the body that follows.
            $sel.Font.Bold  = $false
            $sel.Font.Size  = 9.5
            $sel.Font.Color = $C_AUTO
            $sel.ParagraphFormat.SpaceBefore = 0
            $sel.ParagraphFormat.SpaceAfter  = 0
        }

        # --- Helper: line with right-aligned tab at 6.5" (left text | right text) ---
        function Write-TabbedLine {
            param(
                [string]$LeftText, [bool]$LeftBold = $false, [bool]$LeftItalic = $false,
                [string]$RightText, [bool]$RightBold = $false, [bool]$RightItalic = $false
            )
            $sel.ParagraphFormat.TabStops.ClearAll()
            $null = $sel.ParagraphFormat.TabStops.Add((Inches 6.5), $wdAlignRight)
            Write-Run $LeftText $LeftBold $LeftItalic
            $sel.TypeText("`t")
            Write-Run $RightText $RightBold $RightItalic
            $sel.TypeParagraph()
            $sel.ParagraphFormat.TabStops.ClearAll()
        }

        # --- Helper: bulleted list from an array of strings ---
        # Word's default bullet template uses 0.5"/0.25" hanging indent — too
        # wide for a single-page resume. The original docx used 0.25"/0.167",
        # so we override after applying the bullet.
        function Write-Bullets([string[]]$items) {
            foreach ($t in $items) {
                $sel.TypeText($t)
                $sel.Range.ListFormat.ApplyBulletDefault()
                # Tighten indents on the just-bulleted paragraph (points: 0.25" = 18pt, hanging 0.167" = 12pt)
                $p = $sel.Paragraphs.Item(1)
                $p.LeftIndent       = 18
                $p.FirstLineIndent  = -12
                $sel.TypeParagraph()
            }
            # Detach the trailing paragraph from the bullet list and reset indent.
            $sel.Range.ListFormat.RemoveNumbers()
            $sel.ParagraphFormat.LeftIndent      = 0
            $sel.ParagraphFormat.FirstLineIndent = 0
        }

        # ============== HEADER ==============
        $sel.ParagraphFormat.Alignment = $wdAlignCenter
        $sel.Font.Bold  = $true
        $sel.Font.Size  = 18
        $sel.Font.Color = $C_DARK
        $sel.TypeText($R.Name)
        $sel.TypeParagraph()

        # Contact line (centered, 9.5pt regular)
        $sel.Font.Bold  = $false
        $sel.Font.Size  = 9.5
        $sel.Font.Color = $C_AUTO
        $sel.TypeText($R.Contact)
        $sel.TypeParagraph()

        # Back to left-aligned for everything below.
        $sel.ParagraphFormat.Alignment = $wdAlignLeft

        # ============== SUMMARY ==============
        Write-SectionHeading $R.SummaryHeading
        $sel.TypeText($R.Summary)
        $sel.TypeParagraph()

        # ============== EXPERIENCE ==============
        Write-SectionHeading $R.ExperienceHeading
        foreach ($job in $R.Jobs) {
            Write-TabbedLine -LeftText $job.Role -LeftBold $true `
                             -RightText $job.Period -RightItalic $true
            # Company line (italic)
            Write-Run $job.Company $false $true
            $sel.TypeParagraph()
            Write-Bullets $job.Bullets
        }

        # ============== EDUCATION & CERTIFICATIONS ==============
        Write-SectionHeading $R.EducationHeading
        # Education line: "Degree  Field  •  School" bold left, italic year right.
        $eduLeft = "$($R.Degree)  $($R.Field)  $BULLET   $($R.School)"
        Write-TabbedLine -LeftText $eduLeft -LeftBold $true `
                         -RightText $R.GradYear -RightItalic $true
        Write-Bullets $R.Certifications

        # ============== TECHNICAL SKILLS ==============
        Write-SectionHeading $R.SkillsHeading
        foreach ($g in $R.SkillGroups) {
            Write-Run "$($g.Label): " $true $false
            Write-Run $g.Items $false $false
            $sel.TypeParagraph()
        }

        # ============== AREAS OF EXPERTISE ==============
        Write-SectionHeading $R.AreasHeading
        $areasLine = ($R.Areas -join "  $BULLET  ")
        $sel.TypeText($areasLine)
        $sel.ParagraphFormat.SpaceAfter = 1   # tiny breathing room at the very end

        # --- Verify single-page constraint BEFORE saving ---
        $wdStatisticPages = 2
        $pages = $doc.ComputeStatistics($wdStatisticPages)

        # --- Save ---
        $doc.SaveAs([ref]$DocxPath, [ref]$wdFormatXMLDocument)
        if ($PdfPath) {
            $doc.SaveAs([ref]$PdfPath, [ref]$wdExportFormatPDF)
        }
        $doc.Close($false)
        Write-Output "Saved: $DocxPath"
        if ($PdfPath) { Write-Output "Saved: $PdfPath" }
        if ($pages -gt 1) {
            Write-Warning "Resume is $pages pages - content overflows. Trim a bullet or two, shorten the summary, or remove a less-relevant job to get back to 1 page."
        } else {
            Write-Output "Page count: $pages (fits on one page)."
        }
    }
    finally {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

Build-Resume -R $Resume -DocxPath $DocxOut -PdfPath $PdfOut
