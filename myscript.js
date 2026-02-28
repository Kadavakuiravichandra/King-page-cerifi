function getCommand() {
  const term = document.getElementById("termInput").value.trim();

  if (!term) {
    alert("Enter term id");
    return;
  }

  const command = `e -t ${term}`;

  // show output
  document.getElementById("termOutput").innerText = command;

  // try modern copy
  if (navigator.clipboard && window.isSecureContext) {
    navigator.clipboard
      .writeText(command)
      .then(() => showMsg("Copied!"))
      .catch(() => fallbackCopy(command));
  } else {
    fallbackCopy(command);
  }
}

// fallback copy method
function fallbackCopy(text) {
  const temp = document.createElement("textarea");
  temp.value = text;
  document.body.appendChild(temp);
  temp.select();
  temp.setSelectionRange(0, 99999);

  try {
    document.execCommand("copy");
    showMsg("Copied!");
  } catch (err) {
    alert("Copy failed");
  }

  document.body.removeChild(temp);
}

// optional message
function showMsg(msg) {
  console.log(msg); // or replace with snackbar
}

function generateCommands() {
  const input = document.getElementById("courseInput").value.trim();

  if (!input) {
    alert("Enter course name");
    return;
  }

  const courses = input.split(/\s+/);

  let commands = [];

  courses.forEach((c) => {
    commands.push(`ux2w ${c}`);
    commands.push(`mtermlist ${c}`);
    commands.push(`ux2w -g ${c}`);
    commands.push(`ux2wm ${c}`);
    commands.push(`mprf ${c}`);
    commands.push(`qcom ${c}`);
    commands.push(`qquest ${c}`);
    commands.push(`moutl ${c}`);
  });

  const finalText = commands.join("; ");

  // show output
  document.getElementById("output").innerText = finalText;

  // ✅ AUTO COPY HERE
  navigator.clipboard
    .writeText(finalText)
    .then(() => {
      showSnackbar("Commands copied!");
    })
    .catch(() => {
      showSnackbar("Copy failed");
    });
}

// Run after page loads
document.addEventListener("DOMContentLoaded", function () {
  const items = document.querySelectorAll("li[data-copy]");

  items.forEach((item) => {
    item.addEventListener("click", function () {
      const text = this.getAttribute("data-copy");

      navigator.clipboard
        .writeText(text)
        .then(() => {
          showSnackbar("Copied: " + text);
          highlightItem(this);
        })
        .catch(() => {
          alert("Copy failed");
        });
    });
  });
});

// Highlight effect
function highlightItem(element) {
  element.classList.add("copied");
  setTimeout(() => {
    element.classList.remove("copied");
  }, 800);
}

/* ============================================
   Uiicons Copy Function
   ============================================ */

function copyText(text) {
  navigator.clipboard.writeText(text);

  let toast = document.createElement("div");
  toast.innerText = text + " copied!";
  toast.style.position = "fixed";
  toast.style.bottom = "20px";
  toast.style.right = "20px";
  toast.style.background = "black";
  toast.style.color = "white";
  toast.style.padding = "10px 15px";
  toast.style.borderRadius = "5px";

  document.body.appendChild(toast);

  setTimeout(() => {
    toast.remove();
  }, 2000);
}

/* ============================================
   create term entity files from text input macro
   ============================================ */

function CreateTerm() {
  const macro = `param(
  [string]$InputFile  = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Terms Creation\\terms.txt'),
  [string]$IdPrefix   = "001",
  [string]$OutputRoot = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Term Creation Output')
)

function Read-TextSmart([string]$path) {
  $bytes = [System.IO.File]::ReadAllBytes($path)
  $utf8 = New-Object System.Text.UTF8Encoding($false, $true)
  try { $t1 = $utf8.GetString($bytes) } catch { $t1 = [System.Text.Encoding]::UTF8.GetString($bytes) }

  $hasMojibake = ($t1.IndexOf([char]0x00C3) -ge 0) -or ($t1.IndexOf([char]0x00C2) -ge 0) -or ($t1.IndexOf([char]0x00E2) -ge 0)
  if ($hasMojibake) {
    $cp1252 = [System.Text.Encoding]::GetEncoding(1252)
    return $cp1252.GetString($bytes)
  }
  return $t1
}

function Escape-Xml([string]$s) {
  if ($null -eq $s) { return "" }
  $s = $s -replace "&","&amp;"
  $s = $s -replace "<","&lt;"
  $s = $s -replace ">","&gt;"
  $s = $s -replace '"',"&quot;"
  $s = $s -replace "'","&apos;"
  return $s
}

function Normalize-Text([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return "" }

  $s = $s -replace [char]0x00A0, " "
  $s = $s -replace "[\\u200B-\\u200D\\uFEFF]", ""
  $s = $s -replace '^[\\s\\p{Cf}]*(?:[\\u2022\\u25CF\\u25AA\\u00B7\\u2219\\u2013\\u2014\\-\\.\\u00B7]+)\\s*', ''

  $s = $s.Replace([string][char]0x2014, "-")
  $s = $s.Replace([string][char]0x2013, "-")

  $s = $s.Replace([string][char]0x00A7, "Section ")
  $s = $s.Replace([string][char]0x00BD, "1/2")

  $s = ($s -replace "\\s+"," ").Trim()
  $s = $s -replace '\\s+([,.;:!?])','$1'

  return $s
}

function Is-ChapterLine([string]$line) {
  return ($line -match '^\\s*Chapter\\s+\\d+\\s*$')
}

function Is-NumberedStart([string]$line) {
  return ($line -match '^\\s*\\d+\\s*[\\.\\)]\\s*\\S')
}

function Strip-Numbering([string]$line) {
  return ([regex]::Replace($line, '^\\s*\\d+\\s*[\\.\\)]\\s*', '')).Trim()
}

function Split-TitleDef([string]$line) {
  if ($line -match '^(?<t>.+?)\\s+--\\s+(?<d>.+)$') { return @{ Title=$matches.t; Def=$matches.d } }
  if ($line -match '^(?<t>.+?)\\s+-\\s+(?<d>.+)$')  { return @{ Title=$matches.t; Def=$matches.d } }
  return $null
}

if (!(Test-Path $InputFile)) { Write-Host "ERROR: Input file not found: $InputFile"; exit 1 }

$raw = Read-TextSmart $InputFile
$raw = $raw -replace "\`r\`n","\`n"
$raw = $raw -replace "\`r","\`n"
$lines = $raw -split "\`n"

$terms   = New-Object System.Collections.Generic.List[object]
$skipped = New-Object System.Collections.Generic.List[string]

$i = 0
while ($i -lt $lines.Count) {
  $lineRaw = $lines[$i]
  $line = Normalize-Text $lineRaw
  $i++

  if ($line -eq "" -or (Is-ChapterLine $line)) { continue }

  if (Is-NumberedStart $line) {
    $line = Normalize-Text (Strip-Numbering $line)
  }

  $split = Split-TitleDef $line
  if ($null -ne $split) {
    $title = Normalize-Text $split.Title
    $def   = Normalize-Text $split.Def
    if ($title -ne "" -and $def -ne "") {
      $terms.Add([pscustomobject]@{ Title=$title; Def=$def })
    } else {
      $skipped.Add($lineRaw)
    }
    continue
  }

  if ($line -ne "" -and $i -lt $lines.Count) {
    $title = $line
    $defParts = New-Object System.Collections.Generic.List[string]

    while ($i -lt $lines.Count) {
      $peekRaw = $lines[$i]
      $peek = Normalize-Text $peekRaw

      if ($peek -eq "" -or (Is-ChapterLine $peek) -or (Is-NumberedStart $peek)) { break }

      $defParts.Add($peek)
      $i++
    }

    $def = Normalize-Text (($defParts -join " ").Trim())
    if ($title -ne "" -and $def -ne "") {
      $terms.Add([pscustomobject]@{ Title=$title; Def=$def })
    } else {
      $skipped.Add($lineRaw)
    }
    continue
  }

  $skipped.Add($lineRaw)
}

$terms = $terms | Sort-Object Title

$outDir = Join-Path $OutputRoot ("out_" + $IdPrefix)
if (!(Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

Get-ChildItem -LiteralPath $outDir -Filter ($IdPrefix + "*.ent") -ErrorAction SilentlyContinue |
  Remove-Item -Force -ErrorAction SilentlyContinue

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

$index = New-Object System.Collections.Generic.List[string]
$counter = 1

foreach ($t in $terms) {
  $id = $IdPrefix + ("{0:D2}" -f $counter)

  $xml = @"
<dlentry id="$id">
<title>$(Escape-Xml $t.Title)</title>
<definition>
<para>$(Escape-Xml $t.Def)</para>
</definition>
</dlentry>
"@
  $xml = $xml -replace "\`n","\`r\`n"
  [System.IO.File]::WriteAllText((Join-Path $outDir ($id + ".ent")), $xml, $utf8NoBom)
  $index.Add("$id\`t$($t.Title)")
  $counter++
}

[System.IO.File]::WriteAllLines((Join-Path $outDir "index.txt"), $index, $utf8NoBom)
[System.IO.File]::WriteAllLines((Join-Path $outDir "skipped.txt"), $skipped, $utf8NoBom)

Write-Host ("Output folder : {0}" -f $outDir)
Write-Host ("Parsed terms  : {0}" -f $terms.Count)
Write-Host ("Created files : {0}" -f (Get-ChildItem -LiteralPath $outDir -Filter ($IdPrefix + "*.ent") -ErrorAction SilentlyContinue).Count)
Write-Host ("Skipped lines : {0}" -f $skipped.Count)
Write-Host ("Skipped file  : {0}" -f (Join-Path $outDir "skipped.txt"))`;

  navigator.clipboard.writeText(macro).then(() => {
    showSnackbar("Macro copied!");
  });
}

// ============================================
// Detecting Paragraphs Not Using Approved Styles and Extracting to New Document Macro
// ============================================

function NotApprovedStyles() {
  const macro = `Option Explicit

Public Sub Extract_Text_Outside_CustomStyleList_ToNewDoc_NoFindKey()

    Dim src As Document
    Dim outDoc As Document
    Dim allowed As Object
    Dim styleList As Variant
    Dim key As Variant

    Dim p As Paragraph
    Dim stName As String
    Dim rng As Range
    Dim outRng As Range

    Dim copied As Long
    Dim pg As Long
    Dim snippet As String

    Set src = ActiveDocument
    Set allowed = CreateObject("Scripting.Dictionary")

    styleList = Array( _
        "acronym", "ans-correct", "answer-table", "authority", "blank", "blue", _
        "bold", "bolditalic", "booktitle", "bsl body", "bsl title", _
        "chapter desc", "chapter title", "ednote", "frac-a", "graphic", _
        "head", "head2", "head3", "inline-head", "irc", "italic", _
        "list 1", "list 1 cont", "list 2", "list 2 cont", "list 3", "list 3 cont", _
        "list bullet 1", "list bullet 2", "list bullet 3", _
        "list forced 1", "list forced 2", "list forced 3", _
        "objectives", "panel", "para", "remedial", "screen-break", _
        "study question", "sub", "subchapter title", "sup", "supplement", _
        "tbldu", "tbldu-center", "tbldu-right", "tblu", "tblu-center", "tblu-right", _
        "td", "td-center", "td-right", "term", "test question", _
        "th", "th-center", "th-right", "notenum", "url", _
        "wrapper-begin", "wrapper-end", _
        "num-of-test-questions-to-display", _
        "instruction", _
        "printonly-begin", _
        "printonly-end", _
        "ul-attributes" _
    )

    For Each key In styleList
        allowed(LCase$(CStr(key))) = True
    Next key

    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    Application.StatusBar = "Extracting paragraphs..."

    Set outDoc = Documents.Add
    outDoc.Content.Text = "PARAGRAPHS NOT USING APPROVED STYLES" & vbCrLf & _
                          "===================================" & vbCrLf & vbCrLf

    Set outRng = outDoc.Range
    outRng.Collapse wdCollapseEnd

    copied = 0

    For Each p In src.StoryRanges(wdMainTextStory).Paragraphs

        stName = LCase$(CStr(p.Range.Style))

        If Not allowed.Exists(stName) Then

            Set rng = p.Range.Duplicate

            If rng.Characters.Count > 0 Then
                If rng.Characters.Last.Text = Chr(7) Then rng.End = rng.End - 1
            End If

            snippet = CleanPara(rng.Text)
            If Len(snippet) = 0 Then GoTo NextP

            pg = 0
            On Error Resume Next
            pg = p.Range.Information(wdActiveEndAdjustedPageNumber)
            On Error GoTo 0

            outRng.InsertAfter _
                "Page: " & IIf(pg = 0, "?", CStr(pg)) & " | Style: " & CStr(p.Range.Style) & vbCrLf & _
                snippet & vbCrLf & _
                "----------------------------------------" & vbCrLf & vbCrLf

            outRng.Collapse wdCollapseEnd
            copied = copied + 1

        End If

NextP:
    Next p

    Application.StatusBar = False
    Application.DisplayAlerts = wdAlertsAll
    Application.ScreenUpdating = True

    MsgBox "Done! Extracted " & copied & " items.", vbInformation

End Sub


Private Function CleanPara(ByVal s As String) As String

    Dim t As String
    t = s

    t = Replace(t, vbCr, " ")
    t = Replace(t, Chr(7), " ")
    t = Replace(t, vbTab, " ")
    t = Replace(t, Chr(160), " ")

    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop

    CleanPara = Trim$(t)

End Function`;

  navigator.clipboard.writeText(macro).then(() => {
    showSnackbar("Macro copied!");
  });
}

// ============================================
// Detect Negative Questions and Extract to New Document Macro
// ============================================

function DetectNegativewordsinSQTQ() {
  const macro = `Option Explicit

Sub ExtractNegativeQuestions_ToNewDoc_Tolerant()

    Dim p As Paragraph
    Dim txt As String, tLow As String
    Dim inBlock As Boolean

    Dim negWords As Variant
    Dim report As String
    Dim count As Long
    Dim repDoc As Document

    Dim qText As String

    negWords = Array( _
        "false", "incorrect", "not", "except", "never", _
        "no", "none", "neither", "wrong", "without", "cannot", _
        "can't", "dont", "don't", "doesnt", "doesn't", "isnt", "isn't", _
        "wont", "won't", "shouldnt", "shouldn't", "mustnt", "mustn't" _
    )

    report = "Questions Containing Negative Words" & vbCrLf & _
             "===================================" & vbCrLf & vbCrLf

    count = 0
    inBlock = False
    qText = ""

    Application.ScreenUpdating = False
    Application.StatusBar = "Extracting negative-word questions..."

    For Each p In ActiveDocument.Paragraphs

        txt = Trim(Replace(p.Range.Text, Chr(13), ""))
        tLow = LCase$(txt)

        If InStr(tLow, "study question") > 0 Or InStr(tLow, "test question") > 0 Then
            inBlock = True
            qText = ""
            GoTo NextP
        End If

        If InStr(tLow, "answer table") > 0 Or Left$(Trim$(tLow), 6) = "answer" Then
            If Len(Trim$(qText)) > 0 Then
                If ContainsNegFlexible(qText, negWords) Then
                    count = count + 1
                    report = report & count & ". " & Trim$(qText) & vbCrLf & _
                             "----------------------------------------" & vbCrLf
                End If
            End If
            inBlock = False
            qText = ""
            GoTo NextP
        End If

        If inBlock Then

            If LooksLikeQuestionStart(txt) Then

                If Len(Trim$(qText)) > 0 Then
                    If ContainsNegFlexible(qText, negWords) Then
                        count = count + 1
                        report = report & count & ". " & Trim$(qText) & vbCrLf & _
                                 "----------------------------------------" & vbCrLf
                    End If
                End If

                qText = txt

            Else
                If Len(Trim$(qText)) > 0 Then
                    If Len(txt) > 0 Then qText = qText & " " & txt
                End If
            End If

        End If

NextP:
    Next p

    If Len(Trim$(qText)) > 0 Then
        If ContainsNegFlexible(qText, negWords) Then
            count = count + 1
            report = report & count & ". " & Trim$(qText) & vbCrLf & _
                     "----------------------------------------" & vbCrLf
        End If
    End If

    Application.StatusBar = False
    Application.ScreenUpdating = True

    Set repDoc = Documents.Add
    repDoc.Content.Text = report & vbCrLf & "Total Questions Found: " & count

    MsgBox "Done! Extracted " & count & " questions into a new document.", vbInformation

End Sub


Private Function LooksLikeQuestionStart(ByVal s As String) As Boolean
    Dim t As String
    t = Trim$(s)

    If Len(t) = 0 Then
        LooksLikeQuestionStart = False
        Exit Function
    End If

    If t Like "[0-9]#.*" Then LooksLikeQuestionStart = True: Exit Function
    If t Like "[0-9]#)*" Then LooksLikeQuestionStart = True: Exit Function
    If t Like "[0-9]#-* " Then LooksLikeQuestionStart = True: Exit Function

    If UCase$(Left$(t, 1)) = "Q" Then
        If Mid$(t, 2, 1) Like "[0-9]" Then LooksLikeQuestionStart = True: Exit Function
    End If

    If Left$(t, 1) = "•" Or Left$(t, 1) = "-" Then LooksLikeQuestionStart = True: Exit Function

    If InStr(t, "?") > 0 Then LooksLikeQuestionStart = True: Exit Function

    LooksLikeQuestionStart = False
End Function


Private Function ContainsNegFlexible(ByVal q As String, ByVal negWords As Variant) As Boolean
    Dim t As String, i As Long

    t = LCase$(q)

    t = Replace(t, vbTab, " ")
    t = Replace(t, ".", " ")
    t = Replace(t, ",", " ")
    t = Replace(t, "?", " ")
    t = Replace(t, "!", " ")
    t = Replace(t, ":", " ")
    t = Replace(t, ";", " ")
    t = Replace(t, "(", " ")
    t = Replace(t, ")", " ")
    t = Replace(t, "[", " ")
    t = Replace(t, "]", " ")
    t = Replace(t, """", " ")
    t = Replace(t, "’", "'")

    t = Replace(t, "'", "")

    t = " " & t & " "

    For i = LBound(negWords) To UBound(negWords)
        If InStr(1, t, " " & Replace(negWords(i), "'", "") & " ", vbTextCompare) > 0 Then
            ContainsNegFlexible = True
            Exit Function
        End If
    Next i

    ContainsNegFlexible = False
End Function`;

  navigator.clipboard.writeText(macro).then(() => {
    showSnackbar("Macro copied!");
  });
}

// ============================================
// Hylighting All Hyperlinks in the Document and Creating a Report Macro
// ============================================

function DetectHyperlinks() {
  const macro = `Sub RemoveAllHyperlinks_AndCreateReport()

    Dim i As Long
    Dim linkText As String
    Dim linkAddr As String

    Dim reportText As String
    Dim newDoc As Document

    reportText = "Deleted Hyperlinks Report" & vbCrLf & _
                 "==========================" & vbCrLf & vbCrLf

    ' Loop backwards safely
    For i = ActiveDocument.Hyperlinks.Count To 1 Step -1

        linkText = ActiveDocument.Hyperlinks(i).TextToDisplay
        linkAddr = ActiveDocument.Hyperlinks(i).Address

        ' Save hyperlink details in report
        reportText = reportText & i & ". " & linkText & vbCrLf & _
                     "   URL: " & linkAddr & vbCrLf & _
                     "-----------------------------------" & vbCrLf

        ' Delete hyperlink but keep text
        ActiveDocument.Hyperlinks(i).Delete

    Next i

    ' Create new document for report
    Set newDoc = Documents.Add
    newDoc.Content.Text = reportText

    MsgBox "All hyperlinks removed!" & vbCrLf & _
           "A new report document was created with deleted link details."

End Sub`;

  navigator.clipboard.writeText(macro).then(() => {
    showSnackbar("Macro copied!");
  });
}
// ============================================
// DOC Macro Style Formating Add Colors to  DOCX
// ============================================

function HighlightingFormats() {
  const macro = `Option Explicit

Sub MarkBoldItalicUnderline_WholeDoc_LightMode()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    Application.StatusBar = "Marking Bold/Italic/Underline..."

    MarkInAllStories ActiveDocument
    MarkInAllTextBoxes ActiveDocument

    Application.StatusBar = False
    Application.DisplayAlerts = wdAlertsAll
    Application.ScreenUpdating = True

    MsgBox "Done!" & vbCrLf & _
           "Red = Bold+Italic" & vbCrLf & _
           "Blue = Bold only" & vbCrLf & _
           "Purple = Italic only" & vbCrLf & _
           "Green = Underline only" & vbCrLf & _
           "Orange = Bold+Underline" & vbCrLf & _
           "Teal = Italic+Underline" & vbCrLf & _
           "Black = Bold+Italic+Underline"

End Sub


Private Sub MarkInAllStories(ByVal doc As Document)

    Dim s As Range
    For Each s In doc.StoryRanges
        MarkInRangeWords s

        Do While Not s.NextStoryRange Is Nothing
            Set s = s.NextStoryRange
            MarkInRangeWords s
        Loop
    Next s

End Sub


Private Sub MarkInAllTextBoxes(ByVal doc As Document)

    Dim shp As Shape
    For Each shp In doc.Shapes
        If shp.TextFrame.HasText Then
            MarkInRangeWords shp.TextFrame.TextRange
        End If
    Next shp

End Sub


Private Sub MarkInRangeWords(ByVal rng As Range)

    Dim w As Range
    Dim t As String
    Dim isB As Boolean, isI As Boolean, isU As Boolean

    For Each w In rng.Words

        t = w.Text
        t = Replace(t, Chr(13), "")
        t = Replace(t, Chr(7), "")
        t = Trim$(t)

        If Len(t) = 0 Then GoTo NextW

        isB = (w.Font.Bold = True)
        isI = (w.Font.Italic = True)
        isU = (w.Font.Underline <> wdUnderlineNone)

        If isB And isI And isU Then
            w.Font.Color = RGB(0, 0, 0)
        ElseIf isB And isI Then
            w.Font.Color = RGB(255, 0, 0)
        ElseIf isB And isU Then
            w.Font.Color = RGB(255, 140, 0)
        ElseIf isI And isU Then
            w.Font.Color = RGB(0, 128, 128)
        ElseIf isB Then
            w.Font.Color = RGB(0, 0, 255)
        ElseIf isI Then
            w.Font.Color = RGB(128, 0, 128)
        ElseIf isU Then
            w.Font.Color = RGB(0, 128, 0)
        End If

NextW:
        DoEvents
    Next w

End Sub`;

  navigator.clipboard.writeText(macro).then(() => {
    showSnackbar("Macro copied!");
  });
}

// ============================================
// DOC Macro XML to  DOCX Conversion Macro
// ============================================

function copyMacroXmltoDoc() {
  const macro = `Option Explicit

Sub BatchConvertXMLtoDOCX_ClearOutput_DeleteXML_Reliable()

    ' --- Rename folders (if old names exist) ---
    Dim baseDesktop As String
    baseDesktop = "C:\\Users\\KadavakutiRavichandr\\Desktop\\"

    If Dir(baseDesktop & "XML to DOCS", vbDirectory) = "" Then
        If Dir(baseDesktop & "New folder", vbDirectory) <> "" Then
            Name baseDesktop & "New folder" As baseDesktop & "XML to DOCS"
        End If
    End If

    If Dir(baseDesktop & "XML to DOCS\\DOCS", vbDirectory) = "" Then
        If Dir(baseDesktop & "XML to DOCS\\output", vbDirectory) <> "" Then
            Name baseDesktop & "XML to DOCS\\output" As baseDesktop & "XML to DOCS\\DOCS"
        End If
    End If

    Dim folderPath As String, outputPath As String
    Dim outFile As String
    Dim xmlName As String
    Dim xmlFiles() As String
    Dim n As Long, i As Long

    Dim xmlFullPath As String, docxFullPath As String
    Dim doc As Document

    Dim converted As Long, deleted As Long, failed As Long
    Dim failedList As String

    folderPath = baseDesktop & "XML to DOCS\\"
    outputPath = folderPath & "DOCS\\"

    converted = 0: deleted = 0: failed = 0: failedList = ""

    If Dir(outputPath, vbDirectory) = "" Then MkDir outputPath

    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone

    outFile = Dir(outputPath & "*.docx")
    Do While outFile <> ""
        On Error Resume Next
        SetAttr outputPath & outFile, vbNormal
        Kill outputPath & outFile
        On Error GoTo 0
        outFile = Dir()
    Loop

    n = 0
    xmlName = Dir(folderPath & "*.xml")
    Do While xmlName <> ""
        n = n + 1
        ReDim Preserve xmlFiles(1 To n)
        xmlFiles(n) = xmlName
        xmlName = Dir()
    Loop

    If n = 0 Then
        Application.DisplayAlerts = wdAlertsAll
        Application.ScreenUpdating = True
        MsgBox "No XML files found in: " & folderPath
        Exit Sub
    End If

    For i = 1 To n

        xmlFullPath = folderPath & xmlFiles(i)
        docxFullPath = outputPath & Replace(xmlFiles(i), ".xml", ".docx")

        On Error Resume Next
        Set doc = Documents.Open(FileName:=xmlFullPath, ReadOnly:=True, AddToRecentFiles:=False)

        If Err.Number <> 0 Or doc Is Nothing Then
            failed = failed + 1
            failedList = failedList & "- Open failed: " & xmlFiles(i) & vbCrLf
            Err.Clear
            On Error GoTo 0
            GoTo NextI
        End If
        On Error GoTo 0

        On Error Resume Next
        doc.SaveAs2 FileName:=docxFullPath, FileFormat:=wdFormatXMLDocument
        doc.Close SaveChanges:=False
        Set doc = Nothing

        If Err.Number <> 0 Then
            failed = failed + 1
            failedList = failedList & "- SaveAs failed: " & xmlFiles(i) & vbCrLf
            Err.Clear
            On Error GoTo 0
            GoTo NextI
        End If
        On Error GoTo 0

        converted = converted + 1

        If Dir(docxFullPath) <> "" Then
            On Error Resume Next
            SetAttr xmlFullPath, vbNormal
            Kill xmlFullPath
            If Err.Number = 0 Then
                deleted = deleted + 1
            Else
                failed = failed + 1
                failedList = failedList & "- Delete failed: " & xmlFiles(i) & vbCrLf
                Err.Clear
            End If
            On Error GoTo 0
        Else
            failed = failed + 1
            failedList = failedList & "- DOCX missing after save: " & xmlFiles(i) & vbCrLf
        End If

NextI:
    Next i

    Application.DisplayAlerts = wdAlertsAll
    Application.ScreenUpdating = True

    MsgBox "Done!" & vbCrLf & _
           "XML found: " & n & vbCrLf & _
           "DOCX created: " & converted & vbCrLf & _
           "XML deleted: " & deleted & vbCrLf & _
           "Failed: " & failed & vbCrLf & vbCrLf & _
           IIf(failed > 0, "Failures:" & vbCrLf & failedList, "No failures.")

End Sub`;

  navigator.clipboard.writeText(macro).then(() => {
    showSnackbar("Macro copied!");
  });
}

// ============================================
// Navigation and Content Management
// ============================================

function showContent(sectionId) {
  // Hide all content sections
  const sections = document.querySelectorAll(".content-section");
  sections.forEach((section) => section.classList.remove("active"));

  // Show selected section
  const section = document.getElementById(sectionId);
  if (section) {
    section.classList.add("active");
    section.scrollIntoView({ behavior: "smooth", block: "start" });
  }
}

// Navigation function handlers
function p1function() {
  showContent("instruction");
}
function p3function() {
  showContent("imagetag");
}
function p4function() {
  showContent("raptag");
}
function p5function() {
  showContent("scot");
}
function p6function() {
  showContent("audiobuttag");
}
function p7function() {
  showContent("videotag");
}
function p8function() {
  showContent("elinks");
}
function p9function() {
  showContent("entities");
}
function p10function() {
  showContent("reviewpt");
}
function p11function() {
  showContent("finalreview");
}
function p12function() {
  showContent("processsteps");
}
function p13function() {
  showContent("kiosk");
}
function p14function() {
  showContent("powvideo");
}
function p15function() {
  showContent("qms");
}
function p16function() {
  showContent("macros");
}
function p17function() {
  showContent("panel");
}
function p18function() {
  showContent("icons");
}

// ============================================
// Macro Files Copy Function
// ============================================

function copyMacroRef(filename) {
  copyToClipboard(filename);
  showSnackbar("Copied: " + filename);
}

// ============================================
// Icons Copy Function
// ============================================

function copyIconText(iconName) {
  copyToClipboard(iconName);
  showSnackbar("Copied: " + iconName);
}

// ============================================
// Text-to-Speech Functionality
// ============================================

function speekout() {
  const text = document.getElementById("myspeech").value;
  if (text.trim() === "") {
    showSnackbar("Please enter text to speak");
    return;
  }

  const utterance = new SpeechSynthesisUtterance(text);
  utterance.rate = 1;
  utterance.pitch = 1;
  utterance.volume = 1;

  speechSynthesis.speak(utterance);
}

// ============================================
// Copy to Clipboard Functionality
// ============================================

function copyToClipboard(text) {
  navigator.clipboard
    .writeText(text)
    .then(() => {
      showSnackbar("Text copied to clipboard!");
    })
    .catch((err) => {
      console.error("Failed to copy:", err);
      showSnackbar("Failed to copy text");
    });
}

// Entity copy functions
function fequals() {
  copyToClipboard("&equals;");
}
function fplus() {
  copyToClipboard("&plus;");
}
function fhellip4() {
  copyToClipboard("&hellip4;");
}
function fhellip() {
  copyToClipboard("&hellip;");
}
function fldquo() {
  copyToClipboard("&ldquo;");
}
function frdquo() {
  copyToClipboard("&rdquo;");
}
function flsquo() {
  copyToClipboard("&lsquo;");
}
function frsquo() {
  copyToClipboard("&rsquo;");
}
function fapos() {
  copyToClipboard("&apos;");
}
function fminus() {
  copyToClipboard("&minus;");
}
function ftimes() {
  copyToClipboard("&times;");
}
function fmdash() {
  copyToClipboard("&mdash;");
}
function fndash() {
  copyToClipboard("&ndash;");
}
function ffrac12() {
  copyToClipboard("&frac12;");
}
function ffrac14() {
  copyToClipboard("&frac14;");
}
function fsect11() {
  copyToClipboard("&sect;");
}
function fle() {
  copyToClipboard("&le;");
}
function fge() {
  copyToClipboard("&ge;");
}
function fpara() {
  copyToClipboard("&para;");
}
function fcent() {
  copyToClipboard("&cent;");
}

// Underline functions
function singleunderline() {
  copyToClipboard("<u>text</u>");
}
function doubleunderline() {
  copyToClipboard('<u style="text-decoration: underline double;">text</u>');
}

// ============================================
// Image Tagging Functions
// ============================================

function replaceImageRight() {
  const imgName = document.getElementById("imgname").value;
  const imgType = document.getElementById("imgtype").value;

  if (!imgName) {
    showSnackbar("Please enter image name");
    return;
  }

  const tag = `<avc float="right"><avo avoref="${imgName}.${imgType}" align="center"/></avc>`;
  document.getElementById("imgtag").innerHTML = `<pre>${escapeHtml(tag)}</pre>`;
  copyToClipboard(tag);
}

function replaceImageCenter() {
  const imgName = document.getElementById("imgname").value;
  const imgType = document.getElementById("imgtype").value;

  if (!imgName) {
    showSnackbar("Please enter image name");
    return;
  }

  const tag = `<avc><avo avoref="${imgName}.${imgType}" align="center"/></avc>`;
  document.getElementById("imgtag").innerHTML = `<pre>${escapeHtml(tag)}</pre>`;
  copyToClipboard(tag);
}

function replaceImageLeft() {
  const imgName = document.getElementById("imgname").value;
  const imgType = document.getElementById("imgtype").value;

  if (!imgName) {
    showSnackbar("Please enter image name");
    return;
  }

  const tag = `<avc><avo avoref="${imgName}.${imgType}" align="left"/></avc>`;
  document.getElementById("imgtag").innerHTML = `<pre>${escapeHtml(tag)}</pre>`;
  copyToClipboard(tag);
}

// ============================================
// Raptivity Tagging
// ============================================

function replaceEleTag() {
  const eleName = document.getElementById("elename").value;

  if (!eleName) {
    showSnackbar("Please enter element name");
    return;
  }

  const tag = `<avc><avo avoref="${eleName}.swf" avohtmlraptivityref="${eleName}" avohtmlraptivitylaunchref="${eleName}.html" htmlraptivityheight="500" htmlraptivitywidth="715" align="center"/><printOnly></printOnly></avc>`;

  document.getElementById("eletag").textContent = tag;

  copyToClipboard(tag);
} // ============================================
// Audio Tagging
// ============================================

function replaceAudTag() {
  const audName = document.getElementById("audname").value;
  const audFName = document.getElementById("audfname").value;

  if (!audFName) {
    showSnackbar("Please enter the audio name");
    return;
  }

  const imgRef = audName ? `${audName}.jpg` : "audio-icon.jpg";
  const audioRef = `${audFName}.mp3`;

  const tag = `<avc><avo avoref="${imgRef}" align="center"/><avo avoref="${audioRef}" isAudioControlVisible="true" audioControlWidth="400" align="center"/></avc><transcript></transcript>`;

  // ✅ correct
  document.getElementById("audtag").textContent = tag;

  copyToClipboard(tag);
} // ============================================
// Video Tagging
// ============================================

function replaceVidTag() {
  const vidInput = document.getElementById("vidname");
  const imgInput = document.getElementById("vimgname");

  let vidName = vidInput.value.trim();
  let imgName = imgInput.value.trim();

  // If one is empty, copy from the other
  if (vidName && !imgName) {
    imgInput.value = vidName;
    imgName = vidName;
  } else if (imgName && !vidName) {
    vidInput.value = imgName;
    vidName = imgName;
  }

  // If both empty
  if (!vidName && !imgName) {
    showSnackbar("Please enter video or image name");
    return;
  }

  // Use one base name
  const baseName = vidName || imgName;

  const tag = `<avc use-in-jaws="yes"><avo avoref="${baseName}.flv,${baseName}.mp4" avoimageref="${baseName}.jpg"/></avc>`;

  // Display
  document.getElementById("vidtag").textContent = tag;

  // Copy
  copyToClipboard(tag);

  // Snackbar
  showSnackbar("Tag copied!");
}
// ============================================
// SCOT Commands
// ============================================
// ================= COPY FUNCTION =================
function showOutput(id, text) {
  const el = document.getElementById(id);

  if (!el) {
    console.error("❌ Output element not found:", id);
    return;
  }

  el.textContent = text;

  // auto copy
    navigator.clipboard.writeText(text).then(() => {
      showSnackbar("Command copied to clipboard!");
    console.log("Copied:", text);
  });
}

// ================= GET INPUT =================
function getAcc() {
  const el = document.getElementById("acc");

  if (!el) {
    console.error("❌ Input #acc not found");
    return null;
  }

  const value = el.value.trim(); // ✅ FIX

  if (value === "") {
    showSnackbar("Please enter course acronym");
    el.focus();
    return null;
  }

  return value;
}

// ================= COMMON BUTTON HANDLER =================
function runCommand(cmd) {
  const acc = getAcc();
  if (!acc) return;

  showOutput("scottag", `${cmd} ${acc}`);
}

// ================= MULTI COMMANDS =================
function replaceScotTag() {
  const acc = getAcc();
  if (!acc) return;

  const cmd = `mprf ${acc}; qcom ${acc}; qedit ${acc}; qedit -m ${acc}; qedit -n ${acc}; qpicfiles ${acc}; mtermlist ${acc}; qedit -g ${acc}; qtermtitles ${acc}; qterms ${acc}; qquest ${acc}; wqsr2 ${acc}; ux2w ${acc}; ux2wm ${acc}; ux2w -g ${acc}`;

  showOutput("scottag", cmd);
}

function ux2w() {
  const acc = getAcc();
  if (!acc) return;

  const cmd = `ux2w ${acc}; ux2wm ${acc}; ux2w -g ${acc}`;
  showOutput("scottag", cmd);
}

// ================= PREVIEW =================
function fpreview() {
  const screenId = document.getElementById("preview").value.trim();

  if (!screenId) {
    alert("Enter screen ID");
    return;
  }

  showOutput("previewtag", `preview -e ${screenId}`);
}

// ================= PANEL =================
function apwPanel() {
  showOutput("scottag", `<n-panel panel="apw"/>`);
}

function lpPanel() {
  showOutput("scottag", `<n-panel panel="lp"/>`);
}

// ================= TERM COMMAND =================
function termCmd() {
  const term = document.getElementById("termid").value.trim();

  if (!term) {
    alert("Enter term id");
    return;
  }

  showOutput("scottag", `e -t ${term}`);
}
function apwPanel() {
  showOutput("scottag", `<n-panel panel = "apw"/>`);
}

function lpPanel() {
  showOutput("scottag", `<n-panel panel = "lp"/>`);
}
function getTermCommand() {
  const term = document.getElementById("termid").value.trim();
  if (!term) {
    alert("Enter term ID");
    return;
  }

  showOutput("ettag", `e -t ${term}`);
} // ============================================
// Checklist Functions
// ============================================

function checkReset() {
  const checkboxes = document.querySelectorAll(
    '#processsteps input[type="checkbox"]',
  );
  checkboxes.forEach((checkbox) => (checkbox.checked = false));
  document.getElementById("accfi").value = "";
  document.getElementById("courseversion").value = "update course";
  document.getElementById("worktype").value = "Converted";
  document.getElementById("mailid").selectedIndex = 0;
  document.getElementById("finalmail").selectedIndex = 0;
  showSnackbar("Checklist reset");
}

function checkReset2() {
  const checkboxes = document.querySelectorAll(
    '#reviewpt input[type="checkbox"]',
  );
  checkboxes.forEach((checkbox) => (checkbox.checked = false));
  showSnackbar("Checklist reset");
}

function checkReset3() {
  const checkboxes = document.querySelectorAll(
    '#finalreview input[type="checkbox"]',
  );
  checkboxes.forEach((checkbox) => (checkbox.checked = false));
  showSnackbar("Checklist reset");
}

function sendMail() {
  const accfi = document.getElementById("accfi").value;
  const mailid = document.getElementById("mailid").value;

  if (!accfi || !mailid) {
    showSnackbar("Please fill in both course acronym and QA person");
    return;
  }

  showSnackbar("Email would be sent to: " + mailid);
  console.log("Send mail for course: " + accfi + " to: " + mailid);
}

function sendFinalMail() {
  const finalmail = document.getElementById("finalmail").value;

  if (!finalmail) {
    showSnackbar("Please select a TE");
    return;
  }

  showSnackbar("Email would be sent to: " + finalmail);
  console.log("Send final mail to: " + finalmail);
}

// ============================================
// Opportunity Calculator
// ============================================

function filterdata(event) {
  const oppcal = document.getElementById("oppcal").value;
  if (!oppcal.trim()) {
    showSnackbar("Please paste the mprf table");
    return;
  }
  showSnackbar("Data filtered and copied");
  console.log("Filter data:", oppcal);
}

function calopp(event) {
  const chapters =
    parseInt(document.getElementById("chapternumber").textContent) || 0;
  const propage = parseInt(document.getElementById("propage").textContent) || 0;
  const ppe = parseInt(document.getElementById("ppe").textContent) || 0;
  const testques =
    parseInt(document.getElementById("testques").textContent) || 0;
  const tqe = parseInt(document.getElementById("tqe").textContent) || 0;
  const suplele = parseInt(document.getElementById("suplele").textContent) || 0;
  const see = parseInt(document.getElementById("see").textContent) || 0;
  const termele = parseInt(document.getElementById("termele").textContent) || 0;
  const tee = parseInt(document.getElementById("tee").textContent) || 0;
  const metad = parseInt(document.getElementById("metad").textContent) || 4;
  const mde = parseInt(document.getElementById("mde").textContent) || 0;
  const scottt = parseInt(document.getElementById("scottt").textContent) || 13;
  const scote = parseInt(document.getElementById("scote").textContent) || 0;
  const intialsc =
    parseInt(document.getElementById("intialsc").textContent) || 4;
  const iee = parseInt(document.getElementById("iee").textContent) || 0;

  const totalOpp =
    chapters +
    propage +
    testques +
    suplele +
    termele +
    metad +
    scottt +
    intialsc;
  const totalError = ppe + tqe + see + tee + mde + scote + iee;
  const quality =
    totalOpp > 0 ? (((totalOpp - totalError) / totalOpp) * 100).toFixed(2) : 0;

  document.getElementById("totalopp").textContent = totalOpp;
  document.getElementById("totalerror").textContent = totalError;
  document.getElementById("qualitt").textContent = quality + "%";

  showSnackbar("Opportunity calculated: " + quality + "%");
}

// ============================================
// Instruction Functions
// ============================================

function handleSelectChange(event) {
  const selectedLevel = event.target.value;

  // Hide all sections
  document.getElementById("l2").style.display = "none";
  document.getElementById("l4").style.display = "none";

  let displayContent = "";
  let copyContent = "";

  // LEVEL 2
  if (selectedLevel === "level2") {
    displayContent = `<para><instruction>This course includes text, exercises, study questions, and a final exam. The following are examples of images that will display more information when selected. In addition, this course may provide links to view glossary terms, access Internet sites, and open supplements. The glossary and the supplements are also available in the resources section on the top right of the screen.</instruction> &nbsp;</para> <avc><avo avoref="instruction.jpg" align="center"/></avc>`;

    copyContent = `This course includes text, exercises, study questions, and a final exam. The following are examples of images that will display more information when selected. In addition, this course may provide links to view glossary terms, access Internet sites, and open supplements. The glossary and the supplements are also available in the resources section on the top right of the screen.`;

    document.getElementById("l2").style.display = "block";
    document.getElementById("level2").textContent = displayContent;
  }

  // LEVEL 4
  else if (selectedLevel === "level4") {
    displayContent = `<para><instruction>This course includes text, exercises, study questions, and a final exam. The following are examples of images that will display more information when selected. In addition, this course may provide links to view glossary terms, access Internet sites, and open supplements. The glossary and the supplements are also available in the resources section on the top right of the screen. If an activity has audio, it can be muted by selecting the speaker icon. You may select the transcript button on the top right of that particular screen to read it.</instruction> &nbsp;</para> <avc><avo avoref="instruction.jpg" align="center"/></avc>`;

    copyContent = `This course includes text, exercises, study questions, and a final exam. The following are examples of images that will display more information when selected. In addition, this course may provide links to view glossary terms, access Internet sites, and open supplements. The glossary and the supplements are also available in the resources section on the top right of the screen. If an activity has audio, it can be muted by selecting the speaker icon. You may select the transcript button on the top right of that particular screen to read it.`;

    document.getElementById("l4").style.display = "block";
    document.getElementById("level4").textContent = displayContent;
  }

  // Copy WITHOUT tags
  if (copyContent) {
    navigator.clipboard.writeText(copyContent);
    showSnackbar("Copied without tags!");
  }
}

// CLICK BOX → COPY AGAIN
function copyLevel(levelId) {
  let text = "";

  if (levelId === "level2") {
    text = document.getElementById("level2").textContent;
  } else if (levelId === "level4") {
    text = document.getElementById("level4").textContent;
  }

  // Remove tags before copying
  const cleanText = text.replace(/<[^>]*>/g, "");

  navigator.clipboard.writeText(cleanText);
  showSnackbar("Copied!");
}

// HIGHLIGHT EFFECT
function highlightElement(elementId) {
  const element = document.getElementById(elementId);
  element.style.backgroundColor = "#FFFF00";

  setTimeout(() => {
    element.style.backgroundColor = "";
  }, 1500);
}

// ============================================
// Utility Functions
// ============================================

function escapeHtml(text) {
  const map = {
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#039;",
  };
  return text.replace(/[&<>"']/g, (m) => map[m]);
}

function showSnackbar(message) {
  const snackbar = document.getElementById("snackbar");
  snackbar.textContent = message;
  snackbar.classList.add("show");
  setTimeout(() => {
    snackbar.classList.remove("show");
  }, 3000);
}

function getAuthorPacket() {
  const accs = document.getElementById("accs").value;
  if (!accs.trim()) {
    showSnackbar("Please enter course acronyms");
    return;
  }
  const cmd = `Author Packet for: ${accs}`;
  document.getElementById("authpacket").innerHTML = `<pre>${cmd}</pre>`;
  copyToClipboard(cmd);
}

// // Additional SCOT command table functions
// function mprfTag() {
//   copyToClipboard("mprf");
//   showSnackbar("Command copied");
// }
// function qeditnTag() {
//   copyToClipboard("qedit -n");
//   showSnackbar("Command copied");
// }
// function qeditgTag() {
//   copyToClipboard("qedit -g");
//   showSnackbar("Command copied");
// }
// function wqsr2Tag() {
//   copyToClipboard("wqsr2");
//   showSnackbar("Command copied");
// }
// function ux2wTag() {
//   copyToClipboard("ux2w");
//   showSnackbar("Command copied");
// }
// function qeditmTag() {
//   copyToClipboard("qedit -m");
//   showSnackbar("Command copied");
// }
// function mtermlistTag() {
//   copyToClipboard("mtermlist");
//   showSnackbar("Command copied");
// }
// function qtermtitlesTag() {
//   copyToClipboard("qtermtitles");
//   showSnackbar("Command copied");
// }
// function qpicfilesTag() {
//   copyToClipboard("qpicfiles");
//   showSnackbar("Command copied");
// }
// function ux2wmTag() {
//   copyToClipboard("ux2wm");
//   showSnackbar("Command copied");
// }
// function qeditTag() {
//   copyToClipboard("qedit");
//   showSnackbar("Command copied");
// }
// function qtermsTag() {
//   copyToClipboard("qterms");
//   showSnackbar("Command copied");
// }
// function qquestTag() {
//   copyToClipboard("qquest");
//   showSnackbar("Command copied");
// }
// function qcomTag() {
//   copyToClipboard("qcom");
//   showSnackbar("Command copied");
// }
// function ux2wgTag() {
//   copyToClipboard("ux2w -g");
//   showSnackbar("Command copied");
// }

// ============================================
// Initialize on Page Load
// ============================================

function setupKeyupListener(inputId, buttonId) {
  const input = document.getElementById(inputId);
  if (input) {
    input.addEventListener("keyup", (e) => {
      if (e.keyCode === 13) {
        const button = document.getElementById(buttonId);
        if (button) button.click();
      }
    });
  }
}

document.addEventListener("DOMContentLoaded", function () {
  // Set home page as active by default
  showContent("elinks");
  const homeLink = document.getElementById("a1");
  if (homeLink) homeLink.classList.add("active");

  // Set up navigation active states
  document.querySelectorAll(".nav-link").forEach((link) => {
    link.addEventListener("click", function (e) {
      document
        .querySelectorAll(".nav-link")
        .forEach((l) => l.classList.remove("active"));
      this.classList.add("active");
    });
  });

  // Setup keyup listeners for form inputs
  setupKeyupListener("myspeech", "listenbutton");
  setupKeyupListener("imgname", "imgbutton");
  setupKeyupListener("elename", "elebutton");
  setupKeyupListener("acc", "scotcmdbutton");
  setupKeyupListener("audname", "audbutton");
  setupKeyupListener("audfname", "audbutton");
  setupKeyupListener("vimgname", "vidbutton");
  setupKeyupListener("vidname", "vidbutton");
  setupKeyupListener("ppe", "calcbutton");
  setupKeyupListener("tqe", "calcbutton");
  setupKeyupListener("see", "calcbutton");
  setupKeyupListener("tee", "calcbutton");
  setupKeyupListener("mde", "calcbutton");
  setupKeyupListener("scote", "calcbutton");
  setupKeyupListener("iee", "calcbutton");
  setupKeyupListener("totalerror", "calcbutton");

  // Attach listen button functionality
  const listenBtn = document.getElementById("listenbutton");
  if (listenBtn) {
    listenBtn.addEventListener("click", speekout);
  }

  console.log("CPOE Online Team - Application Initialized");
});

// Mychecklist Copy Icon Functionality

document.addEventListener("DOMContentLoaded", function () {
  document.querySelectorAll(".copy-list li").forEach((li) => {
    const icon = document.createElement("span");
    icon.className = "copy-icon";
    icon.innerText = "📋";

    li.appendChild(icon);

    icon.addEventListener("click", function (e) {
      e.stopPropagation();

      const value = li.dataset.copy || li.innerText;

      navigator.clipboard.writeText(value).then(() => {
        // Optional: small visual feedback instead of alert
        icon.innerText = "✔";
        setTimeout(() => (icon.innerText = "📋"), 1000);
      });
    });
  });
});
