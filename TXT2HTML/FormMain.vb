' --------------------------------
' --- FormMain.vb - 02/02/2016 ---
' --------------------------------

' ----------------------------------------------------------------------------------------------------
' 02/02/2016 - SBakker
'            - Added "%" for outdenting paragraph.
' 10/11/2015 - SBakker
'            - Added support for "]" making lines right justified.
' 08/16/2015 - SBakker
'            - Fixed "Can't Copy" error to throw an error and stop.
' 06/18/2015 - SBakker
'            - ...And another go at making "***" look good...
' 04/05/2015 - SBakker
'            - Changed "***" to <hr> so it looks better.
' 12/20/2014 - SBakker
'            - Check for "***" after removing all spaces and tabs, just in case.
' 11/26/2014 - SBakker
'            - Only include "+" as a subsection if the original line starts with it. If it is indented
'              then it is just a plain old line.
' 11/25/2014 - SBakker
'            - Changed <h4> "+" subsections to display as <h3>, since they should be the same size.
' 07/17/2014 - SBakker
'            - Create a new function to fix table rows for tabs, &nbsp;, and additional spaces.
' 07/11/2014 - SBakker
'            - Check for either "^", "^^" or vbTab+"^", vbTab+"^^" and handle them equally.
' 07/08/2014 - SBakker
'            - Handle "<hr />" horizontal lines properly.
' 06/03/2014 - SBakker
'            - Fix spacer lines with only a Null (&#0;).
' 05/24/2014 - SBakker
'            - Added ButtonStop to stop safely in the middle.
' 05/17/2014 - SBakker
'            - Fixed typo with &#0; handling.
' 05/13/2014 - SBakker
'            - Remove any HTML Null &#0; characters. They are great for ignoring bad formatting, like
'              partial-word italics: <i>some</i>&#0;body
'            - Added "quoteblockflush" for lines beginning with vbTab + vbTab + "|".
' 03/26/2014 - SBakker
'            - Changed "* * *" to be "<br />* * *<br />", so the lines take exactly the same space.
' 03/15/2014 - SBakker
'            - Added Bootstrapping to a a local area instead of using the ClickOnce install.
'            - Added the Settings Provider to save settings in the local area with the program.
'            - Added AboutMain.vb which shows the current path in the Status Bar.
' 01/24/2014 - SBakker
'            - Changed scenebreak in css to have no extra space above or below. Instead, add a blank
'              line above and below. This will keep all the lines even from page to page. Before, a
'              scenebreak would mess up the line positions from the pages before and after.
'            - Added a ChangedCount to show how many HTML files were updated.
' 12/08/2013 - SBakker
'            - Took out special handling of <pre> tag. Switched all books to use <code> tag, which
'              looks better anyway.
' 10/23/2013 - SBakker
'            - Added &nbsp; between single-quote and double-quote, so they don't get spread too far
'              apart.
' 07/28/2013 - SBakker
'            - Added StatusBarMain to handle messages, so there's no popup message box.
' 04/21/2013 - SBakker
'            - Remove blank lines at end of chapters before next chapter starts.
' 03/10/2013 - SBakker
'            - Added blank lines between items of header line. More like a book title page.
' 01/27/2013 - SBakker
'            - Added blank line before and after "* * *".
' 08/16/2012 - SBakker
'            - If images have to be copied, make sure that the HTML file is rebuilt too.
' 07/11/2012 - SBakker
'            - FixInlineImages in Blockquote, "^", "^^", and "|" sections too.
' 06/11/2012 - SBakker
'            - Made lines starting with "^^" be centered and bold and big.
' 06/11/2012 - SBakker
'            - Made lines starting with "^" be centered and bold.
' 04/21/2012 - SBakker
'            - Made CurrLine.StartsWith(":") be changed to an <h6> heading, which has
'              "display: none;", yet is still in the Table of Contents. For when the chapter
'              starts with an image.
' 04/15/2012 - SBakker
'            - Added handling of \" (escaped quote) to change it to a normal quote. These
'              are escaped so they are ignored by quote checking in BookCleaner.
' 04/14/2012 - SBakker
'            - Added handling for lines staring with vbTab + "|" as having no indent.
' 04/06/2012 - SBakker
'            - Added AboutMain.vb form to this project.
' 02/26/2012 - SBakker
'            - Remove trailing "·" used to mark an OK but odd line ending.
' 01/06/2012 - Added check for "<overline>" to convert it to a proper inline style tag.
' 12/14/2011 - Added check for "<pre>" tag, and add CSS formatting only if it is found.
'            - Added TextBoxResults to show which files were changed.
' 12/02/2011 - Changed "***" to "* * *" centered using <p class="break">.
'            - Added p.break to center section breaks.
' 10/03/2011 - Changed to use better formatting, centering on headers/chapters, and FINALLY
'              fixed "start" id.
' 08/24/2011 - SBakker
'            - Changed default encoding to be whatever, typically ASCII. Needs to have
'              better file encoding checker whenever a BOM isn't found!
' 08/23/2011 - SBakker
'            - Added settings.
' 08/15/2011 - SBakker
'            - Fixed images with quotes around them.
'            - Allow HTML tags to remain unchanged.
' 08/03/2011 - SBakker
'            - Added default file encoding of UTF-8, in case it isn't possible to determine
'              from a byte-order-mark (BOM).
' 06/25/2011 - SBakker
'            - Use middot for nbsp instead of "&middot;".
' 06/16/2011 - SBakker
'            - Leave leading and trailing spaces around emdashs alone.
' 06/12/2011 - SBakker
'            - Preserve italics, bold, and underline HTML code when found.
' 06/10/2011 - SBakker
'            - Changed formatting, trying to get chapter titles and table of contents
'              working.
' 05/20/2011 - SBakker
'            - Changed to a formatting more in line with Kindle books.
' 01/29/2011 - SBakker
'            - Fixed title of HTM document to not have the file extension.
' 01/26/2011 - SBakker
'            - Check "Not StartsWithTab" to see if a line is a chapter heading.
' 01/20/2011 - SBakker
'            - Added <h2> tags to first line of the file, and <h3> tags around each chapter
'              heading.
' ----------------------------------------------------------------------------------------------------

Imports Arena_Utilities.FileUtils
Imports System.IO
Imports System.Text

Public Class FormMain

    Private Shared ReadOnly ObjName As String = System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.FullName

    Private FromPath As String = ""
    Private ToPath As String = ""
    Private StopRequested As Boolean = False

    Private Sub FormMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name

        Try
            If Arena_Bootstrap.BootstrapClass.CopyProgramsToLaunchPath Then
                Me.Close()
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show(FuncName + vbCrLf + ex.Message, My.Application.Info.AssemblyName, MessageBoxButtons.OK)
            Me.Close()
            Exit Sub
        End Try

        ' --- First call Upgrade to load setting from last version ---
        If My.Settings.CallUpgrade Then
            My.Settings.Upgrade()
            My.Settings.CallUpgrade = False
            My.Settings.Save()
        End If

        TextBoxFromPath.Text = My.Settings.DefaultFromPath
        TextBoxToPath.Text = My.Settings.DefaultToPath

    End Sub

    Private Sub ButtonStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonStart.Click
        Dim FileCount As Integer = 0
        Dim ChangedCount As Integer = 0
        Const FileCountMsg As String = "Files Found: "
        Const ChangedCountMsg As String = " - Files Changed: "
        ' ----------------------------------------------------
        If TextBoxFromPath.Text = "" Then Exit Sub
        If Not Directory.Exists(TextBoxFromPath.Text) Then Exit Sub
        If TextBoxToPath.Text = "" Then Exit Sub
        If Not Directory.Exists(TextBoxToPath.Text) Then Exit Sub
        ' --- Begin ---
        StopRequested = False
        ButtonStop.Enabled = True
        ButtonStop.Visible = True
        ButtonStart.Enabled = False
        ButtonStart.Visible = False
        ButtonStop.Focus()
        My.Settings.DefaultFromPath = TextBoxFromPath.Text
        My.Settings.DefaultToPath = TextBoxToPath.Text
        My.Settings.Save()
        TextBoxResults.Text = ""
        TextBoxFromPath.ReadOnly = True
        TextBoxToPath.ReadOnly = True
        FileCount = 0
        ChangedCount = 0
        ToolStripStatusLabelMain.Text = FileCountMsg + FileCount.ToString + ChangedCountMsg + ChangedCount.ToString
        Application.DoEvents()
        Dim FromFiles() As String = Directory.GetFiles(TextBoxFromPath.Text, "*.txt")
        ToPath = TextBoxToPath.Text
        For Each FileName As String In FromFiles
            If StopRequested Then Exit For
            FileCount += 1
            FromPath = FileName.Substring(0, FileName.LastIndexOf("\"c))
            Try
                If ConvertToHTM(FileName) Then
                    ChangedCount += 1
                End If
            Catch ex As Exception
                MessageBox.Show("Error converting file: " + ex.Message)
                Exit For
            End Try
            ToolStripStatusLabelMain.Text = FileCountMsg + FileCount.ToString + ChangedCountMsg + ChangedCount.ToString
            Application.DoEvents()
        Next
        ' --- Done ---
        ToolStripStatusLabelMain.Text += " - Done"
        TextBoxFromPath.ReadOnly = False
        TextBoxToPath.ReadOnly = False
        ButtonStart.Enabled = True
        ButtonStart.Visible = True
        ButtonStop.Enabled = False
        ButtonStop.Visible = False
        ButtonStart.Focus()
    End Sub

    Private Function ConvertToHTM(ByVal FileName As String) As Boolean
        Dim BaseFileName As String = FileName.Substring(FileName.LastIndexOf("\"c) + 1)
        Dim BaseFileNameNoExt As String = BaseFileName.Substring(0, BaseFileName.LastIndexOf("."c))
        Dim TargetText As New StringBuilder
        Dim FirstLine As Boolean = True
        Dim StartsWithTab As Boolean = False
        Dim StartsWithTwoTabs As Boolean = False
        Dim TempCurrLine As String
        ''Dim NoIndentTag As Boolean = False
        Dim ImagesChanged As Boolean = False
        Dim BlankLineCount As Integer = 0
        ' ---------------------------------------
        TextBoxFileName.Text = BaseFileName
        With TargetText
            ' --- Build heading ---
            .AppendLine("<html>")
            .AppendLine("<head>")
            '' --- May be needed later ---
            ''.AppendLine("<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">")
            ' --- Add in all metadata ---
            FirstLine = True
            For Each CurrLine As String In File.ReadAllLines(FileName, GetFileEncoding(FileName))
                If CurrLine Is Nothing Then Continue For
                If CurrLine = "&#0;" Then Continue For
                If Not FirstLine AndAlso CurrLine.StartsWith("<") AndAlso Not CurrLine.ToLower.StartsWith("<image=") Then
                    .AppendLine(CurrLine)
                End If
                ''If Not NoIndentTag AndAlso
                ''    (CurrLine.StartsWith("|") OrElse CurrLine.StartsWith(vbTab + "|")) AndAlso
                ''    Not CurrLine.StartsWith(vbTab + "||") Then
                ''    NoIndentTag = True
                ''End If
                FirstLine = False
            Next
            .AppendLine("<link href=""css\ebookstyle.css"" rel=""stylesheet"" type=""text/css"">")
            .AppendLine("</head>")
            .AppendLine("<body>")
            ' --- Fill in the lines ---
            FirstLine = True
            BlankLineCount = 0
            For Each CurrLine As String In File.ReadAllLines(FileName, GetFileEncoding(FileName))
                If CurrLine Is Nothing Then Continue For
                If CurrLine = "&#0;" Then
                    CurrLine = ""
                End If
                If Not FirstLine AndAlso CurrLine.StartsWith("<") AndAlso Not CurrLine.ToLower.StartsWith("<image=") Then
                    Continue For
                End If
                StartsWithTab = CurrLine.StartsWith(vbTab)
                StartsWithTwoTabs = CurrLine.StartsWith(vbTab + vbTab)
                TempCurrLine = FixLine(CurrLine)
                ' --- Split first line into multiple parts ---
                If FirstLine Then
                    Dim HeadingLines() As String = TempCurrLine.Split("—"c)
                    For Each TempHeading As String In HeadingLines
                        If FirstLine Then
                            TempHeading = "<div id=""start""><h1><u>" + TempHeading.Trim + "</u></h1></div>" ' --- Book title ---
                        Else
                            .AppendLine("<p>&nbsp;</p>")
                            TempHeading = "<h2>" + TempHeading.Trim + "</h2>"
                        End If
                        TempHeading = FixTags(TempHeading)
                        .AppendLine(TempHeading)
                        FirstLine = False
                    Next
                    BlankLineCount = 0
                ElseIf CurrLine = "" Then ' --- Must fill in blank lines or they get squished out ---
                    BlankLineCount += 1 ' just count them for now
                ElseIf CurrLine.StartsWith("<image=") Then
                    If CurrLine.Contains("""") Then CurrLine = CurrLine.Replace("""", "")
                    If CurrLine.Contains("—") Then CurrLine = CurrLine.Replace("—", "-") ' fix any bad file names
                    If CurrLine.Contains("  ") Then CurrLine = CurrLine.Replace("  ", " ") ' fix any bad file names
                    Do While BlankLineCount > 0
                        .AppendLine("<p>&nbsp;</p>")
                        BlankLineCount -= 1
                    Loop
                    Dim CurrImageName As String = CurrLine.Replace("<image=", "").Replace(">", "")
                    Dim NewImageName As String = CurrLine.Replace("<image=", "").Replace(">", "")
                    CurrLine = CurrLine.Replace(">", """>").Replace("<image=", "<div style=""text-align:center""><img src=""") + "</img></div>"
                    .AppendLine(CurrLine)
                    Try
                        If Not File.Exists(ToPath + "\" + NewImageName) OrElse
                            File.GetLastWriteTimeUtc(FromPath + "\" + CurrImageName) < File.GetLastWriteTimeUtc(ToPath + "\" + NewImageName) Then
                            File.Copy(FromPath + "\" + CurrImageName, ToPath + "\" + NewImageName, True)
                            ImagesChanged = True
                        End If
                    Catch ex As Exception
                        Throw New SystemException("Can't copy file: " + CurrImageName + vbCrLf + ex.Message)
                    End Try
                ElseIf TempCurrLine.StartsWith("^^") Then ' --- Centered and Bold and Big ---
                    Do While BlankLineCount > 0
                        .AppendLine("<p>&nbsp;</p>")
                        BlankLineCount -= 1
                    Loop
                    .Append("<p class=""break""><b><big>")
                    .Append(FixInlineImages(TempCurrLine.Substring(2).Trim))
                    .AppendLine("</big></b></p>")
                ElseIf TempCurrLine.StartsWith("^") Then ' --- Centered and Bold ---
                    Do While BlankLineCount > 0
                        .AppendLine("<p>&nbsp;</p>")
                        BlankLineCount -= 1
                    Loop
                    .Append("<p class=""break""><b>")
                    .Append(FixInlineImages(TempCurrLine.Substring(1).Trim))
                    .AppendLine("</b></p>")
                ElseIf CurrLine.StartsWith("+") Then ' --- Sub-headings in Table of Contents ---
                    BlankLineCount = 0 ' erase all previous blank lines
                    .Append("<h3>")
                    .Append(TempCurrLine.Substring(1).Trim)
                    .AppendLine("</h3>")
                ElseIf TempCurrLine.StartsWith(":") AndAlso Not TempCurrLine.StartsWith("::") Then ' --- Hidden main heading, but still in Table of Contents (Usually followed by an image) ---
                    BlankLineCount = 0 ' erase all previous blank lines
                    .Append("<h6>")
                    .Append(TempCurrLine.Substring(1).Trim)
                    .AppendLine("</h6>")
                ElseIf TempCurrLine.StartsWith("|") AndAlso Not TempCurrLine.StartsWith("||") Then ' no indent
                    Do While BlankLineCount > 0
                        .AppendLine("<p>&nbsp;</p>")
                        BlankLineCount -= 1
                    Loop
                    .Append("<p class=""noindent"">")
                    .Append(FixInlineImages(TempCurrLine.Substring(1).Trim))
                    .AppendLine("</p>")
                ElseIf TempCurrLine.StartsWith("]") Then ' right-justify
                    Do While BlankLineCount > 0
                        .AppendLine("<p>&nbsp;</p>")
                        BlankLineCount -= 1
                    Loop
                    .Append("<p class=""blockright"">")
                    .Append(TempCurrLine.Substring(1))
                    .AppendLine("</p>")
                ElseIf TempCurrLine.StartsWith("%") Then ' outdent
                    Do While BlankLineCount > 0
                        .AppendLine("<p>&nbsp;</p>")
                        BlankLineCount -= 1
                    Loop
                    .Append("<p class=""outdent"">")
                    .Append(TempCurrLine.Substring(1))
                    .AppendLine("</p>")
                ElseIf Not StartsWithTab Then ' --- Headings in Table of Contents ---
                    BlankLineCount = 0 ' erase all previous blank lines
                    .Append("<h3>")
                    .Append(TempCurrLine)
                    .AppendLine("</h3>")
                ElseIf StartsWithTwoTabs Then ' --- Blockquotes ---
                    Do While BlankLineCount > 0
                        .AppendLine("<p>&nbsp;</p>")
                        BlankLineCount -= 1
                    Loop
                    If CurrLine.StartsWith(vbTab + vbTab + "|") Then
                        .Append("<p class=""quoteblockflush"">")
                    Else
                        .Append("<p class=""quoteblock"">")
                    End If
                    .Append(FixInlineImages(TempCurrLine).Replace("&nbsp;&nbsp;&nbsp;&nbsp;", ""))
                    .AppendLine("</p>")
                ElseIf CurrLine.Replace(" ", "").Replace(vbTab, "") = "***" Then
                    BlankLineCount = 0 ' erase all previous blank lines
                    ''.AppendLine("<hr class=""scenebreak"" />")
                    ''.AppendLine("<hr style=""width:20%;"" />")
                    ''.AppendLine("<p class=""scenebreak""><br />* * *<br /></p>")
                    ''.AppendLine("<p class=""scenebreak"">* * *</p>")
                    .AppendLine("<p>&nbsp;</p><p class=""scenebreak"">* * *</p><p>&nbsp;</p>")
                ElseIf CurrLine.ToLower.StartsWith(vbTab + "<t") OrElse CurrLine.ToLower.StartsWith(vbTab + "</t") Then ' tables
                    Do While BlankLineCount > 0
                        .AppendLine("<p>&nbsp;</p>")
                        BlankLineCount -= 1
                    Loop
                    .AppendLine(FixTableLine(CurrLine))
                ElseIf CurrLine.ToLower.Contains("<hr />") Then
                    Do While BlankLineCount > 0
                        .AppendLine("<p>&nbsp;</p>")
                        BlankLineCount -= 1
                    Loop
                    .AppendLine("<hr />")
                Else
                    Do While BlankLineCount > 0
                        .AppendLine("<p>&nbsp;</p>")
                        BlankLineCount -= 1
                    Loop
                    .Append("<p>")
                    .Append(FixInlineImages(TempCurrLine))
                    .AppendLine("</p>")
                End If
            Next
            ' --- Build ending ---
            .AppendLine("</body>")
            .AppendLine("</html>")
        End With
        ' --- Prepare the result ---
        Dim TargetFilename As String = ToPath + "\" + BaseFileNameNoExt + ".html"
        If File.Exists(TargetFilename) Then
            ' --- Check if file hasn't changed ---
            If Not ImagesChanged Then
                Dim OldFileText As String = File.ReadAllText(TargetFilename, GetFileEncoding(TargetFilename))
                If OldFileText = TargetText.ToString Then
                    Return False
                End If
            End If
            ' --- Check if file is read-only ---
            Dim CurrInfo As FileInfo = My.Computer.FileSystem.GetFileInfo(TargetFilename)
            Do While CurrInfo.IsReadOnly
                Dim Answer As DialogResult = MessageBox.Show("""" + TargetFilename + """ is Read-Only",
                                             "File is Read-Only", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error)
                ' --- abort ---
                If Answer = DialogResult.Abort Then
                    Return False
                End If
                ' --- ignore ---
                If Answer = DialogResult.Ignore Then
                    Return False
                End If
                ' --- retry ---
                CurrInfo = My.Computer.FileSystem.GetFileInfo(TargetFilename)
            Loop
        End If
        ' --- Write out result ---
        File.WriteAllText(TargetFilename, TargetText.ToString, Encoding.UTF8)
        TextBoxResults.AppendText(BaseFileName + vbCrLf)
        Return True
    End Function

    Private Function FixLine(ByVal CurrLine As String) As String
        If CurrLine = "" Then Return "&nbsp;"
        ' --- Remove trailing "·" used to mark an OK but odd line ending ---
        If CurrLine.EndsWith("·") Then CurrLine = CurrLine.Substring(0, CurrLine.Length - 1)
        If CurrLine.StartsWith(vbTab) AndAlso CurrLine.Length > 1 Then CurrLine = CurrLine.Substring(1) ' remove leading tab
        If CurrLine.Contains(vbTab) Then CurrLine = CurrLine.Replace(vbTab, "&nbsp;&nbsp;&nbsp;&nbsp;")
        If CurrLine.Contains(Chr(160)) Then CurrLine = CurrLine.Replace(Chr(160), "") ' non-breakable space
        If CurrLine.Contains(Chr(173)) Then CurrLine = CurrLine.Replace(Chr(173), "-"c) ' soft hyphen
        If CurrLine.Contains("\""") Then CurrLine = CurrLine.Replace("\""", """") ' an escaped quote, so it's ignored by quote checking
        If CurrLine.Contains(""" '") Then CurrLine = CurrLine.Replace(""" '", """&nbsp;'") ' keeps them from drifting apart
        If CurrLine.Contains("' """) Then CurrLine = CurrLine.Replace("' """, "'&nbsp;""") ' keeps them from drifting apart
        If CurrLine.Contains("<") AndAlso Not CurrLine.Contains("</") Then CurrLine = CurrLine.Replace("<", "&lt;")
        If CurrLine.Contains(">") AndAlso Not CurrLine.Contains("</") Then CurrLine = CurrLine.Replace(">", "&gt;")
        If CurrLine.Contains("€") Then CurrLine = CurrLine.Replace("€", "&euro;")
        If CurrLine.Contains("‚") Then CurrLine = CurrLine.Replace("‚", "&sbquo;")
        If CurrLine.Contains("ƒ") Then CurrLine = CurrLine.Replace("ƒ", "&fnof;")
        If CurrLine.Contains("„") Then CurrLine = CurrLine.Replace("„", "&bdquo;")
        If CurrLine.Contains("…") Then CurrLine = CurrLine.Replace("…", "&hellip;")
        If CurrLine.Contains("†") Then CurrLine = CurrLine.Replace("†", "&dagger;")
        If CurrLine.Contains("‡") Then CurrLine = CurrLine.Replace("‡", "&Dagger;")
        If CurrLine.Contains("ˆ") Then CurrLine = CurrLine.Replace("ˆ", "&circ;")
        If CurrLine.Contains("‰") Then CurrLine = CurrLine.Replace("‰", "&permil;")
        If CurrLine.Contains("Š") Then CurrLine = CurrLine.Replace("Š", "&Scaron;")
        If CurrLine.Contains("‹") Then CurrLine = CurrLine.Replace("‹", "&lsaquo;")
        If CurrLine.Contains("Œ") Then CurrLine = CurrLine.Replace("Œ", "&OElig;")
        If CurrLine.Contains("Ž") Then CurrLine = CurrLine.Replace("Ž", "&#0381;")
        If CurrLine.Contains("‘") Then CurrLine = CurrLine.Replace("‘", "&lsquo;")
        If CurrLine.Contains("’") Then CurrLine = CurrLine.Replace("’", "&rsquo;")
        If CurrLine.Contains(Chr(147)) Then CurrLine = CurrLine.Replace(Chr(147), "&ldquo;")
        If CurrLine.Contains(Chr(148)) Then CurrLine = CurrLine.Replace(Chr(148), "&rdquo;")
        If CurrLine.Contains("•") Then CurrLine = CurrLine.Replace("•", "&bull;")
        If CurrLine.Contains("–") Then CurrLine = CurrLine.Replace("–", "&#8211;")
        If CurrLine.Contains("˜") Then CurrLine = CurrLine.Replace("˜", "&tilde;")
        If CurrLine.Contains("™") Then CurrLine = CurrLine.Replace("™", "&trade;")
        If CurrLine.Contains("š") Then CurrLine = CurrLine.Replace("š", "&scaron;")
        If CurrLine.Contains("›") Then CurrLine = CurrLine.Replace("›", "&rsaquo;")
        If CurrLine.Contains("œ") Then CurrLine = CurrLine.Replace("œ", "&oelig;")
        If CurrLine.Contains("ž") Then CurrLine = CurrLine.Replace("ž", "&#0382;")
        If CurrLine.Contains("Ÿ") Then CurrLine = CurrLine.Replace("Ÿ", "&Yuml;")
        If CurrLine.Contains("¡") Then CurrLine = CurrLine.Replace("¡", "&iexcl;")
        If CurrLine.Contains("¢") Then CurrLine = CurrLine.Replace("¢", "&cent;")
        If CurrLine.Contains("£") Then CurrLine = CurrLine.Replace("£", "&pound;")
        If CurrLine.Contains("¤") Then CurrLine = CurrLine.Replace("¤", "&curren;")
        If CurrLine.Contains("¥") Then CurrLine = CurrLine.Replace("¥", "&yen;")
        If CurrLine.Contains("¦") Then CurrLine = CurrLine.Replace("¦", "&brvbar;")
        If CurrLine.Contains("§") Then CurrLine = CurrLine.Replace("§", "&sect;")
        If CurrLine.Contains("¨") Then CurrLine = CurrLine.Replace("¨", "&uml;")
        If CurrLine.Contains("©") Then CurrLine = CurrLine.Replace("©", "&copy;")
        If CurrLine.Contains("ª") Then CurrLine = CurrLine.Replace("ª", "&ordf;")
        If CurrLine.Contains("«") Then CurrLine = CurrLine.Replace("«", "&laquo;")
        If CurrLine.Contains("¬") Then CurrLine = CurrLine.Replace("¬", "&not;")
        If CurrLine.Contains("®") Then CurrLine = CurrLine.Replace("®", "&reg;")
        If CurrLine.Contains("¯") Then CurrLine = CurrLine.Replace("¯", "&macr;")
        If CurrLine.Contains("°") Then CurrLine = CurrLine.Replace("°", "&deg;")
        If CurrLine.Contains("±") Then CurrLine = CurrLine.Replace("±", "&plusmn;")
        If CurrLine.Contains("²") Then CurrLine = CurrLine.Replace("²", "&sup2;")
        If CurrLine.Contains("³") Then CurrLine = CurrLine.Replace("³", "&sup3;")
        If CurrLine.Contains("´") Then CurrLine = CurrLine.Replace("´", "&acute;")
        If CurrLine.Contains("µ") Then CurrLine = CurrLine.Replace("µ", "&micro;")
        If CurrLine.Contains("¶") Then CurrLine = CurrLine.Replace("¶", "&para;")
        If CurrLine.Contains("·") Then CurrLine = CurrLine.Replace("·", "&nbsp;") ' --- use middot for nbsp instead of "&middot;"
        If CurrLine.Contains("¸") Then CurrLine = CurrLine.Replace("¸", "&cedil;")
        If CurrLine.Contains("¹") Then CurrLine = CurrLine.Replace("¹", "&sup1;")
        If CurrLine.Contains("º") Then CurrLine = CurrLine.Replace("º", "&ordm;")
        If CurrLine.Contains("»") Then CurrLine = CurrLine.Replace("»", "&raquo;")
        If CurrLine.Contains("¼") Then CurrLine = CurrLine.Replace("¼", "&frac14;")
        If CurrLine.Contains("½") Then CurrLine = CurrLine.Replace("½", "&frac12;")
        If CurrLine.Contains("¾") Then CurrLine = CurrLine.Replace("¾", "&frac34;")
        If CurrLine.Contains("¿") Then CurrLine = CurrLine.Replace("¿", "&iquest;")
        If CurrLine.Contains("À") Then CurrLine = CurrLine.Replace("À", "&Agrave;")
        If CurrLine.Contains("Á") Then CurrLine = CurrLine.Replace("Á", "&Aacute;")
        If CurrLine.Contains("Â") Then CurrLine = CurrLine.Replace("Â", "&Acirc;")
        If CurrLine.Contains("Ã") Then CurrLine = CurrLine.Replace("Ã", "&Atilde;")
        If CurrLine.Contains("Ä") Then CurrLine = CurrLine.Replace("Ä", "&Auml;")
        If CurrLine.Contains("Å") Then CurrLine = CurrLine.Replace("Å", "&Aring;")
        If CurrLine.Contains("Æ") Then CurrLine = CurrLine.Replace("Æ", "&AElig;")
        If CurrLine.Contains("Ç") Then CurrLine = CurrLine.Replace("Ç", "&Ccedil;")
        If CurrLine.Contains("È") Then CurrLine = CurrLine.Replace("È", "&Egrave;")
        If CurrLine.Contains("É") Then CurrLine = CurrLine.Replace("É", "&Eacute;")
        If CurrLine.Contains("Ê") Then CurrLine = CurrLine.Replace("Ê", "&Ecirc;")
        If CurrLine.Contains("Ë") Then CurrLine = CurrLine.Replace("Ë", "&Euml;")
        If CurrLine.Contains("Ì") Then CurrLine = CurrLine.Replace("Ì", "&Igrave;")
        If CurrLine.Contains("Í") Then CurrLine = CurrLine.Replace("Í", "&Iacute;")
        If CurrLine.Contains("Î") Then CurrLine = CurrLine.Replace("Î", "&Icirc;")
        If CurrLine.Contains("Ï") Then CurrLine = CurrLine.Replace("Ï", "&Iuml;")
        If CurrLine.Contains("Ð") Then CurrLine = CurrLine.Replace("Ð", "&ETH;")
        If CurrLine.Contains("Ñ") Then CurrLine = CurrLine.Replace("Ñ", "&Ntilde;")
        If CurrLine.Contains("Ò") Then CurrLine = CurrLine.Replace("Ò", "&Ograve;")
        If CurrLine.Contains("Ó") Then CurrLine = CurrLine.Replace("Ó", "&Oacute;")
        If CurrLine.Contains("Ô") Then CurrLine = CurrLine.Replace("Ô", "&Ocirc;")
        If CurrLine.Contains("Õ") Then CurrLine = CurrLine.Replace("Õ", "&Otilde;")
        If CurrLine.Contains("Ö") Then CurrLine = CurrLine.Replace("Ö", "&Ouml;")
        If CurrLine.Contains("×") Then CurrLine = CurrLine.Replace("×", "&times;")
        If CurrLine.Contains("Ø") Then CurrLine = CurrLine.Replace("Ø", "&Oslash;")
        If CurrLine.Contains("Ù") Then CurrLine = CurrLine.Replace("Ù", "&Ugrave;")
        If CurrLine.Contains("Ú") Then CurrLine = CurrLine.Replace("Ú", "&Uacute;")
        If CurrLine.Contains("Û") Then CurrLine = CurrLine.Replace("Û", "&Ucirc;")
        If CurrLine.Contains("Ü") Then CurrLine = CurrLine.Replace("Ü", "&Uuml;")
        If CurrLine.Contains("Ý") Then CurrLine = CurrLine.Replace("Ý", "&Yacute;")
        If CurrLine.Contains("Þ") Then CurrLine = CurrLine.Replace("Þ", "&THORN;")
        If CurrLine.Contains("ß") Then CurrLine = CurrLine.Replace("ß", "&szlig;")
        If CurrLine.Contains("à") Then CurrLine = CurrLine.Replace("à", "&agrave;")
        If CurrLine.Contains("á") Then CurrLine = CurrLine.Replace("á", "&aacute;")
        If CurrLine.Contains("â") Then CurrLine = CurrLine.Replace("â", "&acirc;")
        If CurrLine.Contains("ã") Then CurrLine = CurrLine.Replace("ã", "&atilde;")
        If CurrLine.Contains("ä") Then CurrLine = CurrLine.Replace("ä", "&auml;")
        If CurrLine.Contains("å") Then CurrLine = CurrLine.Replace("å", "&aring;")
        If CurrLine.Contains("æ") Then CurrLine = CurrLine.Replace("æ", "&aelig;")
        If CurrLine.Contains("ç") Then CurrLine = CurrLine.Replace("ç", "&ccedil;")
        If CurrLine.Contains("è") Then CurrLine = CurrLine.Replace("è", "&egrave;")
        If CurrLine.Contains("é") Then CurrLine = CurrLine.Replace("é", "&eacute;")
        If CurrLine.Contains("ê") Then CurrLine = CurrLine.Replace("ê", "&ecirc;")
        If CurrLine.Contains("ë") Then CurrLine = CurrLine.Replace("ë", "&euml;")
        If CurrLine.Contains("ì") Then CurrLine = CurrLine.Replace("ì", "&igrave;")
        If CurrLine.Contains("í") Then CurrLine = CurrLine.Replace("í", "&iacute;")
        If CurrLine.Contains("î") Then CurrLine = CurrLine.Replace("î", "&icirc;")
        If CurrLine.Contains("ï") Then CurrLine = CurrLine.Replace("ï", "&iuml;")
        If CurrLine.Contains("ð") Then CurrLine = CurrLine.Replace("ð", "&eth;")
        If CurrLine.Contains("ñ") Then CurrLine = CurrLine.Replace("ñ", "&ntilde;")
        If CurrLine.Contains("ò") Then CurrLine = CurrLine.Replace("ò", "&ograve;")
        If CurrLine.Contains("ó") Then CurrLine = CurrLine.Replace("ó", "&oacute;")
        If CurrLine.Contains("ô") Then CurrLine = CurrLine.Replace("ô", "&ocirc;")
        If CurrLine.Contains("õ") Then CurrLine = CurrLine.Replace("õ", "&otilde;")
        If CurrLine.Contains("ö") Then CurrLine = CurrLine.Replace("ö", "&ouml;")
        If CurrLine.Contains("÷") Then CurrLine = CurrLine.Replace("÷", "&divide;")
        If CurrLine.Contains("ø") Then CurrLine = CurrLine.Replace("ø", "&oslash;")
        If CurrLine.Contains("ù") Then CurrLine = CurrLine.Replace("ù", "&ugrave;")
        If CurrLine.Contains("ú") Then CurrLine = CurrLine.Replace("ú", "&uacute;")
        If CurrLine.Contains("û") Then CurrLine = CurrLine.Replace("û", "&ucirc;")
        If CurrLine.Contains("ü") Then CurrLine = CurrLine.Replace("ü", "&uuml;")
        If CurrLine.Contains("ý") Then CurrLine = CurrLine.Replace("ý", "&yacute;")
        If CurrLine.Contains("þ") Then CurrLine = CurrLine.Replace("þ", "&thorn;")
        If CurrLine.Contains("ÿ") Then CurrLine = CurrLine.Replace("ÿ", "&yuml;")
        Do While CurrLine.Contains("  ")
            CurrLine = CurrLine.Replace("  ", "&nbsp;&nbsp;")
        Loop
        Do While CurrLine.Contains("&nbsp; ")
            CurrLine = CurrLine.Replace("&nbsp; ", "&nbsp;&nbsp;")
        Loop
        ' --- Fix italics ---
        Do While CurrLine.Contains("&lt;i&gt;")
            CurrLine = CurrLine.Replace("&lt;i&gt;", "<i>")
        Loop
        Do While CurrLine.Contains("&lt;/i&gt;")
            CurrLine = CurrLine.Replace("&lt;/i&gt;", "</i>")
        Loop
        ' --- Fix bold ---
        Do While CurrLine.Contains("&lt;b&gt;")
            CurrLine = CurrLine.Replace("&lt;b&gt;", "<b>")
        Loop
        Do While CurrLine.Contains("&lt;/b&gt;")
            CurrLine = CurrLine.Replace("&lt;/b&gt;", "</b>")
        Loop
        ' --- Fix underline ---
        Do While CurrLine.Contains("&lt;u&gt;")
            CurrLine = CurrLine.Replace("&lt;u&gt;", "<u>")
        Loop
        Do While CurrLine.Contains("&lt;/u&gt;")
            CurrLine = CurrLine.Replace("&lt;/u&gt;", "</u>")
        Loop
        ' '' --- Fix mbp:pagebreak ---
        ''Do While CurrLine.Contains("&lt;mbp:pagebreak/&gt;")
        ''    CurrLine = CurrLine.Replace("&lt;mbp:pagebreak/&gt;", "<mbp:pagebreak />")
        ''Loop
        ''Do While CurrLine.Contains("&lt;mbp:pagebreak /&gt;")
        ''    CurrLine = CurrLine.Replace("&lt;mbp:pagebreak /&gt;", "<mbp:pagebreak />")
        ''Loop
        ' --- Fix special HTML characters ---
        Do While CurrLine.Contains("&#0;") ' null
            CurrLine = CurrLine.Replace("&#0;", "")
        Loop
        Do While CurrLine.Contains("&#32;") ' space
            CurrLine = CurrLine.Replace("&#32;", " ")
        Loop
        Do While CurrLine.Contains("&#45;") ' hyphen
            CurrLine = CurrLine.Replace("&#45;", "-")
        Loop

        CurrLine = FixTags(CurrLine)
        ' --- Done ---
        Return CurrLine
    End Function

    Private Function FixTags(ByVal CurrLine As String) As String
        If CurrLine.Contains("<overline>") Then
            CurrLine = CurrLine.Replace("<overline>", "<span style=""text-decoration: overline;"">")
        End If
        If CurrLine.Contains("</overline>") Then
            CurrLine = CurrLine.Replace("</overline>", "</span>")
        End If
        If CurrLine.Contains("<u><u>") Then
            CurrLine = CurrLine.Replace("<u><u>", "<u>")
        End If
        If CurrLine.Contains("</u></u>") Then
            CurrLine = CurrLine.Replace("</u></u>", "</u>")
        End If
        ' --- Done ---
        Return CurrLine
    End Function

    Private Function FixInlineImages(ByVal CurrLine As String) As String
        If CurrLine.Contains("&lt;image=") Then
            CurrLine = CurrLine.Replace("&lt;image=", "<img src=""")
            CurrLine = CurrLine.Replace("&gt;", """></img>")
            CurrLine = CurrLine.Replace("&nbsp;", " ")
            CurrLine = CurrLine.Replace("""""", """")
            Dim Filenames() As String = CurrLine.Split("<"c)
            For Each NewImageName As String In Filenames
                If NewImageName.StartsWith("img src=""") Then
                    NewImageName = NewImageName.Replace("img src=""", "")
                    NewImageName = NewImageName.Substring(0, NewImageName.IndexOf(""""c))
                    Try
                        If Not File.Exists(ToPath + "\" + NewImageName) Then
                            File.Copy(FromPath + "\" + NewImageName, ToPath + "\" + NewImageName, True)
                        End If
                    Catch ex As Exception
                        Throw New SystemException("Can't copy file: " + NewImageName + vbCrLf + ex.Message)
                    End Try
                End If
            Next
        End If
        Return CurrLine
    End Function

    Private Sub AboutToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        Dim TempAbout As New AboutMain
        TempAbout.ShowDialog()
        TempAbout = Nothing
    End Sub

    Private Sub ButtonStop_Click(sender As Object, e As EventArgs) Handles ButtonStop.Click
        StopRequested = True
    End Sub

    Private Function FixTableLine(CurrLine As String) As String
        If CurrLine.Contains(vbTab) Then
            CurrLine = CurrLine.Replace(vbTab, "")
        End If
        If CurrLine.Contains("·") Then
            CurrLine = CurrLine.Replace("·", "&nbsp;")
        End If
        Do While CurrLine.Contains(" </td>")
            CurrLine = CurrLine.Replace(" </td>", "</td>")
        Loop
        Return CurrLine
    End Function

End Class
