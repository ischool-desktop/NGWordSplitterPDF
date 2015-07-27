using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
using StuData = NGWordSplitter.IDNumberLookup.StudentData;

namespace NGWordSplitter
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void btnSplite_Click(object sender, EventArgs e)
        {
            int section_position = 0;

            IDNumberLookup idlookup = new IDNumberLookup("vlookup.xlsx");
            Document doc = new Document(txtDocFN.Text);
            Document doc2 = null;

            Regex rx = new Regex(txtPattern.Text);

            //臺北市南港區南港國民小學
            string fn = string.Empty;
            string fn2 = string.Empty;
            StuData sdata = null;
            string error_file = string.Empty;
            for (int i = 0; i <= doc.Sections[section_position].Body.Tables.Count; i++)
            {
                Table t = doc.Sections[section_position].Body.Tables[i] as Table;
                Paragraph p = doc.Sections[section_position].Body.Paragraphs[i] as Paragraph;

                if (t == null) continue;

                if (t.GetText().IndexOf("臺北市南港區南港國民小學") >= 0)
                {
                    doc2 = NewDocument();
                    fn = string.Empty;
                    sdata = null;
                    error_file = string.Empty;
                }

                if (t != null)
                {
                    t = doc2.ImportNode(t, true) as Table;
                    doc2.Sections[section_position].Body.Tables.Add(t);
                }

                if (p != null)
                {
                    p = doc2.ImportNode(p, true) as Paragraph;
                    doc2.Sections[section_position].Body.Paragraphs.Add(p);
                }

                Match m = rx.Match(t.GetText());

                if (m.Success)
                {
                    int sno;

                    error_file = string.Format("{0}_{1}_{2}", m.Groups[1].Value, m.Groups[2].Value, m.Groups[3].Value);

                    if (int.TryParse(m.Groups[2].Value, out sno))
                    {
                        sdata = idlookup.GetStudentData(m.Groups[1].Value, sno.ToString());

                        if (sdata != null)
                        {
                            fn = Util.GenerateFileName(sdata.IDNumber, "06");
                            fn2 = string.Format("{0}_{1}_{2}_{3}", sdata.IDNumber, m.Groups[1], m.Groups[2], m.Groups[3]);

                            if (sdata.Name.Trim() != m.Groups[3].Value.Trim())
                                fn = "X_" + fn;
                        }
                    }
                }

                if (t.GetText().IndexOf("家長簽章") >= 0)
                {
                    if (fn == string.Empty)
                        fn = error_file;

                    if (doc2 != null)
                    {
                        doc2.Save("output\\" + fn + ".pdf", SaveFormat.Pdf);
                        //doc2.Save("output\\" + fn2 + ".docx", SaveFormat.Docx);
                    }
                }
            }

            MessageBox.Show("Complete!");
        }

        private static Document NewDocument()
        {
            Document doc2 = new Document();

            Section sec = doc2.Sections[0];
            sec.Body.Paragraphs.Clear();

            sec.PageSetup.PaperSize = PaperSize.A4;
            return doc2;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.Save();
        }

        private void btnHTMLSplite_Click(object sender, EventArgs e)
        {
            SpliteByHtml();
        }

        private Tuple<List<string>, int> GetBeginHtml(List<string> lines)
        {
            Tuple<List<string>, int> result = FindToPattern(lines, "<BODY>", 0);

            StringBuilder header = new StringBuilder();

            foreach (string line in result.Item1)
                header.AppendLine(line);
            return result;
        }

        private List<string> ReadAllLine()
        {
            StreamReader sr = new StreamReader(txtDocFN.Text, System.Text.Encoding.UTF8);

            List<string> lines = new List<string>();
            while (sr.Peek() > 0)
            {
                lines.Add(sr.ReadLine());
            }
            sr.Close();
            return lines;
        }

        private void SpliteByHtml()
        {
            IDNumberLookup idlookup = new IDNumberLookup("vlookup.xlsx");
            List<string> sourceLines = ReadAllLine(); //讀取全部的文字行。

            Regex rx = new Regex(txtPattern.Text);

            //臺北市南港區南港國民小學
            string fileNumber = string.Empty;
            string pdfInfo = string.Empty;
            StuData sdata = null;
            string error_file = string.Empty;

            Tuple<List<string>, int> result = GetBeginHtml(sourceLines);
            string template = GenerateTempalte(result); // <%Content%>

            //先找到第一個
            result = FindToPattern(sourceLines, "<div width=\"649\" align=\"center\">", result.Item2);

            while (true)
            {
                //找學生的內容。
                result = FindToPattern(sourceLines, "<div width=\"649\" align=\"center\">", result.Item2 + 1);
                result.Item1.RemoveAt(result.Item1.Count - 1); //移掉最後一行，因為那是下一個學生的。

                StringBuilder content = new StringBuilder();
                content.AppendLine("<div width=\"649\" align=\"center\">");
                foreach (string line in result.Item1)
                    content.AppendLine(line);

                string fullContent = template.Replace("<%Content%>", content.ToString());

                Match m = rx.Match(fullContent);

                if (m.Success)
                {
                    int sno;

                    error_file = string.Format("{0}_{1}_{2}", m.Groups[1].Value, m.Groups[2].Value, m.Groups[3].Value);

                    if (int.TryParse(m.Groups[2].Value, out sno))
                    {
                        sdata = idlookup.GetStudentData(m.Groups[1].Value, sno.ToString());

                        if (sdata != null)
                        {
                            fileNumber = Util.GenerateFileName(sdata.IDNumber, "06");
                            pdfInfo = string.Format("{0}_{1}_{2}_{3}", sdata.IDNumber, m.Groups[1], m.Groups[2], m.Groups[3]);

                            if (sdata.Name.Trim() != m.Groups[3].Value.Trim())
                                fileNumber = "X_" + fileNumber;
                        }
                    }

                    string path = Path.Combine(Application.StartupPath, "output");
                    string filedoc = Path.Combine(path, fileNumber + ".doc");
                    string filedocx = Path.Combine(path, fileNumber + ".docx");
                    string filepdfFinal = Path.Combine(path, fileNumber + ".pdf");
                    string filepdf = Path.Combine(path, pdfInfo + ".pdf");

                    //儲存為 doc
                    StreamWriter sw = new StreamWriter(filedoc, false, Encoding.UTF8);
                    sw.Write(fullContent);
                    sw.Close();

                    Document tempdoc = null;
                    if (chkOutputDocx.Checked)
                    {
                        //轉為 docx
                        tempdoc = new Document(filedoc);
                        tempdoc.Save(filedocx, SaveFormat.Docx);
                    }

                    PdfSaveOptions pso = new PdfSaveOptions();
                    pso.ExportDocumentStructure = true;
                    pso.UseHighQualityRendering = true;

                    if (chkOutputPDF.Checked)
                    {
                        //轉為 pdf
                        tempdoc = new Document(filedocx);
                        tempdoc.Save(filepdfFinal, pso);
                    }
                }
                else
                    break;
            }

            MessageBox.Show("Complete!");
        }

        private static string GenerateTempalte(Tuple<List<string>, int> result)
        {
            StringBuilder template = new StringBuilder();
            foreach (string line in result.Item1)
                template.AppendLine(line);
            template.AppendLine("<%Content%></BODY></HTML>");

            return template.ToString();
        }

        private Tuple<List<string>, int> FindToPattern(List<string> input, string pattern, int startLine)
        {
            List<string> output = new List<string>();
            int endLine = 0;

            for (int line = startLine; line < input.Count; line++)
            {
                string lineText = input[line].Trim();
                output.Add(lineText);

                if (lineText == pattern.Trim())
                {
                    endLine = line;
                    break;
                }
            }

            return new Tuple<List<string>, int>(output, endLine);
        }

        private void chkOutputPDF_CheckedChanged(object sender, EventArgs e)
        {
            if (chkOutputPDF.Checked)
                chkOutputDocx.Checked = true;
        }

        private void chkOutputDocx_CheckedChanged(object sender, EventArgs e)
        {
            if (!chkOutputDocx.Checked)
                chkOutputPDF.Checked = false;
        }
    }
}
