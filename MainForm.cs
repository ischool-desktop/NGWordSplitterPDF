using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Aspose.Words;
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
    }
}
