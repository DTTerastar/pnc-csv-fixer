using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualBasic.FileIO;

namespace Z13.PNC.CSV.Fixer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            textBox1.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        }

        private void Button1Click(object sender, EventArgs e)
        {
            const string fname = "Converted.csv";
            IOrderedEnumerable<Entry> all = Directory.GetFiles(textBox1.Text, "*.csv").Where(a => a != fname).SelectMany(PncCsv.ParseFile).Distinct().OrderBy(a => a.Date);
            var enumerable = all.Where(a => a.Amount != 0).ToArray();
            WaveCsv.WriteFile(Path.Combine(textBox1.Text, fname), enumerable);
            richTextBox1.Text = string.Format("Done.  Wrote {0} records!", enumerable.Length);
        }

        private void Button2Click(object sender, EventArgs e)
        {
            var dialog = new FolderBrowserDialog {SelectedPath = textBox1.Text};
            if (dialog.ShowDialog() == DialogResult.OK)
                textBox1.Text = dialog.SelectedPath;
        }
    }

    internal static class WaveCsv
    {
        public static void WriteFile(string fname, IEnumerable<Entry> items)
        {
            IEnumerable<string> lines = (new[] { "Date,Amount,Description,RefNo" }).Concat(items.Select(a => string.Format("\"{0:yyyy/MM/dd}\",\"{1}\",\"{2}\",\"{3}\"", a.Date, a.Amount, a.Description, a.RefNo)));
            File.WriteAllLines(fname, lines);
        }
    }

    internal static class PncCsv
    {
        public static IEnumerable<Entry> ParseFile(string fname)
        {
            using (var fileReader = new TextFieldParser(fname))
            {
                fileReader.TextFieldType = FieldType.Delimited;
                fileReader.SetDelimiters(",");
                fileReader.HasFieldsEnclosedInQuotes = true;
                while (true)
                {
                    string[] d = fileReader.ReadFields();
                    if (d == null)
                        yield break;

                    DateTime dt;
                    decimal amount;


                    bool validDate = DateTime.TryParse(d[0], out dt);
                    bool validAmount = Decimal.TryParse(d[1], out amount);

                    if (d.Length == 6 && (d[5] == "DEBIT" || d[5] == "CREDIT") && validDate && validAmount)
                    {
                        yield return new Entry
                            {
                                Amount = d[5] == "DEBIT" ? -amount : amount,
                                Date = dt,
                                Description = d[2],
                                RefNo = d[4]
                            };
                    }
                }
            }
        }
    }

    internal class Entry
    {
        public Decimal Amount { get; set; }
        public string Description { get; set; }
        public string RefNo { get; set; }
        public DateTime Date { get; set; }
    }
}