using System;
using word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using Spire.Doc;
using System.Text.RegularExpressions;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void findandreplace(word.Application wordapp, object ToFindText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object nmatchAllforms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object replace = 2;
            object wrap = 1;

            wordapp.Selection.Find.Execute(ref ToFindText,
                ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundLike,
                ref nmatchAllforms, ref forward,
                ref wrap, ref format, ref replaceWithText,
                ref replace, ref matchKashida,
                ref matchDiactitics, ref matchAlefHamza,
                ref matchControl);
        }

        private void createworddocument(object filename, object saveas)
        {
            word.Application wordapp = new word.Application();
            object missing = Missing.Value;
            word.Document myworddoc = null;

            if (File.Exists((string)filename))
            {

                object readOnly = false;
                object isVisible = false;
                wordapp.Visible = false;

                myworddoc = wordapp.Documents.Open(ref filename, ref missing, ref readOnly,
                                        ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing, ref missing);
                myworddoc.Activate();

                findandreplace(wordapp, "xxxx", comboBox6.Text);
                findandreplace(wordapp, "22/10/2019", textBox2.Text);


            }
            else
            {
                MessageBox.Show("Ficheiro em falta!");
            }

            myworddoc.SaveAs2(ref saveas);

            myworddoc.Close();
            wordapp.Quit();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            createworddocument($"\\\\SERVIDOR\\{comboBox1.Text}", $"{textBox3.Text}\\{comboBox1.Text}");
            string str = $"{comboBox1.Text}";
            int pFrom = str.IndexOf(" - ") + " - ".Length;
            int pTo = str.LastIndexOf(".docx");
            string result = str.Substring(pFrom, pTo - pFrom);          textBox1.Text = result;
            Document document = new Document();
            document.LoadFromFile($"{ textBox3.Text}\\{ comboBox1.Text}");
            document.SaveToFile($"{ textBox3.Text}\\Random name {textBox1.Text}.pdf", FileFormat.PDF);
            File.Delete($"{textBox3.Text}\\{comboBox1.Text}");
            if (string.IsNullOrWhiteSpace(comboBox2.Text))
            {
                MessageBox.Show("Sucesso!");
                return;
            }
            else

            createworddocument($"\\\\SERVIDOR\\{comboBox2.Text}", $"{textBox3.Text}\\{comboBox2.Text}");
            string str2 = $"{comboBox2.Text}";
            int pFrom2 = str2.IndexOf(" - ") + " - ".Length;
            int pTo2 = str2.LastIndexOf(".docx");
            string result2 = str2.Substring(pFrom2, pTo2 - pFrom2);
            textBox4.Text = result2;
            Document document1 = new Document();
            document1.LoadFromFile($"{ textBox3.Text}\\{ comboBox2.Text}");
            document1.SaveToFile($"{ textBox3.Text}\\Random name {textBox4.Text}.pdf", FileFormat.PDF);
            File.Delete($"{textBox3.Text}\\{comboBox2.Text}");
            if (string.IsNullOrWhiteSpace(comboBox3.Text))
            {
                MessageBox.Show("Sucesso!");
                return;
            }
            else

            createworddocument($"\\\\SERVIDOR\\{comboBox3.Text}", $"{textBox3.Text}\\{comboBox3.Text}");
            string str3 = $"{comboBox3.Text}";
            int pFrom3 = str3.IndexOf(" - ") + " - ".Length;
            int pTo3 = str3.LastIndexOf(".docx");
            string result3 = str3.Substring(pFrom3, pTo3 - pFrom3);
            textBox5.Text = result3;
            Document document2 = new Document();
            document2.LoadFromFile($"{ textBox3.Text}\\{ comboBox3.Text}");
            document2.SaveToFile($"{ textBox3.Text}\\Random name {textBox5.Text}.pdf", FileFormat.PDF);
            File.Delete($"{textBox3.Text}\\{comboBox3.Text}");
            if (string.IsNullOrWhiteSpace(comboBox4.Text))
            {
                MessageBox.Show("Sucesso!");
                return;
            }
            else

            createworddocument($"\\\\SERVIDOR\\{comboBox4.Text}", $"{textBox3.Text}\\{comboBox4.Text}");
            string str4 = $"{comboBox4.Text}";
            int pFrom4 = str4.IndexOf(" - ") + " - ".Length;
            int pTo4 = str4.LastIndexOf(".docx");
            string result4 = str4.Substring(pFrom4, pTo4 - pFrom4);
            textBox6.Text = result4;
            Document document3 = new Document();
            document3.LoadFromFile($"{ textBox3.Text}\\{ comboBox4.Text}");
            document3.SaveToFile($"{ textBox3.Text}\\Random name {textBox6.Text}.pdf", FileFormat.PDF);
            File.Delete($"{textBox3.Text}\\{comboBox4.Text}");
            if (string.IsNullOrWhiteSpace(comboBox5.Text))
            {
                MessageBox.Show("Sucesso!");
                return;
            }
            else

            createworddocument($"\\\\SERVIDOR\\{comboBox5.Text}", $"{textBox3.Text}\\{comboBox5.Text}");
            string str5 = $"{comboBox5.Text}";
            int pFrom5 = str5.IndexOf(" - ") + " - ".Length;
            int pTo5 = str5.LastIndexOf(".docx");
            string result5 = str5.Substring(pFrom5, pTo5 - pFrom5);
            textBox7.Text = result5;
            Document document4 = new Document();
            document4.LoadFromFile($"{ textBox3.Text}\\{ comboBox5.Text}");
            document4.SaveToFile($"{ textBox3.Text}\\Random name {textBox7.Text}.pdf", FileFormat.PDF);
            File.Delete($"{textBox3.Text}\\{comboBox5.Text}");
            if (string.IsNullOrWhiteSpace(comboBox7.Text))
            {
                MessageBox.Show("Sucesso!");
                return;
            }
            else

            createworddocument($"\\\\SERVIDOR\\{comboBox5.Text}", $"{textBox3.Text}\\{comboBox7.Text}");
            string str6 = $"{comboBox7.Text}";
            int pFrom6 = str6.IndexOf(" - ") + " - ".Length;
            int pTo6 = str6.LastIndexOf(".docx");
            string result6 = str6.Substring(pFrom6, pTo6 - pFrom6);
            textBox8.Text = result6;
            Document document5 = new Document();
            document5.LoadFromFile($"{ textBox3.Text}\\{ comboBox7.Text}");
            document5.SaveToFile($"{ textBox3.Text}\\Random name {textBox8.Text}.pdf", FileFormat.PDF);
            File.Delete($"{textBox3.Text}\\{comboBox7.Text}");

            MessageBox.Show("Sucesso!");
        }
        private void Form1_Load(object filename, EventArgs e)
        {
            textBox3.Visible = false;
            string[] files = Directory.GetFiles(@"\\SERVIDOR", "*docx");
            foreach (string file in files)
            {
                string path = Path.GetFileName(file);
                comboBox1.Items.Add(path);
                comboBox2.Items.Add(path);
                comboBox3.Items.Add(path);
                comboBox4.Items.Add(path);
                comboBox5.Items.Add(path);
                comboBox7.Items.Add(path);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.RootFolder = Environment.SpecialFolder.Desktop;
            fbd.Description = "Escolher pasta onde gravar as Confirmações de Independência";
            fbd.ShowNewFolderButton = true;

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                textBox3.Text = fbd.SelectedPath;
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
       

           
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
 
