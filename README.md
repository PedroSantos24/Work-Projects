<!-- PROJECT LOGO -->
<p align="center">
  <img width="" height="" src="https://pictures.alignable.com/eyJidWNrZXQiOiJhbGlnbmFibGV3ZWItcHJvZHVjdGlvbiIsImtleSI6ImV2ZW50cy9waWN0dXJlcy9tZWRpdW0vMzIzNzQwLzE1MjkzMzQyMDBfYmxvYiIsImVkaXRzIjp7fX0=">
</p>



<!-- TABLE OF CONTENTS -->
<details open="open">
  <summary>Table of Contents</summary>
  <ol>
    <li>
      <a href="#about-the-project">About The Project</a>
      <ul>
        <li><a href="#built-with">Built With</a></li>
      </ul>
    </li>
    <li><a href="#usage">Usage</a></li>
  </ol>
</details>


<!-- ABOUT THE PROJECT -->
## About The Project

Troughtout my work in financial services there was a clear necessity to open alot of word docs and printing them to their pdf form. However in my current employer this was done by hand. This little piece of software simplys does that automaticaly. In this specific case the word doc that is gonna be printed is located in a server and their corresponding names will now populated several drop down menus, enabling the user to select witch document in the servers folder he wants to print.

Aditionally the software will open said work doc and will find a piece of text specficif in the code (that will be the same for all the word docs in the server folder) and replace it with what the user inputs into a comboBox. This is done by this simple function:

                findandreplace(wordapp, "xxxx", comboBox6.Text);
                findandreplace(wordapp, "22/10/2019", textBox2.Text);
                
The population of the comboBox is done when the form loads:

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

After both inputs are writen in their corresponding comboBoxes it is required for the user to locate where the pdf shall be exported. I took the liberty to create a button


     private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.RootFolder = Environment.SpecialFolder.Desktop;
            fbd.Description = "Escolher pasta onde gravar as Confirmações de Independência";
            fbd.ShowNewFolderButton = true;
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                textBox3.Text = fbd.SelectedPath; #string of the location
            }
        }

The click of the "export" button the trigger a loop of this function:

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

The performance of the software sometimes isn't perfect. I assume that it only happnes due to the loop not actully being written as a proper loop, but instead as cycle done trough else's.

For a more visual understanding please refer to the section "Usage".


### Built With

* [Spire](https://www.e-iceblue.com/Introduce/word-for-net-introduce.html)
* [C#](https://docs.microsoft.com/en-us/dotnet/csharp/)



<!-- USAGE EXAMPLES -->
## Usage
<p align="center">
  <img width="" height="" src="https://i.imgur.com/KgUdgBG.jpeg">
</p>



[![Linkedin Badge](https://img.shields.io/badge/linkedin-%230077B5.svg?&style=for-the-badge&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/pedroguedes21/)


<!-- MARKDOWN LINKS & IMAGES -->
<!-- https://www.markdownguide.org/basic-syntax/#reference-style-links -->
[contributors-shield]: https://img.shields.io/github/contributors/othneildrew/Best-README-Template.svg?style=for-the-badge
[contributors-url]: https://github.com/othneildrew/Best-README-Template/graphs/contributors
[forks-shield]: https://img.shields.io/github/forks/othneildrew/Best-README-Template.svg?style=for-the-badge
[forks-url]: https://github.com/othneildrew/Best-README-Template/network/members
[stars-shield]: https://img.shields.io/github/stars/othneildrew/Best-README-Template.svg?style=for-the-badge
[stars-url]: https://github.com/othneildrew/Best-README-Template/stargazers
[issues-shield]: https://img.shields.io/github/issues/othneildrew/Best-README-Template.svg?style=for-the-badge
[issues-url]: https://github.com/othneildrew/Best-README-Template/issues
[license-shield]: https://img.shields.io/github/license/othneildrew/Best-README-Template.svg?style=for-the-badge
[license-url]: https://github.com/othneildrew/Best-README-Template/blob/master/LICENSE.txt
[linkedin-shield]: https://img.shields.io/badge/-LinkedIn-black.svg?style=for-the-badge&logo=linkedin&colorB=555
[linkedin-url]: https://linkedin.com/in/othneildrew
[product-screenshot]: images/screenshot.png
