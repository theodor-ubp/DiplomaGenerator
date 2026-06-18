using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;


namespace DiplomaGenerator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }



        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


        private void LoadExcelDataAndGenerateDiplomas(string excelFilePath)
{
    OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

    using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(excelFilePath)))
    {
        var worksheet = package.Workbook.Worksheets[0];
        int totalRows = worksheet.Dimension.End.Row;
        bool hasErrors = false;
        bool hasSuccesses = false;
        StringBuilder errorMessages = new StringBuilder();

        for (int row = 3; row <= totalRows; row++) // Start from row 3 to skip headers
        {
            string courseModule = worksheet.Cells[row, 1].Text;
            string nume = worksheet.Cells[row, 2].Text;
            string cnp = worksheet.Cells[row, 3].Text;
            string dataEliberareText = worksheet.Cells[row, 4].Text;
            string dataExpirareText = worksheet.Cells[row, 5].Text;

            // Log the row data for debugging
            Console.WriteLine($"Processing row {row}: {courseModule}, {nume}, {cnp}, {dataEliberareText}, {dataExpirareText}");

            // Check if the course module exists in the template folder
            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"{courseModule}.docx");
            if (!File.Exists(templatePath))
            {
                errorMessages.AppendLine($"Eroare la rândul {row}: Modulul curs '{courseModule}' nu există în dosarul de șabloane.");
                hasErrors = true;
                continue; // Skip this row
            }

            // Check if the name is empty
            if (string.IsNullOrEmpty(nume))
            {
                errorMessages.AppendLine($"Eroare la rândul {row}: Câmpul Nume este gol.");
                hasErrors = true;
                continue; // Skip this row
            }

            // Check if the CNP is exactly 13 digits
            if (cnp.Length != 13 || !cnp.All(char.IsDigit))
            {
                errorMessages.AppendLine($"Eroare la rândul {row}: CNP-ul trebuie să aibă exact 13 cifre.");
                hasErrors = true;
                continue; // Skip this row
            }

            // Check if the dates are valid
            if (!DateTime.TryParse(dataEliberareText, out DateTime dataEliberare))
            {
                errorMessages.AppendLine($"Eroare la rândul {row}: 'Dată eliberare' nu este o dată validă.");
                hasErrors = true;
                continue; // Skip this row
            }

            if (!DateTime.TryParse(dataExpirareText, out DateTime dataExpirare))
            {
                errorMessages.AppendLine($"Eroare la rândul {row}: 'Dată expirare' nu este o dată validă.");
                hasErrors = true;
                continue; // Skip this row
            }

            try
            {
                // If no errors, process the row
                FillTemplate(templatePath, nume, cnp, dataEliberare, dataExpirare, false); // false for bulk mode
                hasSuccesses = true;
            }
            catch (Exception ex)
            {
                errorMessages.AppendLine($"Eroare la rândul {row} pentru {nume}: {ex.Message}");
                hasErrors = true;
            }
        }

        if (hasErrors && hasSuccesses)
        {
            MessageBox.Show($"Au fost generate diplomele pentru rândurile valide, însă au existat erori la anumite rânduri:\n\n{errorMessages.ToString()}", "Atenție", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        else if (hasErrors)
        {
            MessageBox.Show($"Nu au fost generate diplome din cauza erorilor următoare:\n\n{errorMessages.ToString()}", "Eroare", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        else
        {
            MessageBox.Show("Diplome generate cu succes pentru toate rândurile valide!", "Succes", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}


        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Disable controls
            Invoke(new Action(() =>
            {
                SetControlsEnabled(false, !string.IsNullOrEmpty(txtExcelFilePath.Text));
            }));

            // Show progress bar and start processing
            progressBarGenerating.Visible = true;
            progressBarGenerating.Style = ProgressBarStyle.Marquee;
            progressBarGenerating.MarqueeAnimationSpeed = 30;

            Task.Run(() =>
            {
                try
                {
                    // Check if an Excel file has been selected
                    if (!string.IsNullOrEmpty(txtExcelFilePath.Text))
                    {
                        // Process the Excel file
                        string excelFilePath = txtExcelFilePath.Text;
                        LoadExcelDataAndGenerateDiplomas(excelFilePath);
                    }
                    else
                    {
                        // Manual entry processing
                        if (!string.IsNullOrEmpty(txtNume.Text) && !string.IsNullOrEmpty(txtCNP.Text))
                        {
                            string cnp = txtCNP.Text;

                            // Validate the CNP length
                            if (cnp.Length != 13)
                            {
                                Invoke(new Action(() =>
                                {
                                    MessageBox.Show("CNP incorect. Trebuie să conțină exact 13 cifre.", "Eroare", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }));
                                return; // Exit the method if the CNP is invalid
                            }

                            string courseModule = string.Empty;

                            // Access the comboBoxCourseModule safely using Invoke
                            Invoke(new Action(() =>
                            {
                                courseModule = comboBoxCourseModule.SelectedItem.ToString();
                            }));

                            // Check if the user has selected the default "Alege model..." option
                            if (courseModule == "Alege model...")
                            {
                                Invoke(new Action(() =>
                                {
                                    MessageBox.Show("Vă rugăm să selectați un model din listă.", "Eroare", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }));
                                return;
                            }

                            string nume = txtNume.Text;
                            DateTime dataEliberare = dateTimePickerEliberare.Value;
                            DateTime dataExpirare = dateTimePickerExpirare.Value;

                            // Construct the template path based on the selected item
                            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"{courseModule}.docx");

                            // Check if the template file exists
                            if (!File.Exists(templatePath))
                            {
                                Invoke(new Action(() =>
                                {
                                    MessageBox.Show("Fișierul model selectat nu a fost găsit.", "Eroare", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }));
                                return;
                            }

                            // Fill the template and generate the files
                            FillTemplate(templatePath, nume, cnp, dataEliberare, dataExpirare, true); // true for manual mode
                        }
                        else
                        {
                            Invoke(new Action(() =>
                            {
                                MessageBox.Show("Nu ati introdus toate campurile.");
                            }));
                        }
                    }
                }
                catch (Exception ex)
                {
                    Invoke(new Action(() =>
                    {
                        MessageBox.Show($"A apărut o eroare în timpul procesării: {ex.Message}", "Eroare", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                }
                finally
                {
                    // Hide progress bar and re-enable controls once processing is complete
                    Invoke(new Action(() =>
                    {
                        progressBarGenerating.Visible = false;
                        SetControlsEnabled(true, !string.IsNullOrEmpty(txtExcelFilePath.Text));
                    }));
                }
            });
        }


        private void ConvertToPdf(string wordFilePath, string pdfFilePath)
        {
            // Confirm the file exists just before attempting to open it
            if (!File.Exists(wordFilePath))
            {
                throw new FileNotFoundException($"Fișierul Word nu a fost găsit: {wordFilePath}");
            }

            try
            {
                var wordApp = new Microsoft.Office.Interop.Word.Application();

                // Open the Word document
                var document = wordApp.Documents.Open(
                    FileName: wordFilePath,
                    ConfirmConversions: false,
                    ReadOnly: false,
                    AddToRecentFiles: false,
                    Format: Microsoft.Office.Interop.Word.WdOpenFormat.wdOpenFormatAuto
                );

                // Convert to PDF
                document.ExportAsFixedFormat(pdfFilePath, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);

                // Close the document and Word application
                document.Close(false);
                wordApp.Quit(false);
            }
            catch (Exception ex)
            {
                throw new Exception($"A apărut o eroare la conversia PDF: {ex.Message}", ex);
            }
        }


        private void dateTimePickerEliberare_ValueChanged(object sender, EventArgs e)
        {
            // Get the selected issue date
            DateTime issueDate = dateTimePickerEliberare.Value;

            // Calculate the expiration date: next year minus 1 day
            DateTime expirationDate = issueDate.AddYears(1).AddDays(-1);

            // Set the expiration date picker to this calculated date
            dateTimePickerExpirare.Value = expirationDate;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            // Initialize the date pickers
            dateTimePickerEliberare.Value = DateTime.Today; // Set Data Eliberare to today's date
            dateTimePickerExpirare.Value = DateTime.Today.AddYears(1).AddDays(-1); // Set Data Expirare to one year minus one day from today

            // Get the directory of the executable
            string exeDirectory = AppDomain.CurrentDomain.BaseDirectory;

            // Get all .docx files in the root directory
            string[] docxFiles = Directory.GetFiles(exeDirectory, "*.docx");

            // Extract and sort the file names numerically based on the numeric part after "HS"
            var sortedFiles = docxFiles.Select(f => Path.GetFileNameWithoutExtension(f))
                                       .OrderBy(f => int.Parse(System.Text.RegularExpressions.Regex.Match(f, @"\d+").Value))
                                       .ToList();

            // Add the default item at the beginning
            comboBoxCourseModule.Items.Add("Alege model...");

            // Populate the ComboBox with sorted file names
            foreach (string fileName in sortedFiles)
            {
                comboBoxCourseModule.Items.Add(fileName);
            }

            // Set the default selected item to "Alege model..."
            comboBoxCourseModule.SelectedIndex = 0;
        }

        private void FillTemplate(string templatePath, string nume, string cnp, DateTime dataEliberare, DateTime dataExpirare, bool isManual)
        {
            Word.Application wordApp = null;
            try
            {
                wordApp = new Word.Application();
                var document = wordApp.Documents.Open(templatePath);

                // Replace placeholders with actual data
                FindAndReplace(wordApp, "{Nume}", nume);
                FindAndReplace(wordApp, "{CNP}", cnp);
                FindAndReplace(wordApp, "{DataEliberare}", dataEliberare.ToShortDateString());
                FindAndReplace(wordApp, "{DataExpirare}", dataExpirare.ToShortDateString());

                // Create folder on the Desktop based on the template name
                string folderName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), Path.GetFileNameWithoutExtension(templatePath));
                Directory.CreateDirectory(folderName);

                // Generate the filename using the course module, name, and issue date
                string fileName = $"{Path.GetFileNameWithoutExtension(templatePath)}_{nume}_{dataEliberare:dd.MM.yyyy}";

                // Save the document with the new name inside the folder
                string outputDocPath = Path.Combine(folderName, $"{fileName}.docx");

                document.SaveAs2(outputDocPath);
                document.Close();

                // Apply delay only in manual mode
                if (isManual)
                {
                    for (int i = 0; i < 5; i++)  // Retry up to 5 times
                    {
                        if (File.Exists(outputDocPath))
                        {
                            break;  // Exit the loop if the file exists
                        }
                        System.Threading.Thread.Sleep(100);  // Wait for 100 ms before retrying
                    }
                }

                // Ensure the document is saved and exists before converting to PDF
                if (File.Exists(outputDocPath))
                {
                    // Convert to PDF
                    string outputPdfPath = Path.Combine(folderName, $"{fileName}.pdf");
                    ConvertToPdf(outputDocPath, outputPdfPath);

                    // If in manual mode, show the success message
                    if (isManual)
                    {
                        Invoke(new Action(() =>
                        {
                            MessageBox.Show("Diploma generata cu succes!", "Succes", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }));
                    }
                }
                else
                {
                    throw new FileNotFoundException($"Fișierul Word nu a fost găsit după salvare: {outputDocPath}");
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Eroare la procesarea sablonului pentru {nume}: {ex.Message}", ex);
            }
            finally
            {
                wordApp?.Quit();
            }
        }



        private void FindAndReplace(Word.Application wordApp, object findText, object replaceWithText)
        {
            wordApp.Selection.Find.ClearFormatting();
            wordApp.Selection.Find.Execute(FindText: ref findText, ReplaceWith: ref replaceWithText, Replace: Word.WdReplace.wdReplaceAll);
        }
        private void btnGenerateDiplomas_Click(object sender, EventArgs e)
        {
            string cnp = txtCNP.Text;

            // Validate the CNP length
            if (cnp.Length != 13)
            {
                MessageBox.Show("CNP incorect. Trebuie să conțină exact 13 cifre.", "Eroare", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return; // Exit the method if the CNP is invalid
            }

            // Get the selected template from the ComboBox
            string selectedTemplate = comboBoxCourseModule.SelectedItem.ToString();

            // Check if the user has selected the default "Alege model..." option
            if (selectedTemplate == "Alege model...")
            {
                MessageBox.Show("Vă rugăm să selectați un model din listă.", "Eroare", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Construct the template path based on the selected item
            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"{selectedTemplate}.docx");

            // Check if the template file exists
            if (!File.Exists(templatePath))
            {
                MessageBox.Show("Fișierul model selectat nu a fost găsit.", "Eroare", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string nume = txtNume.Text;
            DateTime dataEliberare = dateTimePickerEliberare.Value;
            DateTime dataExpirare = dateTimePickerExpirare.Value;

            // Assuming this method is called during manual generation, pass true for isManual
            FillTemplate(templatePath, nume, cnp, dataEliberare, dataExpirare, true);

            MessageBox.Show("Diploma generata cu succes!", "Succes", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnChooseExcel_Click(object sender, EventArgs e)
        {
            // Check if the button text is "Schimba"
            if (btnChooseExcel.Text == "Schimba")
            {
                // Reset the form without showing the dialog
                ResetForm();
                return;
            }

            // Show the file dialog only if the button text is "Alege"
            openFileDialogExcel.Filter = "Excel Files|*.xlsx;*.xls";
            if (openFileDialogExcel.ShowDialog() == DialogResult.OK)
            {
                string excelFilePath = openFileDialogExcel.FileName;

                // Display the file path in the TextBox
                txtExcelFilePath.Text = excelFilePath;

                // Change button text to "Schimba"
                btnChooseExcel.Text = "Schimba";

                // Disable manual fields since Excel file is chosen
                txtNume.Enabled = false;
                txtCNP.Enabled = false;
                comboBoxCourseModule.Enabled = false;
                dateTimePickerEliberare.Enabled = false;
                dateTimePickerExpirare.Enabled = false;
            }
        }

        private void ResetForm()
        {
            // Clear the selected Excel file path
            txtExcelFilePath.Text = string.Empty;

            // Enable manual entry fields
            txtNume.Enabled = true;
            txtCNP.Enabled = true;
            comboBoxCourseModule.Enabled = true;
            dateTimePickerEliberare.Enabled = true;
            dateTimePickerExpirare.Enabled = true;

            // Reset the button text
            btnChooseExcel.Text = "Alege";

            // Optionally reset the content of the manual fields
            txtNume.Text = string.Empty;
            txtCNP.Text = string.Empty;
            comboBoxCourseModule.SelectedIndex = 0; // Assuming the first index is "Alege model..."
            dateTimePickerEliberare.Value = DateTime.Now;
            dateTimePickerExpirare.Value = DateTime.Now.AddYears(1).AddDays(-1);
        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string message = "Acest program utilizează șabloane .docx din directorul bin al aplicației.\n" +
                             "Pentru fiecare rând valid din fișierul Excel sau introducerea manuală, sunt create fișierele DOCX și PDF.\n" +
                             "Aceste fișiere sunt salvate pe desktop în dosarul cu numele modulului cursului.\n" +
                             "\nPuteți să vă creați propriile șabloane .docx, asigurându-vă că sunt poziționate în directorul bin și conțin următoarele câmpuri obligatorii:\n" +
                             "- {Nume} (numele cursantului)\n" +
                             "- {CNP} (CNP-ul cursantului)\n" +
                             "- {DataEliberare} (data emiterii diplomei)\n" +
                             "- {DataExpirare} (data expirării diplomei)\n" +
                             "\nModul Manual:\n" +
                             "- Dacă nu este selectat un fișier Excel, utilizatorul poate introduce manual datele pentru a genera o singură diplomă.\n" +
                             "\nAvertismente:\n" +
                             "- Fișierele existente cu același nume în același dosar vor fi suprascrise.\n" +
                             "- Programul acceptă doar formatul de dată (dd.MM.yyyy). Asigurați-vă că datele sunt introduse corect în acest format.\n";

            MessageBox.Show(message, "Informații Program", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string message = "Cum creez sabloane?\n\n" +
                             "Pentru a crea sabloane pentru diplome, urmati acesti pasi:\n\n" +
                             "1. Deschideti Microsoft Word sau orice editor de text care poate salva fisiere in format .docx.\n" +
                             "2. Creati un design atractiv pentru diploma, folosind text, imagini, si formatare dupa preferinte.\n" +
                             "3. Introduceti campurile speciale care vor fi completate automat de aplicatie:\n" +
                             "   - {Nume} pentru numele beneficiarului.\n" +
                             "   - {CNP} pentru codul numeric personal/serie diploma (exact 13 caractere).\n" +
                             "   - {DataEliberare} pentru data eliberarii.\n" +
                             "   - {DataExpirare} pentru data expirarii.\n\n" +
                             "Fiti creativi si creati cele mai frumoase diplome, dar nu uitati:\n" +
                             "Salvati fisierul .docx in folderul 'bin' al aplicatiei.\n" +
                             "Acest lucru asigura ca sabloanele vor fi disponibile pentru generarea diplomelor.\n\n" +
                             "Mult succes!";

            MessageBox.Show(message, "Cum creez sabloane?", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string message = "Cum creez un xlsx pentru Auto?\n\n" +
                 "Urmati acesti pasi:\n\n" +
                 "1. Deschideti Microsoft Excel sau orice alt editor de foi de calcul care poate salva fisiere in format .xlsx.\n" +
                 "2. In prima foaie (Sheet 1), creati un tabel cu urmatoarele coloane, incepand de la randul 3 (randul 1,2 rezervate pentru detalii, cap de tabel, etc.):\n" +
                 "   - Coloana A: Numele fisierului .docx (fara extensie), care reprezinta sablonul utilizat pentru diploma. Asigurati-va ca acest fisier se afla in folderul 'bin' al aplicatiei.\n" +
                 "   - Coloana B: Nume (numele complet al beneficiarului).\n" +
                 "   - Coloana C: CNP (codul numeric personal al beneficiarului sau seria diplomei, exact 13 caractere).\n" +
                 "   - Coloana D: Data Eliberare (data in formatul dd.MM.yyyy).\n" +
                 "   - Coloana E: Data Expirare (data in formatul dd.MM.yyyy).\n\n" +
                 "3. Completati fiecare rand cu informatiile corespunzatoare pentru fiecare diploma.\n" +
                 "4. Salvati fisierul .xlsx in locatia dorita.\n\n" +
                 "Nu uitati:\n" +
                 "In coloana A trebuie sa introduceti numele fisierului .docx (sablonul) pe care doriti sa il utilizati pentru fiecare diploma. Asigurati-va ca fisierele .docx sunt plasate in folderul 'bin' al aplicatiei.\n" +
                 "Fisierele Excel trebuie sa fie corect completate si sa contina toate campurile obligatorii pentru ca diplomele sa fie generate corect.\n\n" +
                 "Creati si gestionati fisierele Excel cu atentie pentru a obtine cele mai bune rezultate!\n" +
                 "Pentru model, gasiti in folderul instalarii fisierul 'XLS Generator Diplome v1'";

            MessageBox.Show(message, "Cum creez un xlsx pentru Auto?", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void SetControlsEnabled(bool enabled, bool isExcelFileChosen)
        {
            // Enable/Disable general controls
            txtExcelFilePath.Enabled = enabled;
            btnChooseExcel.Enabled = enabled;
            button1.Enabled = enabled; // Generate button
            button2.Enabled = enabled; // Info button
            button3.Enabled = enabled; // Template help button
            button4.Enabled = enabled; // Excel help button

            // Enable/Disable manual entry fields based on whether an Excel file is chosen
            txtNume.Enabled = enabled && !isExcelFileChosen;
            txtCNP.Enabled = enabled && !isExcelFileChosen;
            comboBoxCourseModule.Enabled = enabled && !isExcelFileChosen;
            dateTimePickerEliberare.Enabled = enabled && !isExcelFileChosen;
            dateTimePickerExpirare.Enabled = enabled && !isExcelFileChosen;
        }
    }

}


