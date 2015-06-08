using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace ExcelProject
{
    public partial class ManualFormForGiBProjects : System.Windows.Forms.Form
    {
        private string connectionStringDefault = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Excel_Automation_Training\Projektliste GiB.xlsx;Extended Properties=""Excel 12.0;HDR=YES;""";
        private DataTable form_dt;
        private string[] headers;
        private string[] bearbeiterValues;
        private Dictionary<string,string> previousTextBoxValues = new Dictionary<string,string>();
        private Dictionary<string, string> currentTextBoxValues = new Dictionary<string,string>();
        private string dialogBoxMessage = string.Empty;
        TextBox lockedOne = null;
        private static int dynamicX = 5;
        private static int labelY = 40;
        private static int textboxY = labelY + 15;
        

        public ManualFormForGiBProjects()
        {
            InitializeComponent();
            
            
        }

        private void SetupInitialValues()
        {   
            //form_dt = ExcelLogic.FillDataTable();
            //headers = ExcelLogic.GetExcelColumns();
            //bearbeiterValues = ExcelLogic.GetBearbeiterValues();
            form_dt = SQLiteLogic.FillDataTable();
            headers = SQLiteLogic.GetColumnNames();
            bearbeiterValues = SQLiteLogic.GetBearbeiterValues();
            
        }

        private void ManualFormForGiBProjects_Load(object sender, EventArgs e)
        {

            SetupInitialValues();
            AddUIControlsToPanel2ofSplitCOntainer();
            ExcelGridView.DataSource = form_dt;
            bearbeiter_combobox();
            
        }


        private void bearbeiter_combobox()
        {
            try
            {
                Label bearbeiterLabel = new Label();
            ComboBox bearbeiterComboBox = new ComboBox();
            Button bearbeiterButton = new Button();

            bearbeiterLabel.Name = "bearbeiterLabel";
            bearbeiterComboBox.Name = "bearbeiterComboBox";
            bearbeiterButton.Name = "bearbeiterButton";
            bearbeiterButton.Click += bearbeiterButton_Click;
            bearbeiterLabel.Text = "Projekte des Bearbeiters: ";
            bearbeiterLabel.AutoSize = true;

            foreach (string bearbeiter in bearbeiterValues)
            {
                if (bearbeiter != null)
                {
                    bearbeiterComboBox.Items.Add(bearbeiter);
                }            
            }
            bearbeiterButton.Text = "Anzeigen";

            bearbeiterLabel.Location = new Point(5, 10);
            bearbeiterComboBox.Location = new Point(140, 10);
            bearbeiterButton.Location = new Point(270, 10);

            bearbeiterLabel.Show();
            bearbeiterComboBox.Show();
            bearbeiterButton.Show();

            this.splitContainer1.Panel2.Controls.Add(bearbeiterLabel);
            this.splitContainer1.Panel2.Controls.Add(bearbeiterComboBox);
            this.splitContainer1.Panel2.Controls.Add(bearbeiterButton);
            }
            catch (Exception ex)
            {
                 System.Windows.Forms.MessageBox.Show(ex.Message);
                Application.Exit();    
            }
        }

        void bearbeiterButton_Click(object sender, EventArgs e)
        {   
            // geting selected value in combobox
            ComboBox comb = (ComboBox)this.splitContainer1.Panel2.Controls["bearbeiterComboBox"];
            string name = comb.SelectedItem.ToString();

            //changing default view of datatable to reflect comboboxchoice
            form_dt.DefaultView.RowFilter = "Bearbeiter like '" + name + "'";
        }

        private void  AddUIControlsToPanel2ofSplitCOntainer() 
        {
            //int dynamicX = 5;
            //int labelY = 50;
            //int textboxY = labelY + 39;
            try
            {
                foreach (string columnname in headers)
                {
                    Label hinweise = new Label();
                    int offset_text = 50;
                    hinweise.Font = new Font(FontFamily.GenericSansSerif, 9);
                    string text = "Hinweise zum Ausfühlen der Textfelder:\n\n" +
                                    "Vertragsabschluss:".PadRight(30) + "dd.mm.yyyy und nicht leer\n" +
                                    "Projektnummer:".PadRight(31) + "wird automatisch erstellt\n" +
                                    "Projektname:".PadRight(34) + "beliebig, nicht leer\n" +
                                    "Labornummer:".PadRight(31) + "6 stellige Zahl und Buchstaben\n" +
                                    "PSPElement:".PadRight(32) + "15 stellige Zahl\n" +
                                    "Projekpartner:".PadRight(34) + "nur Buchstaben\n" +
                                    "Auftraggeber:".PadRight(34) + "nur Buchstaben\n" +
                                    "Projektanfang:".PadRight(33) + "dd.mm.yyyy\n" +
                                    "Projektanfang:".PadRight(33) + "dd.mm.yyyy\n" +
                                    "Projektvolumen:".PadRight(31) + "ganze oder Dezimalzahl\n" +
                                    "Bearbeiter:".PadRight(36) + "nur Buchstaben und nicht leer\n" +
                                    "Dienstleistung:".PadRight(33) + "x oder kein x\n" +
                                    "Hoheitliche Forschung:".PadRight(25) + "x oder kein x\n" +
                                    "Gutachten:".PadRight(35) + "x oder kein x\n" +
                                    "Freie Forschung:".PadRight(30) + "x oder kein x\n" +
                                    "Status:".PadRight(39) + "nicht leer\n";
                    hinweise.Text = text;
                    SettingUI(columnname);
                    hinweise.AutoSize=true;
                    hinweise.ForeColor = System.Drawing.Color.Green;
                    hinweise.Location = new Point(700, textboxY+30);
                    hinweise.Show();
                    this.splitContainer1.Panel2.Controls.Add(hinweise);


                     
                }

                //adding Add-, Update- and Clear Buttons
                Create_Add_Update_Clear_Buttons_and_StatusCombo();

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                Application.Exit();
            }
            
        }

        private void Create_Add_Update_Clear_Buttons_and_StatusCombo()
        {
            Button addButton = new Button();
            addButton.Name = "AddButton";
            addButton.Text = "Hinzufügen";
            addButton.Location = new Point(5, (textboxY + 40));
            addButton.Click += addButton_Click;
            addButton.Show();
            this.splitContainer1.Panel2.Controls.Add(addButton);


            Button updateButton = new Button();
            updateButton.Name = "UpdateButton";
            updateButton.Text = "Aktualisieren";
            updateButton.Location = new Point(85, (textboxY + 40));
            updateButton.Click += updateButton_Click;
            updateButton.Show();
            this.splitContainer1.Panel2.Controls.Add(updateButton);
            
            
            Button clearButton = new Button();
            clearButton.Name = "ClearButton";
            clearButton.Text = "Textfelder leeren";
            clearButton.AutoSize = true;
            clearButton.Location = new Point(165, (textboxY + 40));
            clearButton.Click += clearButton_Click;
            clearButton.Show();
            this.splitContainer1.Panel2.Controls.Add(clearButton);
            
            
            Button openDirectoryButton = new Button();
            openDirectoryButton.Name = "OpenDirectoryButton";
            openDirectoryButton.Text = "Projektordner öffnen";
            openDirectoryButton.AutoSize = true;
            openDirectoryButton.Location = new Point(270, (textboxY + 40));
            openDirectoryButton.Click += openDirectoryButton_Click;
            openDirectoryButton.Show();
            this.splitContainer1.Panel2.Controls.Add(openDirectoryButton);

            Button changeSourceFileDir = new Button();
            changeSourceFileDir.Name = "changeSourceFileDirButton";
            changeSourceFileDir.Text = "Ort der SQLite-Datenbank auswählen";
            changeSourceFileDir.AutoSize = true;
            changeSourceFileDir.Location = new Point(5, 200);
            changeSourceFileDir.Click += changeSourceFileDir_Click;
            changeSourceFileDir.Enabled = false;
            changeSourceFileDir.Show();
            this.splitContainer1.Panel2.Controls.Add(changeSourceFileDir);

            Button changeProjectsRootFolderTreeDir = new Button();
            changeProjectsRootFolderTreeDir.Name = "changeProjectsRootFolderTreeDirButton";
            changeProjectsRootFolderTreeDir.Text = "Ort des Projektordners auswählen";
            changeProjectsRootFolderTreeDir.AutoSize = true;
            changeProjectsRootFolderTreeDir.Location = new Point(5, 250);
            changeProjectsRootFolderTreeDir.Click += changeProjectsRootFolderTreeDir_Click;
            changeProjectsRootFolderTreeDir.Enabled = false;
            changeProjectsRootFolderTreeDir.Show();
            this.splitContainer1.Panel2.Controls.Add(changeProjectsRootFolderTreeDir);

            Button showLog = new Button();
            showLog.Name = "showLogButton";
            showLog.Text = "Log anzeigen";
            showLog.AutoSize = true;
            showLog.Location = new Point(5, 300);
            showLog.Click += showLog_Click;
            showLog.Show();
            this.splitContainer1.Panel2.Controls.Add(showLog);

            Label StatusLabel = new Label();
            StatusLabel.Name = "StatusLabelComb";
            StatusLabel.Text = "Status auswählen: ";
            StatusLabel.Location = new Point(435, (textboxY + 40));
            StatusLabel.Show();
            this.splitContainer1.Panel2.Controls.Add(StatusLabel);


            ComboBox StatusComboBox = new ComboBox();
            StatusComboBox.Name = "StatusComboBox";
            StatusComboBox.Items.Add("Vorb.");
            StatusComboBox.Items.Add("Antr./Ange.");
            StatusComboBox.Items.Add("beauf./bew.");
            StatusComboBox.Items.Add("abgeschl.");
            StatusComboBox.Location = new Point(535, (textboxY + 40));
            StatusComboBox.SelectedValueChanged += StatusComboBox_SelectedValueChanged;
            StatusComboBox.Show();
            this.splitContainer1.Panel2.Controls.Add(StatusComboBox);
            
            

            TextBox tempTextBox = (TextBox)this.splitContainer1.Panel2.Controls["TextBoxStatus"];
            tempTextBox.Enabled = false;
            //TextBox temp2TextBox = (TextBox)this.splitContainer1.Panel2.Controls["TextBoxProjektnummer"];
            lockedOne = (TextBox)this.splitContainer1.Panel2.Controls["TextBoxProjektnummer"];
            lockedOne.Enabled = false;
        }

        void showLog_Click(object sender, EventArgs e)
        {
            //TODO: show Log
        }

        void changeProjectsRootFolderTreeDir_Click(object sender, EventArgs e)
        {
            //TODO: create filedialog and save result in ProjectsRootFolderTreeDir in FileSystemServices.cs
        }

        void changeSourceFileDir_Click(object sender, EventArgs e)
        {
            //TODO: create filedialog and save result in connectionStringDefault in SQLLite.cs
        }

        public void SettingUI(string columnname)
        {
            Label automaticLabel = new Label();
            TextBox automaticTextBox = new TextBox();
            automaticLabel.Name = "Label" + columnname;
            //automaticLabel.Text = SettingLabelText(columnname);
            automaticLabel.AutoSize = true;
            automaticTextBox.Name = "TextBox" + columnname;
            automaticLabel.Location = new Point(dynamicX, labelY);
            automaticTextBox.Location = new Point(dynamicX, textboxY);

            switch (columnname)
            {
                case "Vertragsabschluss":
                    automaticLabel.Text = "Vertragsabschl.:";
                    automaticTextBox.Width = 70;
                    dynamicX += automaticTextBox.Width + 10;
                    break;

                case "Projektnummer":
                    automaticLabel.Text = "Projektnum.:";
                    automaticTextBox.Width = 55;
                    dynamicX += automaticTextBox.Width + 10;
                    break;

                case "Projektname":
                    automaticLabel.Text = "Projektname:";
                    automaticTextBox.Width = 110;
                    dynamicX += automaticTextBox.Width + 10;
                    break;

                case "Labornummer":
                    automaticLabel.Text = "Labornum.:";
                    automaticTextBox.Width = 55;
                    dynamicX += automaticTextBox.Width + 10;
                    break;

                case "PSPElement":
                    automaticLabel.Text = "PSPEl.:";
                    dynamicX += automaticTextBox.Width + 10;
                    break;

                case "Projekpartner":
                    automaticLabel.Text = "Projekpartner:";
                    dynamicX += automaticTextBox.Width + 10;
                    break;

                case "ProjektträgerUndAuftraggeber":
                    automaticLabel.Text = "Auftraggeber:";
                    dynamicX += automaticTextBox.Width + 10;
                    break;

                case "ProjektAnfang":
                    automaticLabel.Text = "P-Anfang:";
                    automaticTextBox.Width = 70;
                    dynamicX += automaticTextBox.Width + 10;
                    break;

                case "ProjektEnde":
                    automaticLabel.Text = "P-Ende:";
                    automaticTextBox.Width = 70;
                    dynamicX += automaticTextBox.Width + 10;
                    break;

                case "Projektvolumen":
                    automaticLabel.Text = "P-Volumen:";
                    automaticTextBox.Width = 50;
                    dynamicX += automaticTextBox.Width + 10;
                    break;

                case "Bearbeiter":
                    automaticLabel.Text = "Bearbeiter:";
                    automaticTextBox.Width = 40;
                    dynamicX += automaticTextBox.Width + 20;
                    break;

                case "Dienstleistung":
                    automaticLabel.Text = "Dienstleist.:";
                    automaticTextBox.Width = 20;
                    dynamicX += automaticTextBox.Width + 45;
                    break;

                case "HoheitlicheForschung":
                    automaticLabel.Text = "Hoh. Forsch.:";
                    automaticTextBox.Width = 20;
                    dynamicX += automaticTextBox.Width + 45;
                    break;

                case "Gutachten":
                    automaticLabel.Text = "Gutachten:";
                    automaticTextBox.Width = 20;
                    dynamicX += automaticTextBox.Width + 45;
                    break;

                case "FreieForschung":
                    automaticLabel.Text = "Fr. Forsch.:";
                    automaticTextBox.Width = 20;
                    dynamicX += automaticTextBox.Width + 45;
                    break;

                case "Status":
                    automaticLabel.Text = "Status:";
                    automaticTextBox.Width = 60;
                    dynamicX += automaticTextBox.Width + 40;
                    break;

                default:
                    return ;
            }
            automaticLabel.AutoSize = true;
            //making controls visible
            automaticLabel.Show();
            automaticTextBox.Show();
            //making controls persistant
            this.splitContainer1.Panel2.Controls.Add(automaticLabel);
            this.splitContainer1.Panel2.Controls.Add(automaticTextBox);
            
        }


        //public string SettingLabelText(string columnname)
        //{
        //    switch (columnname)
        //    {   
        //        case "Vertragsabschluss":
        //            return "Vertragsabschluss:\ndd.mm.yyyy und\nnicht leer";

        //        case "Projektnummer":
        //            return "Projektnummer: ";

        //        case "Projektname":
        //            return "Projektname:\nbeliebig, nicht leer";

        //        case "Labornummer":
        //            return "Labornummer:\n6 stellige Zahl und\nBuchstaben";

        //        case "PSPElement":
        //            return "PSPElement:\n15 stellige Zahl";

        //        case "Projekpartner":
        //            return "Projekpartner:\nnur Buchstaben";

        //        case "ProjektträgerUndAuftraggeber":
        //            return "Auftraggeber:\nnur Buchstaben";

        //        case "ProjektAnfang":
        //            return "Projektanfang:\ndd.mm.yyyy";

        //        case "ProjektEnde":
        //            return "Projektende:\ndd.mm.yyyy";

        //        case "Projektvolumen":
        //            return "Projektvolumen:\nganze oder\nDezimalzahl";

        //        case "Bearbeiter":
        //            return "Bearbeiter:\nnur Buchstaben und nicht leer";

        //        case "Dienstleistung":
        //            return "Dienstleistung:\nx oder kein x";

        //        case "HoheitlicheForschung":
        //            return "Hoheitliche Forschung:\nx oder kein x";

        //        case "Gutachten":
        //            return "Gutachten:\nx oder kein x";

        //        case "FreieForschung":
        //            return "Freie Forschung:\nx oder kein x";

        //        case "Status":
        //            return "Status:\nnicht leer";

        //        default:
        //            return "";
        //    }
        //}

        void StatusComboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            ComboBox temp = (ComboBox)this.splitContainer1.Panel2.Controls["StatusComboBox"];
            if (temp.Items[temp.SelectedIndex] != null)
            {
                TextBox tempTextBox = (TextBox)this.splitContainer1.Panel2.Controls["TextBoxStatus"];
                tempTextBox.Text = temp.Items[temp.SelectedIndex].ToString();
            }
        }

        void openDirectoryButton_Click(object sender, EventArgs e)
        {
            //TODO: open directory
            GetCurrentTextBoxValues();
            FileSystemServices.OpenProjectFolder(currentTextBoxValues["Projektnummer"], currentTextBoxValues["Projektname"]);
        }

        void clearButton_Click(object sender, EventArgs e)
        {
            foreach (string columnname in headers)
            {
                TextBox tempTextBox = (TextBox)this.splitContainer1.Panel2.Controls["TextBox" + columnname];
                tempTextBox.Clear();
            }
        }

        void updateButton_Click(object sender, EventArgs e)
        {
            MessageBox.Show(((Button)sender).Name);

            //geting current values of the TextBoxes
            
            GetCurrentTextBoxValues();
             
            //ckecks if project is "abgeschl." and if not unlocks all textboxes
            if (!currentTextBoxValues["Status"].Equals("abgeschl."))
            {
                UnlockEditing();
            }

            // validate current values
            if (EntryValidation.ValidateInput(currentTextBoxValues))
            {

                DialogResult dialogResult = MessageBox.Show(dialogBoxMessage, "Bestätigung der Aktualisierung", MessageBoxButtons.YesNo);

                if (dialogResult == DialogResult.Yes)
                {
                    form_dt = SQLiteLogic.UpdateDataBaseFile(previousTextBoxValues, currentTextBoxValues);

                    ExcelGridView.DataSource = form_dt;
                    dialogBoxMessage = string.Empty;
                }
   
            }
            else
            {
                MessageBox.Show("Die Eingaben entsprechen nicht dem Format. Probieren Sie bitte noch Mal");
            }

        }

        
        private void GetCurrentTextBoxValues()
        {
            currentTextBoxValues.Clear();
            dialogBoxMessage = string.Empty;
            dialogBoxMessage = "Vorher:                     Nachher:\n";
            foreach (string columnname in headers)
            {
                TextBox tempTextBox = (TextBox)this.splitContainer1.Panel2.Controls["TextBox" + columnname];
                currentTextBoxValues.Add(columnname, tempTextBox.Text);
                dialogBoxMessage += string.Format("Vorher {0}: {1} und nachher {2}: {3} \n", columnname, previousTextBoxValues[columnname], columnname, currentTextBoxValues[columnname]);                    
            }
        }

        //locks down all text boxes except status textbox 
        private void LockEditing()
        {
            foreach (string columnname in headers)
            {
                TextBox tempTextBox = (TextBox)this.splitContainer1.Panel2.Controls["TextBox" + columnname];
                if (!tempTextBox.Name.Equals("TextBoxStatus") || !tempTextBox.Name.Equals("TextBoxProjektnummer"))
                {
                    tempTextBox.Enabled = false;
                }
            }
        }


        //unlock locked textboxes
        private void UnlockEditing()
        {
            foreach (string columnname in headers)
            {
                TextBox tempTextBox = (TextBox)this.splitContainer1.Panel2.Controls["TextBox" + columnname];
                if (!tempTextBox.Name.Equals("TextBoxStatus") || !tempTextBox.Name.Equals("TextBoxProjektnummer"))
                {
                    tempTextBox.Enabled = true;
                }
            }
            TextBox locked = (TextBox)this.splitContainer1.Panel2.Controls["TextBoxProjektnummer"];
            locked.Enabled = false;
        }

        void addButton_Click(object sender, EventArgs e)
        {
            GetCurrentTextBoxValues();

            string currentYear = (System.DateTime.Now).ToString().Substring(6, 4);
            int projectNr = 0;
            foreach (DataRow item in form_dt.Rows)
            {
                string numb = item[1].ToString();
                if (numb.Contains(currentYear))
                {
                    MessageBox.Show(numb.Substring(4));
                    int count = int.Parse(numb.Substring(4));
                    if (projectNr<count)
                    {
                        projectNr = count;
                    }
                }
                
            }

            projectNr += 1;
            string projectNrAsStr = string.Empty;
            if (projectNr >9)
            {
                projectNrAsStr = currentYear+ "" + projectNr;
            }
            else
            {
                projectNrAsStr = currentYear + "0" + projectNr;
            }

            MessageBox.Show("Nechst Project would be: " + projectNrAsStr);
            currentTextBoxValues["Projektnummer"] = projectNrAsStr;

            if (EntryValidation.ValidateInput(currentTextBoxValues))
            {
                if (currentTextBoxValues.Keys.Contains("Projektnummer") &&
                    currentTextBoxValues.Keys.Contains("Projektname") && 
                    currentTextBoxValues.Keys.Contains("Bearbeiter"))
                {
                    //TODO: check if entry already exist
                    foreach (DataRow row in form_dt.Rows)
                    {
                        if (row["Projektname"].ToString().Equals(currentTextBoxValues["Projektname"]) && row["Bearbeiter"].ToString().Equals(currentTextBoxValues["Bearbeiter"]))
                        {
                            MessageBox.Show("Dieses Projekt gibt es schon", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }

                    int colaffected = SQLiteLogic.AddBearbeiterToTheDataBase(currentTextBoxValues);
                    if (colaffected==1)
                    {
                        form_dt = SQLiteLogic.FillDataTable();
                        FileSystemServices.CreateFolders(currentTextBoxValues["Projektnummer"], currentTextBoxValues["Projektname"]);
                        ExcelGridView.DataSource = form_dt;
                        MessageBox.Show("Erfolgreich hinzugefügt", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Hinzufügen fehlgeschlagen", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Es müssen mindestens Projektname- und Bearbeiter Felder gefüllt sein damit man Projekthinzufügen kann.","Fehler",MessageBoxButtons.OK,MessageBoxIcon.Error);
                }
                
            }
            else
            {
                MessageBox.Show("Die Eingaben entsprechen nicht dem Format. Probieren Sie bitte noch Mal", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void ExcelGridView_SelectionChanged(object sender, EventArgs e)
        {
            
            if (ExcelGridView.SelectedRows.Count >0)
            {
                
                previousTextBoxValues = new Dictionary<string, string>();
              
                foreach (string columnName in headers)
                {
                    String cellText=string.Empty;
                    //geting current value of a cell in the gridview
                    if (ExcelGridView.SelectedRows[0].Cells[columnName] != null)
                    {
                        cellText = ExcelGridView.SelectedRows[0].Cells[columnName].Value.ToString(); 
                    }
                    
                    //saving the gridview cell value to the TextBox
                    TextBox tempTextBox = (TextBox)this.splitContainer1.Panel2.Controls["TextBox" + columnName];
                    tempTextBox.Text = cellText;

                    //saving the gridview cell value to internal list
                    previousTextBoxValues.Add(columnName,cellText);     
                }
                TextBox statusTextBox = (TextBox)this.splitContainer1.Panel2.Controls["TextBoxStatus"];
                if (statusTextBox.Text.Equals("abgeschl."))
                {
                    LockEditing();
                }
                else
                {
                    UnlockEditing();
                }
            }
                
        }

        private void ExcelGridView_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            var height = 20;
            foreach (DataGridViewRow dr in ExcelGridView.Rows)
            {
                height += dr.Height;
            }

            ExcelGridView.Height = height;
            splitContainer1.SplitterDistance = height;
            this.Height = splitContainer1.SplitterDistance + splitContainer1.Panel2.Height;

        }


    }
}
