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
            int dynamicX = 5;
            int labelY = 50;
            int textboxY = labelY + 39;
            try
            {
                foreach (string columnname in headers)
                {
                    // creation of controls
                    Label automaticLabel = new Label();
                    TextBox automaticTextBox = new TextBox();
                    // naming of controls
                    automaticLabel.Name = "Label" + columnname;
                    automaticLabel.Text = SettingLabelText(columnname);
                    automaticLabel.AutoSize = true;
                    //automaticLabel.Text = columnname + ": ";
                    automaticTextBox.Name = "TextBox" + columnname;
                    // positioning of controls
                    automaticLabel.Location = new Point(dynamicX, labelY);
                    automaticTextBox.Location = new Point(dynamicX, textboxY);
                    //making controls visible
                    automaticLabel.Show();
                    automaticTextBox.Show();
                    //making controls persistant
                    this.splitContainer1.Panel2.Controls.Add(automaticLabel);
                    this.splitContainer1.Panel2.Controls.Add(automaticTextBox);
                    //offseting controls
                    if (automaticLabel.Name.Equals("LabelProjektträgerUndAuftraggeber") || automaticLabel.Name.Equals("LabelHoheitlicheForschung"))
                    {
                        dynamicX += 140;
                    }
                    else
                    {
                        dynamicX += 110;
                    }

                     
                }

                //adding Add-, Update- and Clear Buttons
                Button addButton = new Button();
                Button updateButton = new Button();
                Button clearButton = new Button();
                Button openDirectoryButton = new Button();

                Label StatusLabel = new Label();

                ComboBox StatusComboBox = new ComboBox();

                StatusComboBox.Items.Add("Vorb.");
                StatusComboBox.Items.Add("Antr./Ange.");
                StatusComboBox.Items.Add("beauf./bew.");
                StatusComboBox.Items.Add("abgel.");

                addButton.Name = "AddButton";
                updateButton.Name = "UpdateButton";
                clearButton.Name = "ClearButton";
                openDirectoryButton.Name = "OpenDirectoryButton";
                StatusLabel.Name = "StatusLabelComb";
                StatusComboBox.Name = "StatusComboBox";

                addButton.Text = "Hinzufügen";
                updateButton.Text = "Aktualisieren";
                clearButton.Text = "Textfelder leeren";
                openDirectoryButton.Text = "Projektordner öffnen";
                clearButton.AutoSize = true;
                openDirectoryButton.AutoSize = true;
                StatusLabel.Text = "Status auswählen: ";

                addButton.Location = new Point(5, (textboxY + 40));
                updateButton.Location = new Point(85, (textboxY + 40));
                clearButton.Location = new Point(165, (textboxY + 40));
                openDirectoryButton.Location = new Point(270, (textboxY + 40));
                StatusLabel.Location = new Point(1535, (textboxY + 40));
                StatusComboBox.Location = new Point(1635, (textboxY + 40));

                addButton.Click += addButton_Click;
                updateButton.Click += updateButton_Click;
                clearButton.Click += clearButton_Click;
                openDirectoryButton.Click += openDirectoryButton_Click;
                StatusComboBox.SelectedValueChanged += StatusComboBox_SelectedValueChanged;


                addButton.Show();
                updateButton.Show();
                clearButton.Show();
                openDirectoryButton.Show();
                StatusLabel.Show();
                StatusComboBox.Show();


                this.splitContainer1.Panel2.Controls.Add(addButton);
                this.splitContainer1.Panel2.Controls.Add(updateButton);
                this.splitContainer1.Panel2.Controls.Add(clearButton);
                this.splitContainer1.Panel2.Controls.Add(openDirectoryButton);
                this.splitContainer1.Panel2.Controls.Add(StatusLabel);
                this.splitContainer1.Panel2.Controls.Add(StatusComboBox);

                TextBox tempTextBox = (TextBox)this.splitContainer1.Panel2.Controls["TextBoxStatus"];
                tempTextBox.Enabled = false;
                TextBox temp2TextBox = (TextBox)this.splitContainer1.Panel2.Controls["TextBoxProjektnummer"];
                temp2TextBox.Enabled = false;

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                Application.Exit();
            }
            
        }
        
        public string SettingLabelText(string columnname)
        {
            switch (columnname)
            {   
                case "Vertragsabschluss":
                    return "Vertragsabschluss: \ndd.mm.yyyy und\nnicht leer";

                case "Projektnummer":
                    return "Projektnummer: ";

                case "Projektname":
                    return "Projektname: \nbeliebig, nicht leer";

                case "Labornummer":
                    return "Labornummer: \n6 stellige Zahl und\nBuchstaben";

                case "PSPElement":
                    return "PSPElement: \n15 stellige Zahl";

                case "Projekpartner":
                    return "Projekpartner: \n keine Zahl,\nsonst beliebig";

                case "ProjektträgerUndAuftraggeber":
                    return "Projektträger/Auftraggeber: \nkeine Zahl,\nsonst beliebig";

                case "ProjektAnfang":
                    return "Projektanfang: \ndd.mm.yyyy";

                case "ProjektEnde":
                    return "Projektende: \ndd.mm.yyyy";

                case "Projektvolumen":
                    return "Projektvolumen: \nganze oder\nDezimalzahl";

                case "Bearbeiter":
                    return "Bearbeiter: \nkeine Zahl, nicht leer";

                case "Dienstleistung":
                    return "Dienstleistung: \nx oder kein x";

                case "HoheitlicheForschung":
                    return "Hoheitliche Forschung: \nx oder kein x";

                case "Gutachten":
                    return "Gutachten: \nx oder kein x";

                case "FreieForschung":
                    return "Freie Forschung: \nx oder kein x";

                case "Status":
                    return "Status: \nnicht leer";

                default:
                    return "";
            }
        }

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
                if (!tempTextBox.Name.Equals("TextBoxStatus"))
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
                if (!tempTextBox.Name.Equals("TextBoxStatus"))
                {
                    tempTextBox.Enabled = true;
                }
            }
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
