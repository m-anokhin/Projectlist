using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace ExcelProject
{
    class EntryValidation
    {
        public static bool ValidateInput(Dictionary<string, string> formInput)
        {
            foreach (string columnName in formInput.Keys)
            {
                switch (columnName)
                {   
                    case "Vertragsabschluss":
                        if (!Regex.IsMatch(formInput[columnName], @"^\d{2}\.\d{2}\.\d{4}$") && !formInput[columnName].Equals(string.Empty))
                        {
                            System.Windows.Forms.MessageBox.Show("ERROR: "+ columnName + ": " + formInput[columnName]+ " entspricht nicht den Validierungsregeln");
                            return false;
                        }
                        break;
                        
                    case "Projektnummer":
                        if (!Regex.IsMatch(formInput[columnName], @"^\d{6}$"))
                        {
                            System.Windows.Forms.MessageBox.Show("ERROR: " + columnName + ": " + formInput[columnName] + " entspricht nicht den Validierungsregeln");
                            return false;
                        }
                        break;
                        
                    case "Projektname":
                        if (formInput[columnName].Equals(string.Empty))
                        {
                            System.Windows.Forms.MessageBox.Show("ERROR: " + columnName + ": " + formInput[columnName] + " entspricht nicht den Validierungsregeln");
                            return false;
                        }
                        break;
                        
                    case "Labornummer":
                        if (!Regex.IsMatch(formInput[columnName], @"^\d{6}\w*$") && !formInput[columnName].Equals(string.Empty))
                        {
                            System.Windows.Forms.MessageBox.Show("ERROR: " + columnName + ": " + formInput[columnName] + " entspricht nicht den Validierungsregeln");
                            return false;
                        }
                        break;

                    case "PSPElement":
                        if (!Regex.IsMatch(formInput[columnName], @"^\d{15}$") && !formInput[columnName].Equals(string.Empty))
                        {
                            System.Windows.Forms.MessageBox.Show("ERROR: " + columnName + ": " + formInput[columnName] + " entspricht nicht den Validierungsregeln");
                            return false;
                        }
                        break;
                        
                    case "Projekpartner":
                        if (!formInput[columnName].Equals(string.Empty) && Regex.IsMatch(formInput[columnName], @"^\d$"))
                        {
                            System.Windows.Forms.MessageBox.Show("ERROR: " + columnName + ": " + formInput[columnName] + " entspricht nicht den Validierungsregeln");
                            return false;
                        }
                        break;

                    case "ProjektträgerUndAuftraggeber":
                        if (formInput[columnName].Equals(string.Empty) && Regex.IsMatch(formInput[columnName], @"^\d$"))
                        {
                            System.Windows.Forms.MessageBox.Show("ERROR: " + columnName + ": " + formInput[columnName] + " entspricht nicht den Validierungsregeln");
                            return false;
                        }
                        break;

                    case "ProjektAnfang":
                        if (!Regex.IsMatch(formInput[columnName], @"^\d{2}\.\d{2}\.\d{4}$") && !formInput[columnName].Equals(string.Empty))
                        {
                            System.Windows.Forms.MessageBox.Show("ERROR: " + columnName + ": " + formInput[columnName] + " entspricht nicht den Validierungsregeln");
                            return false;
                        }
                        break;

                    case "ProjektEnde":
                        if (!Regex.IsMatch(formInput[columnName], @"^\d{2}\.\d{2}\.\d{4}$") && !formInput[columnName].Equals(string.Empty))
                        {
                            System.Windows.Forms.MessageBox.Show("ERROR: " + columnName + ": " + formInput[columnName] + " entspricht nicht den Validierungsregeln");
                            return false;
                        }
                        break;
                        
                    case "Projektvolumen":
                        if (!Regex.IsMatch(formInput[columnName], @"^[0-9]+\,?[0-9]*$") && !formInput[columnName].Equals(string.Empty))
                        {
                            System.Windows.Forms.MessageBox.Show("ERROR: " + columnName + ": " + formInput[columnName] + " entspricht nicht den Validierungsregeln");
                            return false;
                        }
                        break;
                        
                    case "Bearbeiter":
                        if (Regex.IsMatch(formInput[columnName], @"^\d$") &&!formInput[columnName].Equals(string.Empty) )
                        {
                            System.Windows.Forms.MessageBox.Show("ERROR: " + columnName + ": " + formInput[columnName] + " entspricht nicht den Validierungsregeln");
                            return false;
                        }
                        break;
                        
                    case "Dienstleistung":
                        if (!Regex.IsMatch(formInput[columnName], @"^x?$"))
                        {
                            System.Windows.Forms.MessageBox.Show("ERROR: " + columnName + ": " + formInput[columnName] + " entspricht nicht den Validierungsregeln");
                            return false;
                        }
                        break;
                        
                    case "HoheitlicheForschung":
                        if (!Regex.IsMatch(formInput[columnName], @"^x?$"))
                        {
                            System.Windows.Forms.MessageBox.Show("ERROR: " + columnName + ": " + formInput[columnName] + " entspricht nicht den Validierungsregeln");
                            return false;
                        }
                        break;
                        
                    case "Gutachten":
                        if (!Regex.IsMatch(formInput[columnName], @"^x?$"))
                        {
                            System.Windows.Forms.MessageBox.Show("ERROR: " + columnName + ": " + formInput[columnName] + " entspricht nicht den Validierungsregeln");
                            return false;
                        }
                        break;
                        
                    case "FreieForschung":
                        if (!Regex.IsMatch(formInput[columnName], @"^x?$"))
                        {
                            System.Windows.Forms.MessageBox.Show("ERROR: " + columnName + ": " + formInput[columnName] + " entspricht nicht den Validierungsregeln");
                            return false;
                        }
                        break;
                        
                    case "Status":
                        if (formInput[columnName].Equals(string.Empty))
                        {
                            System.Windows.Forms.MessageBox.Show("ERROR: " + columnName + ": " + formInput[columnName] + " entspricht nicht den Validierungsregeln");
                            return false;
                        }
                        break;
                        
                    default:
                        return false;
                        
                       
                }
                
            }
            return true;
        }
    }
}
