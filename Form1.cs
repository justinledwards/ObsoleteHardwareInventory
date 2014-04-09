using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;



namespace Surplusing
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void countyBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void typeBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void stateBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void savePrintButton_Click(object sender, EventArgs e)
        {
            
            var surplusEntry = new Surplus {
                  Type = typeBox.Text,
                  Court = courtBox.Text,
                  County = countyBox.Text,
                  State = stateBox.Text,
                  Serial = serialBox.Text,
                  Make = makeBox.Text,
                  Model = modelBox.Text,
                  Useable_Parts = useablePartsBox.Text,
                  Reuseable_Equipment = reuseableEquipmentBox.Text,
                  Reason_for_Surplus = surplusReasonBox.Text,
                  TimeStamp = DateTime.Now.ToString("MM-dd-yyyy hh:mm:ss")
            };
            SaveToExcel(surplusEntry);
            PrintLabel(surplusEntry);
            typeBox.Text = "";
            courtBox.Text = "";
            countyBox.Text = "";
            stateBox.Text = "";
            serialBox.Text = "";
            makeBox.Text = "";
            modelBox.Text = "";
            useablePartsBox.Text = "";
            reuseableEquipmentBox.Text = "";
            surplusReasonBox.Text = "";
            typeBox.Select();
            typeBox.Focus();

        }

        public class Surplus
        {
            public string Type { get; set; }
            public string Court { get; set; }
            public string County { get; set; }
            public string State { get; set; }
            public string Serial { get; set; }
            public string Make { get; set; }
            public string Model { get; set; }
            public string Useable_Parts { get; set; }
            public string Reuseable_Equipment { get; set; }
            public string Reason_for_Surplus { get; set; }
            public string TimeStamp { get; set; }

        }


        static void SaveToExcel(Surplus surplus)
        {
            // Set file location to desktop
            string savefolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string xlfilename = savefolder + "\\surplus.xlsx";
            var excelApp = new Excel.Application();
            // Make excel invisible and turn off "file already exists alert before save".
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;


            // Check if file exists, if not create it and then open it at the end.
            if (File.Exists(xlfilename))
            {
            }
            else
            {
                var newworkbook = excelApp.Workbooks.Add();
                newworkbook.SaveAs(xlfilename, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            };
            var workbook = excelApp.Workbooks.Open(xlfilename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);


            // Just set the column header each time, less code and work right now than checking it all and then setting it. 
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            workSheet.Cells[1, "A"] = "Type";
            workSheet.Cells[1, "B"] = "Court";
            workSheet.Cells[1, "C"] = "County";
            workSheet.Cells[1, "D"] = "State";
            workSheet.Cells[1, "E"] = "Serial";
            workSheet.Cells[1, "F"] = "Make";
            workSheet.Cells[1, "G"] = "Model";
            workSheet.Cells[1, "H"] = "Useable Parts";
            workSheet.Cells[1, "I"] = "Reuseable Equipment";
            workSheet.Cells[1, "J"] = "Reason for Surplus";
            workSheet.Cells[1, "K"] = "TimeStamp";

            var row = workSheet.UsedRange.Row +workSheet.UsedRange.Rows.Count - 1;

            row++;
            workSheet.Cells[row, "A"] = surplus.Type;
            workSheet.Cells[row, "B"] = surplus.Court;
            workSheet.Cells[row, "C"] = surplus.County;
            workSheet.Cells[row, "D"] = surplus.State;
            workSheet.Cells[row, "E"] = surplus.Serial;
            workSheet.Cells[row, "F"] = surplus.Make;
            workSheet.Cells[row, "G"] = surplus.Model;
            workSheet.Cells[row, "H"] = surplus.Useable_Parts;
            workSheet.Cells[row, "I"] = surplus.Reuseable_Equipment;
            workSheet.Cells[row, "J"] = surplus.Reason_for_Surplus;
            workSheet.Cells[row, "K"] = surplus.TimeStamp;
            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].AutoFit();
            workSheet.Columns[3].AutoFit();
            workSheet.Columns[4].AutoFit();
            workSheet.Columns[5].AutoFit();
            workSheet.Columns[6].AutoFit();
            workSheet.Columns[7].AutoFit();
            workSheet.Columns[8].AutoFit();
            workSheet.Columns[9].AutoFit();
            workSheet.Columns[10].AutoFit();
            workSheet.Columns[11].AutoFit();
            
            workbook.SaveAs(xlfilename, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
             false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
             Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workbook.Close();

        }


        public void FindAndReplace(Word.Application doc, string findText, string replaceWithText)
        {
            
            //options
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            //execute find and replace
            doc.Selection.Find.Execute(findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }
        public void PrintLabel(Surplus surplus)
        {
            string savefolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string wordFilename = savefolder + "\\surplus.docx";
            Word.Application wordApp = new Word.Application { Visible = false };
            wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            Word.Document wordDoc = wordApp.Documents.Open(wordFilename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            wordDoc.Activate();

            FindAndReplace(wordApp, "{type}", surplus.Type);
            FindAndReplace(wordApp, "{court}", surplus.Court);
            FindAndReplace(wordApp, "{county}", surplus.County);
            FindAndReplace(wordApp, "{state}", surplus.State);
            FindAndReplace(wordApp, "{serial}", surplus.Serial);
            FindAndReplace(wordApp, "{make}", surplus.Make);
            FindAndReplace(wordApp, "{model}", surplus.Model);
            FindAndReplace(wordApp, "{useable_parts}", surplus.Useable_Parts);
            FindAndReplace(wordApp, "{reuseable}", surplus.Reuseable_Equipment);
            //FindAndReplace(wordApp, "{type}", surplus.Reason_for_Surplus);
            FindAndReplace(wordApp, "{date}", DateTime.Now.ToString("MM-dd-yyyy"));
            wordDoc.PrintOut();
            wordDoc.Close(SaveChanges: false);
            wordApp.Quit(Type.Missing, Type.Missing, Type.Missing);

        }
    }
}
