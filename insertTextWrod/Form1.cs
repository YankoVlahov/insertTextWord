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
using Aspose.Words;
using Microsoft.Office.Interop.Word;
using word= Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Globalization;
using System.Threading;

namespace insertTextWrod
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();
            using (OpenFileDialog wordRead = new OpenFileDialog())

            {
                if (wordRead.ShowDialog() == DialogResult.OK)
                {

                    object file = wordRead.FileName;

                    //MessageBox.Show(file);
                    word.Application wordAPp = new word.Application();
                    //  word.Document aDoc = null;
                    object missing = System.Reflection.Missing.Value;
                    wordAPp.Visible = true;
                    Microsoft.Office.Interop.Word.Document docs = wordAPp.Documents.Open
                    (ref file, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing);


                    // wordRead.OpenFile();


                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            CreateDocument();
        }

        private void CreateDocument()
        {
            try
            {
                //Create an instance for word app  
                Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();

                //Set animation status for word application  
                winword.ShowAnimation = false;


                winword.Visible = false;


                object missing = System.Reflection.Missing.Value;


                Microsoft.Office.Interop.Word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);



                // Save the document
                object filename = @"D:\temp2.doc";
                document.SaveAs2(ref filename);
                document.Close(ref missing, ref missing, ref missing);
                document = null;
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;
                MessageBox.Show("Document created successfully !");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }



        public partial class SqlDb
        {
            
            public string Ip { get; set; }
            public string Name { get; set; }
            public string Position { get; set; }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            //brow на word
            

                   

                    List<SqlDb> filTheText = new List<SqlDb>();
                    var data = new SqlDb { Ip = textBox1.Text, Name = textBox2.Text, Position = "чистач на кенефи" };
                    filTheText.Add(data);

                    //string name = textBox1.Text;
                    //string position = textBox2.Text;
                    //string department = textBox3.Text;
                    //string from_date = dateTimePicker1.Text;

                    ExportWord(filTheText);
        }

        private void ExportWord(List<SqlDb> filTheText)
        {
            using (OpenFileDialog wordRead = new OpenFileDialog())

            {
                if (wordRead.ShowDialog() == DialogResult.OK)
                {
                    object file = wordRead.FileName;
                    //word.Application wordAPp = new word.Application();
                    object missing = System.Reflection.Missing.Value;
                    object readOnly = false;
                    var fileName = Path.Combine(System.Windows.Forms.Application.StartupPath, wordRead.FileName);
                    Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application { Visible = false };
                    Microsoft.Office.Interop.Word.Document aDoc = wordApp.Documents.Open(fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                    // Microsoft.Office.Interop.Word.Document aDoc = wordApp.Documents.Open();
                    foreach (var item in filTheText)
                    {
                        foreach (var prop in item.GetType().GetProperties())
                        {

                            FindAndReplace(wordApp, prop.Name, prop.GetValue(item, null));
                        }
                    }
                    //aDoc.Activate();
                    wordApp.Visible = true;
                }
            }
        }

        private void FindAndReplace(Microsoft.Office.Interop.Word.Application fileOpen , object findText, object replaceWithText)
                {
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
                    object visible = false;
                    object replace = 2;
                    object wrap = 1;
            //find and replace 
            fileOpen.Selection.SetRange(fileOpen.ActiveDocument.Content.Start, fileOpen.ActiveDocument.Content.End);
                    fileOpen.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                        ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                        ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
                }

        

        private void Form1_Load(object sender, EventArgs e)
        {
            DateTime dt1 = dateTimePicker1.Value;
            string theDate = dateTimePicker1.Value.ToString("dd.MM.yyyy");
           // MessageBox.Show(theDate);

        }
    }
}
    

    
    


