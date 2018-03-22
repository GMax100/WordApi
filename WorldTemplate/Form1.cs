using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace WorldTemplate
{
    public partial class Form1 : Form
    {
        //private readonly string TemplateFileName = @"D:\File.docx";//путь к файлу
        private Word.Application wordapp;
        private Word.Document worddocument;
        


        public Form1()
        {
            InitializeComponent();
          
        }
        //метод замены заглушек
        private void ReplaceWordStub(string stubToReplace, string text, Word.Document worddocument)
        {
            var range = worddocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text);
        }
        //кнопка открытия
        private void btnStart_Click(object sender, EventArgs e)
        {
            //открытие файла
            #region
            try
            {
                wordapp = new Word.Application();//запускаем Word
                wordapp.Visible = false;//делаем не видимым его
                Object filename = @"D:\A1.docx";
                worddocument = wordapp.Documents.Open(ref filename);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            #endregion

            //переменные с формы
            #region
          
            string bosname = textBox1.Text;
            string postion = textBox2.Text;
            string applicant = textBox3.Text;
            var date = endData.Value.ToShortDateString();

            #endregion


            ReplaceWordStub("{bosname}", bosname, worddocument);
            ReplaceWordStub("{position}", postion, worddocument);
            ReplaceWordStub("{applicant}", applicant, worddocument);
            ReplaceWordStub("{date}", date, worddocument);
            //сохраняем документ
            worddocument.SaveAs2(@"D:\result.docx");
        }
        //кнопка закрытия
        private void button2_Click(object sender, EventArgs e)
        {
            Object saveChanges = Word.WdSaveOptions.wdSaveChanges;
            Object originalFormat = Word.WdOriginalFormat.wdWordDocument;
            Object routeDocument = Type.Missing;
            wordapp.Quit(ref saveChanges,ref originalFormat, ref routeDocument);
            wordapp = null;
        }

        
    }
}
