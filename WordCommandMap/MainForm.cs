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

namespace WordCommandMap {
    public partial class MainForm : Form {
        public MainForm() {
            InitializeComponent();

            //RibbonPanel rp = new RibbonPanel("TEST");
            //Controls.Add(rp);
        }

        private void button1_Click(object sender, EventArgs e) {
            ribbon1.Hide();
            /*Word.Application app = new Word.Application();
            app.Visible = true;
            app.Documents.Open(@"chiParallelInterfaces-1-js.docx");
            app.ActiveDocument.CommandBars.FindControl(Id: 109).Execute();*/
        }
    }
}
