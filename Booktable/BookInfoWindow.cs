using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Booktable
{
    public partial class BookInfoWindow : Form
    {
        string bookname = null;
        bool IsSearch = false;

        public BookInfoWindow(string book)
        {
            InitializeComponent();
            this.bookname = book;
            this.webBrowser1.Navigate("http://ivp.co.kr/books/s01.html");
            this.webBrowser1.DocumentCompleted += this.webBrowser1_DocumentCompleted;

        }
           

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (!this.IsSearch)
            {
                // Get search component
                if (this.webBrowser1.Document != null)
                {
                    HtmlDocument doc = this.webBrowser1.Document;

                    if (doc != null)
                    {
                        HtmlElementCollection eles = doc.GetElementsByTagName("input");

                        foreach (HtmlElement ele in eles)
                        {

                            if (ele.GetAttribute("className").Equals("frm_input"))
                            {
                                Console.WriteLine(ele.GetAttribute("name"));
                                ele.SetAttribute("Value", this.bookname);
                            }
                            else if (ele.GetAttribute("className").Equals("btn_submit"))
                            {
                                ele.InvokeMember("click");
                            }
                        }
                    }
                }
                this.IsSearch = true;
            }
        }
    }
}
