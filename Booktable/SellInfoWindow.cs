using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace Booktable
{
    public partial class SellInfoWindow : Form
    {
        List<string> owndata = new List<string>();
        List<string> colnames = new List<string>();


        public SellInfoWindow(List<string> colnames, List<string> onedatalist)
        {
            InitializeComponent();
            this.owndata = onedatalist;
            this.colnames = colnames;
            this.showInfos();
        }

        private void showInfos()
        {
            this.book_label.Text = this.owndata[this.colnames.IndexOf("책이름")];
            this.author_label.Text = this.owndata[this.colnames.IndexOf("저자")];
            //this.realprice_label.Text = this.owndata[this.colnames.IndexOf("정가")];
            this.sellprice_label.Text = this.owndata[this.colnames.IndexOf("판매가")];
            this.whobuy_label.Text = this.owndata[this.colnames.IndexOf("구매자")];
            this.extra_label.Text = this.owndata[this.colnames.IndexOf("비고")];
            this.whensold_label.Text = this.owndata[this.colnames.IndexOf("판매시간")];
            this.howsold_label.Text = this.owndata[this.colnames.IndexOf("결제방법")];

        }
    }

    
}
