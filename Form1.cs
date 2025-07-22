using AddressParser.Helpers;
using System;
using System.Windows.Forms;

namespace AddressParser
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void ButtonLoadFile_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel 檔案 (*.xlsx)|*.xlsx",
                Title = "選擇 Excel 檔案"
            };



            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string path = openFileDialog.FileName;
                ExcelAddressTranslator.TranslateColumnAtoB(path);

                MessageBox.Show("地址翻譯完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            return;
        }
    }
}
