using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;
using System.IO;

using System.Runtime.Serialization;

using SpellCheck.Common;
using SpellCheck.Common.XmlData;


namespace SpellCheck
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();


        }

        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                DataGridViewCell item = row.Cells["AddToExclusion"];

                if (item is DataGridViewCheckBoxCell)
                {
                    DataGridViewCheckBoxCell checkBoxCell = item as DataGridViewCheckBoxCell;
                    checkBoxCell.Value = true;
                }

            }
        }

        private void createCustomDictionaryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Collections.Specialized.StringCollection strings = new System.Collections.Specialized.StringCollection();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                DataGridViewCell item = row.Cells["AddToExclusion"];

                if (item is DataGridViewCheckBoxCell)
                {
                    DataGridViewCheckBoxCell checkBoxCell = item as DataGridViewCheckBoxCell;

                    bool add = (bool)checkBoxCell.EditedFormattedValue;
                    if (add)
                    {
                        string str = row.Cells["Error"].Value as string;

                        if (str != null && !strings.Contains(str))
                        {
                            strings.Add(str);
                        }
                    }
                }
            }

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                StreamWriter writer = new StreamWriter(saveFileDialog1.FileName, false, Encoding.Unicode);
                foreach (string str in strings)
                { 
                    writer.WriteLine(str);
                }
                writer.Close();
            }
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileStream fs2 = new FileStream(openFileDialog1.FileName, FileMode.Open);
                XmlDictionaryReader reader = XmlDictionaryReader.CreateTextReader(fs2, new XmlDictionaryReaderQuotas());
                
                SpellingErrors errors = new SpellingErrors();
                try
                {
                    DataContractSerializer serializer = new DataContractSerializer(typeof(SpellingErrors));
                    errors = (SpellingErrors)serializer.ReadObject(reader, false);
                }
                catch (InvalidOperationException exception)
                {
                    MessageBox.Show(this, exception.Message, "Invalid Xml file");
                    return;
                }
                finally
                {
                    reader.Close();
                    fs2.Close();
                }

                foreach (string strongName in errors.HandlerTable.Keys)
                {
                    Handler handler = errors.HandlerTable[strongName];
                    foreach (FileData fileData in handler.Files)
                    {
                        foreach (StringValue strVal in fileData.GetStringValueList)
                        {
                            foreach (Error error in strVal.Errors)
                            {
                                dataGridView1.Rows.Add(strongName, fileData.FileName, strVal.Key, strVal.Value, error.Value, false);
                            }
                        }
                    }

                }
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            openToolStripMenuItem_Click(this, new EventArgs());
        }
    }
}