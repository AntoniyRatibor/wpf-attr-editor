using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml;

namespace IntDataParser
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        
        private void openBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog browse = new OpenFileDialog();
            browse.DefaultExt = ".xml";
            browse.Filter = "XML Files (*.xml)|*.xml";
            Nullable<bool> result = browse.ShowDialog();

            if (result == true)
            {
                inFilePath.Text = browse.FileName;
            }
        }

        private void saveBtn_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.Filter = "XML Files (*.xml)|*.xml" + "|Text Files (*.txt)|*.txt";
            Nullable<bool> result = save.ShowDialog();

            if (result == true)
            {
                outFilePath.Text = save.FileName;
            }
        }

        private void createIntRelaysBtn_Click(object sender, RoutedEventArgs e)
        {
            int number = Convert.ToInt32(objNumber.Text);
            XmlDocument doc = new XmlDocument();
            doc.Load(inFilePath.Text);

            StreamWriter file = new StreamWriter(outFilePath.Text, true);
            file.AutoFlush = true;

            XmlNodeList inTag = doc.GetElementsByTagName("in");
            for (int i = 0; i < inTag.Count; i++)
            {
                string attr = inTag[i].Attributes["type"].Value;
                if ((attr.Substring(0, 1) == "И") && ((attr.Substring(attr.Length - 1) == "У") || (attr.Substring(attr.Length - 1) == "1") || (attr.Substring(attr.Length - 1) == "2")))
                {
                    string inTagName = inTag[i].Attributes["type"].Value;
                    string objTagName = inTag[i].ParentNode.Attributes["name"].Value;
                    string intRelay = inTag[i].Attributes["name"].Value;

                    file.WriteLine("<obj objtype=\"Интерфейсное_реле\" name=\"" + objTagName + "_" + inTagName + "\" objnum=\"" + number + "\" subtype=\"1\">");
                    file.WriteLine("\t<attributes jsonValue=\"{}\"/>");
                    file.WriteLine("\t<in name=\"" + intRelay + "\" type=\"И\"/>");
                    file.WriteLine("</obj>");
                    file.WriteLine();
                    number++;
                }
            }
            MessageBox.Show("File \"" + outFilePath.Text + "\" was created.", "Info");
        }

        private void reNumbObjBtn_Click(object sender, RoutedEventArgs e)
        {
            int number = Convert.ToInt32(objNumber.Text);
            XmlDocument doc = new XmlDocument();
            doc.PreserveWhitespace = true;
            doc.Load(inFilePath.Text);

            string objFrom = objNameFrom.Text;
            string objTo = objNameTo.Text;

            XmlNodeList nodeList = doc.GetElementsByTagName("obj");
            for (int j = 0; j < nodeList.Count; j++)
            {
                if (nodeList[j].Attributes != null)
                {
                    if (nodeList[j].Attributes["name"].Value == objFrom)
                    {
                        while (nodeList[j].Attributes["name"].Value != objTo)
                        {
                            nodeList[j].Attributes["objnum"].Value = number.ToString();
                            j++;
                            number++;
                        }
                        nodeList[j].Attributes["objnum"].Value = number.ToString();
                        break;
                    }
                }
            }
            doc.Save(outFilePath.Text);
            MessageBox.Show("File \"" + outFilePath.Text + "\" was created.", "Info");
        }
    }
}
