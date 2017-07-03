using Microsoft.Win32;
using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;
using System.Xml;

namespace IntDataParser
{
    public partial class CompareWindow : Window
    {
        Dictionary<string, string> OCdataDict = new Dictionary<string, string>();

        public CompareWindow()
        {
            InitializeComponent();
        }

        private void openOCData_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog browse = new OpenFileDialog();
            browse.DefaultExt = ".xml";
            browse.Filter = "XML Files (*.xml)|*.xml";
            Nullable<bool> result = browse.ShowDialog();

            if (result == true)
            {
                OCDataPath.Text = browse.FileName;
            }
        }

        private void openIntData_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog browse = new OpenFileDialog();
            browse.DefaultExt = ".xml";
            browse.Filter = "XML Files (*.xml)|*.xml";
            Nullable<bool> result = browse.ShowDialog();

            if (result == true)
            {
                IntDataPath.Text = browse.FileName;
            }
        }

        private void compareBtn_Click(object sender, RoutedEventArgs e)
        {
            XmlDocument docOCData = new XmlDocument();
            docOCData.PreserveWhitespace = true;
            docOCData.Load(OCDataPath.Text);

            XmlNodeList outsOCData = docOCData.GetElementsByTagName("out");
            for (int j = 0; j < outsOCData.Count; j++)
            {
                try
                {
                    OCdataDict.Add(outsOCData[j].Attributes["name"].Value, null);
                }
                catch (ArgumentException)
                {
                    textBox.Text += "Warning! Duplicated input " + outsOCData[j].Attributes["name"].Value + " in OCData.xml\n";
                }
            }

            XmlNodeList inputsOCData = docOCData.GetElementsByTagName("in");
            for (int j = 0; j < inputsOCData.Count; j++)
            {
                try
                {
                    OCdataDict.Add(inputsOCData[j].Attributes["name"].Value, null);
                }
                catch (ArgumentException)
                {
                    textBox.Text += "Warning! Duplicated input " + inputsOCData[j].Attributes["name"].Value + " in OCData.xml\n";
                }
            }

            XmlNodeList registersCData = docOCData.GetElementsByTagName("register");
            for (int j = 0; j < registersCData.Count; j++)
            {
                if (registersCData[j].Attributes["type"].Value == "input")
                {
                    try
                    {
                        OCdataDict.Add(registersCData[j].Attributes["name"].Value, null);
                    }
                    catch (ArgumentException)
                    {
                        textBox.Text += "Warning! Duplicated register " + registersCData[j].Attributes["name"].Value + " in OCData.xml\n";
                    }
                }
            }

            XmlDocument docIntData = new XmlDocument();
            docIntData.PreserveWhitespace = true;
            docIntData.Load(IntDataPath.Text);

            XmlNodeList outsIntData = docIntData.GetElementsByTagName("out");
            for (int j = 0; j < outsIntData.Count; j++)
            {
                string name = outsIntData[j].Attributes["name"].Value;
                if (OCdataDict.ContainsKey(name))
                {
                    OCdataDict[name] += "(" + outsIntData[j].ParentNode.Attributes["objtype"].Value + " " + outsIntData[j].ParentNode.Attributes["name"].Value + ") ";
                }
                else
                {
                    textBox.Text += "Warning! There is no " + name + " in OCData.xml\n";
                }
            }

            XmlNodeList inputsIntData = docIntData.GetElementsByTagName("in");
            for (int j = 0; j < inputsIntData.Count; j++)
            {
                string name = inputsIntData[j].Attributes["name"].Value;
                if (OCdataDict.ContainsKey(name))
                {
                    OCdataDict[name] += "(" + inputsIntData[j].ParentNode.Attributes["objtype"].Value + " " + inputsIntData[j].ParentNode.Attributes["name"].Value + ") ";
                }
                else
                {
                    textBox.Text += "Warning! There is no " + name + " in OCData.xml\n";
                }
            }

            ICollection<string> keys = OCdataDict.Keys;
            
            foreach (string i in keys)
            {
                textBox.Text += i + " = " + OCdataDict[i] + "\n";
            }
            OCdataDict.Clear();
        }

        private void clearBtn_Click(object sender, RoutedEventArgs e)
        {
            textBox.Clear();
        }
    }
}
