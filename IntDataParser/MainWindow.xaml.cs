using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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
using Excel = Microsoft.Office.Interop.Excel;

namespace IntDataParser
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Dictionary<string, string> Dict = new Dictionary<string, string>();
        Dictionary<string, string> ModuleDict = new Dictionary<string, string>();
        Dictionary<string, string> Inputs = new Dictionary<string, string>();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void openBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog browse = new OpenFileDialog();
            browse.DefaultExt = ".xml";
            browse.Filter = "All Files (*.xml, *.xls, *.xlsx)|*.xml; *.xls; *.xlsx";
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
                    string intRelay = inTag[i].Attributes["name"].Value;
                    if (intRelay != "")
                    {
                        string inTagName = inTag[i].Attributes["type"].Value;
                        string objTagName = inTag[i].ParentNode.Attributes["name"].Value;

                        file.WriteLine("<obj objtype=\"Интерфейсное_реле\" name=\"" + objTagName + "_" + inTagName + "\" objnum=\"" + number + "\" subtype=\"1\">");
                        file.WriteLine("\t<attributes jsonValue=\"{}\"/>");
                        file.WriteLine("\t<in name=\"" + intRelay + "\" type=\"И\"/>");
                        file.WriteLine("</obj>");
                        number++;
                    }
                }
            }
            file.Close();
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

        private void button_Click(object sender, RoutedEventArgs e)
        {
            int i = 3;
            Excel.Application xlApp = new Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!");
                return;
            }

            Excel.Workbook workBook;
            Excel.Worksheet workSheet;
            object misValue = System.Reflection.Missing.Value;

            workBook = xlApp.Workbooks.Add(misValue);
            workSheet = workBook.Worksheets.get_Item(1);

            XmlDocument doc = new XmlDocument();
            doc.Load(inFilePath.Text);
            XmlNodeList nodeList = doc.GetElementsByTagName("obj");

            foreach (XmlNode node in nodeList)
            {
                int j = 3;
                workSheet.Cells[i, j] = node.Attributes["objtype"].Value;
                j++;
                workSheet.Cells[i, j] = node.Attributes["name"].Value;
                j++;
                workSheet.Cells[i, j] = node.Attributes["subtype"].Value;
                j++;
                workSheet.Cells[i, j] = node.Attributes["objnum"].Value;
                i++;
            }
            workBook.SaveAs("d:\\csharp-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            workBook.Close(true, misValue, misValue);
            xlApp.Quit();
            MessageBox.Show("File was created.", "Info");
        }

        private void editOutName_Click(object sender, RoutedEventArgs e)
        {
            int number = Convert.ToInt32(objNumber.Text);
            XmlDocument doc = new XmlDocument();
            doc.PreserveWhitespace = true;
            doc.Load(inFilePath.Text);

            XmlNodeList nodeList = doc.GetElementsByTagName("out");
            int val = nodeList.Count;
            foreach (XmlNode node in nodeList)
            {
                node.Attributes["name"].Value += "-У";
            }
            doc.Save(outFilePath.Text);
            MessageBox.Show("File \"" + outFilePath.Text + "\" was created.", "Info");
        }

        private void file_compare_Click(object sender, RoutedEventArgs e)
        {
            CompareWindow compareWin = new CompareWindow();
            compareWin.Show();
        }

        private void menuItemExit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void createOCD_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application xlApp = new Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!");
                return;
            }

            Excel.Workbook workBook;
            Excel.Worksheet workSheet;
            Excel.Range range;
            object misValue = System.Reflection.Missing.Value;

            workBook = xlApp.Workbooks.Open(inFilePath.Text);
            workSheet = workBook.Worksheets.get_Item(1);

            range = workSheet.UsedRange;
            int row = range.Rows.Count;
            int column = range.Columns.Count;

            string signalType;

            XmlTextWriter xmlWriter = new XmlTextWriter(outFilePath.Text, Encoding.UTF8);
            xmlWriter.WriteStartElement("checksum");
            xmlWriter.WriteEndElement();
            xmlWriter.Flush();
            xmlWriter.Close();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(outFilePath.Text);

            XmlNode config = xmlDoc.CreateElement("config");
            XmlNode line = xmlDoc.CreateElement("line");
            XmlNode output = xmlDoc.CreateElement("output");
            XmlNode input = xmlDoc.CreateElement("input");
            XmlNode modbusRegisters = xmlDoc.CreateElement("ModbusRegisters");

            XmlAttribute attrLeft = xmlDoc.CreateAttribute("left");
            XmlAttribute attrRight = xmlDoc.CreateAttribute("right");
            XmlAttribute attrTypeLine = xmlDoc.CreateAttribute("type");

            xmlDoc.DocumentElement.AppendChild(config);
            config.AppendChild(line);
            line.Attributes.Append(attrLeft);
            line.Attributes.Append(attrRight);
            line.Attributes.Append(attrTypeLine);
            xmlDoc.DocumentElement.AppendChild(output);
            xmlDoc.DocumentElement.AppendChild(input);
            xmlDoc.DocumentElement.AppendChild(modbusRegisters);

            xmlDoc.Save(outFilePath.Text);

            for (int i = 1; i <= row; i++)
            {
                if (range.Cells[i, 1].Value == null)
                {
                    int cellCounter = 0;
                    for (int k = 1; k <= column; k++)
                    {
                        if (range.Cells[i, k].Value == null)
                        {
                            cellCounter++;
                        }
                    }
                    MessageBox.Show("Warning! " + cellCounter + " cell(s) in " + i + " row is empty. Program aborted.", "Warning");
                    break;
                }

                XmlAttribute attrName = xmlDoc.CreateAttribute("name");
                XmlAttribute attrAddr = xmlDoc.CreateAttribute("addr");
                XmlAttribute attrCcmi_id = xmlDoc.CreateAttribute("ccmi_id");
                XmlAttribute attrType = xmlDoc.CreateAttribute("type");

                signalType = range.Cells[i, 2].Value;
                attrType.Value = range.Cells[i, 3].Value;
                attrCcmi_id.Value = range.Cells[i, 5].Value.ToString();
                attrAddr.Value = range.Cells[i, 6].Value.ToString();

                if (!Dict.ContainsKey(attrCcmi_id.Value))
                {
                    Dict[attrCcmi_id.Value] = range.Cells[i, 4].Value.ToString();
                }

                if (signalType == "ТС")
                {
                    attrName.Value = range.Cells[i, 1].Value;
                    XmlNode newInput = xmlDoc.CreateElement("in");
                    input.AppendChild(newInput);
                    newInput.Attributes.Append(attrType);
                    newInput.Attributes.Append(attrCcmi_id);
                    newInput.Attributes.Append(attrAddr);
                    newInput.Attributes.Append(attrName);
                }
                else if (signalType == "ТУ")
                {
                    attrName.Value = range.Cells[i, 1].Value + "-У";
                    XmlNode newOut = xmlDoc.CreateElement("out");
                    output.AppendChild(newOut);
                    newOut.Attributes.Append(attrCcmi_id);
                    newOut.Attributes.Append(attrAddr);
                    newOut.Attributes.Append(attrName);
                }
                else if (signalType == "modbus")
                {
                    XmlAttribute attrController_id = xmlDoc.CreateAttribute("controllerId");
                    XmlAttribute attrAddress = xmlDoc.CreateAttribute("address");
                    attrController_id.Value = range.Cells[i, 5].Value.ToString();
                    attrAddress.Value = range.Cells[i, 6].Value.ToString();

                    attrName.Value = range.Cells[i, 1].Value;
                    XmlNode newRegister = xmlDoc.CreateElement("register");
                    modbusRegisters.AppendChild(newRegister);
                    newRegister.Attributes.Append(attrType);
                    newRegister.Attributes.Append(attrController_id);
                    newRegister.Attributes.Append(attrAddress);
                    newRegister.Attributes.Append(attrName);
                }

                xmlDoc.Save(outFilePath.Text);
            }

            if (!output.HasChildNodes)
            {
                xmlDoc.DocumentElement.RemoveChild(output);
            }

            HeaderCreator(xmlDoc, line);

            XmlWriterSettings settings = new XmlWriterSettings
            {
                Indent = true,
                IndentChars = "\t",
                OmitXmlDeclaration = true
            };

            using (XmlWriter writer = XmlWriter.Create(outFilePath.Text, settings))
            {
                xmlDoc.Save(writer);
            }

            workBook.Close(false, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(workBook);
            Marshal.ReleaseComObject(workSheet);
            Marshal.ReleaseComObject(xlApp);

            Dict.Clear();

            MessageBox.Show("Done!");
        }

        private void HeaderCreator(XmlDocument xmlDoc, XmlNode line)
        {
            ICollection<string> keys = Dict.Keys;

            foreach (string i in keys)
            {
                XmlAttribute attrType = xmlDoc.CreateAttribute("type");
                XmlAttribute attrId = xmlDoc.CreateAttribute("id");
                XmlAttribute attrAdamR = xmlDoc.CreateAttribute("adamR");
                XmlAttribute attrAdamL = xmlDoc.CreateAttribute("adamL");
                XmlAttribute attrIRegs = xmlDoc.CreateAttribute("iRegs");
                XmlAttribute attrORegs = xmlDoc.CreateAttribute("oRegs");
                XmlAttribute attrCoils = xmlDoc.CreateAttribute("coils");
                XmlAttribute attrInputs = xmlDoc.CreateAttribute("inputs");

                attrId.Value = i;
                attrType.Value = Dict[i];

                XmlNode ccmi = xmlDoc.CreateElement("ccmi");
                line.AppendChild(ccmi);
                ccmi.Attributes.Append(attrType);
                ccmi.Attributes.Append(attrId);
                ccmi.Attributes.Append(attrAdamL);
                ccmi.Attributes.Append(attrAdamR);
                ccmi.Attributes.Append(attrIRegs);
                ccmi.Attributes.Append(attrORegs);
                ccmi.Attributes.Append(attrCoils);
                ccmi.Attributes.Append(attrInputs);

                xmlDoc.Save(outFilePath.Text);
            }
        }

        private void createDiagBtn_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application xlApp = new Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!");
                return;
            }

            Excel.Workbook workBook;
            Excel.Worksheet workSheet;
            Excel.Range range;
            object misValue = System.Reflection.Missing.Value;

            workBook = xlApp.Workbooks.Open(inFilePath.Text);
            workSheet = workBook.Worksheets.get_Item(1);

            range = workSheet.UsedRange;
            int row = range.Rows.Count;

            XmlTextWriter xmlWriter = new XmlTextWriter(outFilePath.Text, Encoding.UTF8);
            xmlWriter.WriteStartElement("checksum");
            xmlWriter.WriteEndElement();
            xmlWriter.Flush();
            xmlWriter.Close();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(outFilePath.Text);

            for (int i = 1; i <= row; i++)
            {
                if (range.Cells[i, 1].Value == null)
                {
                    MessageBox.Show("Warning! First cell in " + i + " row is empty.", "Warning");
                }

                if ((range.Cells[i, 7].Value != null) && (!ModuleDict.ContainsKey(range.Cells[i, 7].Value)))
                {
                    if (range.Cells[i, 2].Value == "modbus")
                    {
                        if (range.Cells[i, 8].Value == "diagnostic")
                        {
                            ModuleDict[range.Cells[i, 7].Value] = range.Cells[i, 3].Value;
                        }
                    }
                    else
                    {
                        ModuleDict[range.Cells[i, 7].Value] = range.Cells[i, 3].Value;
                    }
                }
            }

            for (int j = 1; j <= row; j++)
            {
                if (range.Cells[j, 7].Value != null && range.Cells[j, 3].Value != null)
                {
                    if (range.Cells[j, 2].Value == "modbus")
                    {
                        if (range.Cells[j, 8].Value == "diagnostic")
                        {
                            Inputs[range.Cells[j, 1].Value] = range.Cells[j, 7].Value;
                        }
                    }
                    else
                    {
                        Inputs[range.Cells[j, 1].Value] = range.Cells[j, 7].Value;
                    }
                }
            }

            ICollection<string> ModuleKeys = ModuleDict.Keys;
            foreach (string moduleName in ModuleKeys)
            {
                XmlNode obj = ObjCreator(xmlDoc, moduleName);

                int k = 0;

                ICollection<string> InputKeys = Inputs.Keys;
                foreach (string inputName in InputKeys)
                {
                    if (Inputs[inputName] == moduleName)
                    {
                        XmlNode newInput = xmlDoc.CreateElement("in");

                        XmlAttribute attrName = xmlDoc.CreateAttribute("name");
                        XmlAttribute attrType = xmlDoc.CreateAttribute("type");

                        if (ModuleDict[moduleName] == "input")
                        {
                            attrType.Value = "СОСТ";
                        }
                        else
                        {
                            attrType.Value = "inp" + k.ToString();
                        }
                        
                        attrName.Value = inputName;

                        obj.AppendChild(newInput);
                        newInput.Attributes.Append(attrType);
                        newInput.Attributes.Append(attrName);
                        k++;
                        xmlDoc.Save(outFilePath.Text);
                    }
                }
                xmlDoc.Save(outFilePath.Text);
            }

            XmlWriterSettings settings = new XmlWriterSettings
            {
                Indent = true,
                IndentChars = "\t",
                OmitXmlDeclaration = true
            };

            using (XmlWriter writer = XmlWriter.Create(outFilePath.Text, settings))
            {
                xmlDoc.Save(writer);
            }

            workBook.Close(false, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(workBook);
            Marshal.ReleaseComObject(workSheet);
            Marshal.ReleaseComObject(xlApp);

            ModuleDict.Clear();
            Inputs.Clear();

            MessageBox.Show("Done!");
        }

        private XmlNode ObjCreator(XmlDocument doc, string moduleName)
        {
            XmlNode obj = doc.CreateElement("obj");
            doc.DocumentElement.AppendChild(obj);

            XmlAttribute attrObjType = doc.CreateAttribute("objtype");
            XmlAttribute attrSubType = doc.CreateAttribute("subtype");
            XmlAttribute attrObjName = doc.CreateAttribute("name");
            XmlAttribute attrObjNum = doc.CreateAttribute("objnum");

            if (ModuleDict[moduleName] == "Т" || ModuleDict[moduleName] == "М" || ModuleDict[moduleName] == "Д" || ModuleDict[moduleName] == "input")
            {
                attrObjType.Value = "Модуль_ввода";
            }
            else
            {
                attrObjType.Value = "Модуль_вывода";
            }

            if (ModuleDict[moduleName] == "input")
            {
                attrSubType.Value = "3";
            }
            else if (ModuleDict[moduleName] == "М" || ModuleDict[moduleName] == "Д")
            {
                attrSubType.Value = "2";
            }
            else
            {
                attrSubType.Value = "1";
            }

            if (ModuleDict[moduleName] == "input")
            {
                attrObjName.Value = "Модуль_" + moduleName;
            }
            else
            {
                attrObjName.Value = moduleName;
            }

            doc.DocumentElement.AppendChild(obj);
            obj.Attributes.Append(attrObjType);
            obj.Attributes.Append(attrSubType);
            obj.Attributes.Append(attrObjName);
            obj.Attributes.Append(attrObjNum);

            XmlNode attributes = doc.CreateElement("attributes");
            XmlAttribute attrJsonValue = doc.CreateAttribute("jsonValue");

            attrJsonValue.Value = "{}";

            obj.AppendChild(attributes);
            attributes.Attributes.Append(attrJsonValue);

            doc.Save(outFilePath.Text);
            return obj;
        }
    }
}