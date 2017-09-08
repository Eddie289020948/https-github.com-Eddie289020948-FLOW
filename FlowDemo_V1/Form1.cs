using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FlowDemo_BL;
using System.Xml;
using System.Xml.Linq;
using System.Reflection;
using System.Collections;

namespace FlowDemo_V1
{
    public partial class Form1 : Form
    {
        List<TestFlow> oListTestFlow;
        List<TestFlow> cListTestFlow;
        Dictionary<int, List<FunctionParameter>> cParamsByID = new Dictionary<int, List<FunctionParameter>>();
        
        Dictionary<string, List<FunctionParameter>> dictFuncs;
        List<int> dictID = new List<int>();
        
        ComboBox cboFuncs = new ComboBox();
        string xmlPath = string.Empty;
        Assembly assAPP;
        Type tApp;
        int numberToContinue = 0;

        public Form1()
        {
            InitializeComponent();

            InitializeComponentValue();
        }

        public Form1(string sXMLPath)
        {
            InitializeComponent();

            InitializeComponentValue();

            xmlPath = sXMLPath;
            LoadData();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cboFuncs.Visible = false;

            cboFuncs.SelectedIndexChanged += new EventHandler(cboFuncs_SelectedIndexChanged);

            dgvTestFlow.Controls.Add(cboFuncs);
        }

        private void InitializeComponentValue()
        {
            textBox1.Text = Definition.PLEASE_IDENTIFY_OUTPUT_FILE_PATH;

            cboMode.Text = "Edit Mode";
        }

        private void fileToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Filter = "FLOW File|*.flw";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    xmlPath = dialog.FileName;
                    LoadData();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dLLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Todo: if old dll path is not equal to new dll path, clear all grids and data
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "DLLFile|*.dll";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                txtDLLPath.Text = dialog.FileName;
                LoadDictFuncs(txtDLLPath.Text);
            }
        }

        private void btnUp_Click(object sender, EventArgs e)
        {
            if (dgvTestFlow.SelectedRows.Count <= 0)
            {
                MessageBox.Show("No row is selected!");
                return;
            }

            if (!CheckContinue())
            {
                MessageBox.Show("The selected rows are not continuous!");
                return;
            }

            List<DataGridViewRow> sortedSelectedRows = dgvTestFlow.SelectedRows.Sort();
            int firstSelectedRowIndex = sortedSelectedRows[0].Index;
            if (firstSelectedRowIndex == 0)
            {
                MessageBox.Show("Data has already been the top one(s)!");
                return;
            }

            DataGridViewRow row = dgvTestFlow.Rows[firstSelectedRowIndex - 1];
            dgvTestFlow.Rows.RemoveAt(firstSelectedRowIndex - 1);

            int lastSelectedRowIndex = sortedSelectedRows[sortedSelectedRows.Count - 1].Index;
            dgvTestFlow.Rows.Insert(lastSelectedRowIndex + 1, row);
        }

        private void btnDown_Click(object sender, EventArgs e)
        {
            if (dgvTestFlow.SelectedRows.Count <= 0)
            {
                MessageBox.Show("No row is selected!");
                return;
            }

            if (!CheckContinue())
            {
                MessageBox.Show("The selected rows are not continuous!");
                return;
            }

            List<DataGridViewRow> sortedSelectedRows = dgvTestFlow.SelectedRows.Sort();
            int lastSelectedRowIndex = sortedSelectedRows[sortedSelectedRows.Count - 1].Index;
            if (lastSelectedRowIndex == dgvTestFlow.Rows[dgvTestFlow.Rows.Count - 1].Index)
            {
                MessageBox.Show("Data has already been the bottom one(s)!");
                return;
            }

            DataGridViewRow row = dgvTestFlow.Rows[lastSelectedRowIndex + 1];
            dgvTestFlow.Rows.RemoveAt(lastSelectedRowIndex + 1);

            int firstSelectedRowIndex = sortedSelectedRows[0].Index;
            dgvTestFlow.Rows.Insert(firstSelectedRowIndex, row);
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            int index = dgvTestFlow.Rows.Add();

            //Todo: need a method to generate ID
            dgvTestFlow.Rows[index].Cells["ID"].Value = index + 1;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (ValidateTestFlowConfig())
            {
                //
                if (!string.IsNullOrEmpty(xmlPath))
                    BackupFile();
                else
                {
                    //Todo: if xmlPath is not equal to dllPath
                    FolderBrowserDialog dialog = new FolderBrowserDialog();
                    //dialog.RootFolder = Environment.SpecialFolder.MyComputer;
                    dialog.ShowDialog();
                    if (!string.IsNullOrEmpty(dialog.SelectedPath))
                    {
                        string fileName = txtDLLPath.Text.Substring(txtDLLPath.Text.LastIndexOf("\\"));
                        fileName = fileName.Substring(0, fileName.LastIndexOf("."));
                        xmlPath = dialog.SelectedPath + fileName + ".flw";
                        GenerateXmlFile(xmlPath);
                    }
                    else
                    {
                        return;
                    }
                }

                //
                CacheData();
                //
                SaveData();
            }
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            dgvTestFlow.Rows.Clear();
            cParamsByID.Clear();

            if (oListTestFlow != null)
            {
                foreach (TestFlow testFlow in oListTestFlow)
                {
                    int index = dgvTestFlow.Rows.Add();
                    dgvTestFlow.Rows[index].Cells["ID"].Value = testFlow.ID;
                    cParamsByID.Add(testFlow.ID, new List<FunctionParameter>());
                    dgvTestFlow.Rows[index].Cells["TestNumber"].Value = testFlow.TestNumber;
                    dgvTestFlow.Rows[index].Cells["TestName"].Value = testFlow.TestName;
                    dgvTestFlow.Rows[index].Cells["TestFunction"].Value = testFlow.TestFunction;
                    dgvTestFlow.Rows[index].Cells["UpperLimit"].Value = testFlow.UpperLimit;
                    dgvTestFlow.Rows[index].Cells["LowerLimit"].Value = testFlow.LowerLimit;
                    dgvTestFlow.Rows[index].Cells["Unit"].Value = testFlow.Unit;
                    dgvTestFlow.Rows[index].Cells["SoftBin"].Value = testFlow.SoftBin;
                    dgvTestFlow.Rows[index].Cells["HardBin"].Value = testFlow.HardBin;
                    dgvTestFlow.Rows[index].Cells["Action"].Value = testFlow.Action;

                    foreach(var fp in testFlow.TestFunctionParameters)
                    {
                        cParamsByID[testFlow.ID].Add(fp);
                    }
                }
            }

            //dgvTestFlow.ClearSelection();
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            dgvTestFlow.Rows.Clear();
            cParamsByID.Clear();
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            for (int i = dgvTestFlow.SelectedRows.Count; i > 0; i--)
            {
                dgvTestFlow.Rows.RemoveAt(dgvTestFlow.SelectedRows[i - 1].Index);
            }

            //if (MessageBox.Show("Delete the selected row(s)? ", "Confirm", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            //{
                
            //}
        }

        private void btnClone_Click(object sender, EventArgs e)
        {
            List<DataGridViewRow> sortedSelectedRows = dgvTestFlow.SelectedRows.Sort();

            int index = sortedSelectedRows[sortedSelectedRows.Count - 1].Index;

            for (int i = 0; i < sortedSelectedRows.Count; i++)
            {
                dgvTestFlow.Rows.Insert(index + 1 + i, sortedSelectedRows[i].CloneWithValues());
            }
        }

        private void cboFuncs_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataGridViewCell CurrentCell = dgvTestFlow.CurrentCell;
            if (CurrentCell != null && CurrentCell.OwningColumn.Name == "TestFunction")
            {
                CurrentCell.Value = ((ComboBox)sender).Text;
                CurrentCell.Tag = ((ComboBox)sender).Text;

                LoadFunctionParameter(Convert.ToInt32(CurrentCell.OwningRow.Cells["ID"].Value), ((ComboBox)sender).Text);
            }
        }

        private void dgvTestFlow_CurrentCellChanged(object sender, EventArgs e)
        {
            try
            {
                DataGridViewCell CurrentCell = dgvTestFlow.CurrentCell;
                if (CurrentCell != null && CurrentCell.OwningColumn.Name == "TestFunction")
                {
                    Rectangle rect = dgvTestFlow.GetCellDisplayRectangle(CurrentCell.ColumnIndex, CurrentCell.RowIndex, false);
                    cboFuncs.Text = (CurrentCell.Value == null) ? string.Empty : CurrentCell.Value.ToString();
                    cboFuncs.Size = rect.Size;
                    cboFuncs.Top = rect.Top;
                    cboFuncs.Left = rect.Left;
                    cboFuncs.Width = rect.Width;
                    cboFuncs.Height = rect.Height;
                    cboFuncs.Visible = true;
                }
                else
                {
                    cboFuncs.Visible = false;
                }
            }
            catch
            {

            }
        }

        private void dgvTestFlow_Scroll(object sender, ScrollEventArgs e)
        {
            cboFuncs.Visible = false;
        }

        private void dgvTestFlow_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            cboFuncs.Visible = false;
        }

        private void dgvFunctionParameter_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewCell CurrentCell = dgvFunctionParameter.CurrentCell;
            if (CurrentCell != null && CurrentCell.OwningColumn.Name == "ParameterValue" && CurrentCell.OwningRow.Index > 0)
            {
                CurrentCell.Style.BackColor = Color.Empty;
                CurrentCell.Style.SelectionBackColor = Color.Empty;

                int id = Convert.ToInt32(dgvTestFlow.CurrentRow.Cells["ID"].Value);
                string sParameterName = CurrentCell.OwningRow.Cells["ParameterName"].Value.ToString();
                string sParameterType = CurrentCell.OwningRow.Cells["ParameterType"].Value.ToString();
                string sParameterValue = string.Empty;
                if (CurrentCell.Value == null)
                {
                    PopupError(dgvFunctionParameter, "ParameterValue", CurrentCell.OwningRow.Index, "Parameter Value");
                    return;
                }
                else
                {
                    sParameterValue = CurrentCell.Value.ToString();
                }
                

                if (!cParamsByID.Any(x => x.Key == id))
                {
                    cParamsByID.Add(id, new List<FunctionParameter>() { new FunctionParameter() { ParameterName = sParameterName, ParameterType = sParameterType, ParameterValue = sParameterValue } });
                }
                else
                {
                    if (!cParamsByID[id].Any(x => x.ParameterName == sParameterName))
                    {
                        cParamsByID[id].Add(new FunctionParameter() { ParameterName = sParameterName, ParameterType = sParameterType, ParameterValue = sParameterValue });
                    }
                    else
                    {
                        cParamsByID[id].FirstOrDefault(x => x.ParameterName == sParameterName).ParameterValue = sParameterValue;
                    }
                }
            }
        }

        private void dgvTestFlow_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvTestFlow.SelectedRows.Count == 1)
            {
                var row = dgvTestFlow.SelectedRows[0];
                if (row.Cells["ID"].Value != null && row.Cells["TestFunction"].Value != null)
                    LoadFunctionParameter(Convert.ToInt32(row.Cells["ID"].Value), row.Cells["TestFunction"].Value.ToString());
            }
        }

        private void cboMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            setButtonVisible(cboMode.Text == "Edit Mode");
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (tApp != null)
            {
                object obj = Activator.CreateInstance(tApp);
                
                MethodInfo mi = tApp.GetMethod("Start");

                object[] parameters = new object[0];

                int res = (int)mi.Invoke(obj, parameters);
            }
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            if (tApp != null)
            {
                object obj = Activator.CreateInstance(tApp);

                MethodInfo mi = tApp.GetMethod("Stop");

                object[] parameters = new object[0];

                int res = (int)mi.Invoke(obj, parameters);
            }
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            if (backgroundWorker1.IsBusy != true)
            {
                //Need to disable other input fields.
                btnRun.Enabled = false;
                btnHalt.Enabled = true;

                foreach(DataGridViewRow row in dgvTestFlow.Rows)
                {
                    row.DefaultCellStyle.BackColor = Color.Empty;
                }

                // Start the asynchronous operation. 
                backgroundWorker1.RunWorkerAsync(numberToContinue);
            }
        }

        private void btnHalt_Click(object sender, EventArgs e)
        {
            if (backgroundWorker1.WorkerSupportsCancellation == true)
            {
                // Cancel the asynchronous operation. 
                backgroundWorker1.CancelAsync();

                btnHalt.Enabled = false;
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;

            if (tApp != null)
            {
                object obj = Activator.CreateInstance(tApp);

                for (int i = (int)e.Argument; i < cListTestFlow.Count; i++)
                {
                    if (worker.CancellationPending == true)
                    {
                        e.Cancel = true;
                        break;
                    }
                    else
                    {
                        var test = cListTestFlow[i];
                        MethodInfo mi = tApp.GetMethod(test.TestFunction);

                        object[] parameters = new object[test.TestFunctionParameters.Count];

                        for (int j = 0; j < test.TestFunctionParameters.Count; j++)
                        {
                            FunctionParameter fp = test.TestFunctionParameters[j];

                            parameters[j] = ValueTranslator(fp.ParameterType, fp.ParameterValue);
                        }

                        mi.Invoke(obj, parameters);

                        System.Threading.Thread.Sleep(500);

                        worker.ReportProgress((int)((i + 1) / cListTestFlow.Count * 100), i);
                    }
                }
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            int rowIndex = (int)e.UserState;

            if (rowIndex > 0)
                dgvTestFlow.Rows[rowIndex - 1].DefaultCellStyle.BackColor = Color.Empty;

            dgvTestFlow.Rows[rowIndex].DefaultCellStyle.BackColor = Color.Green;
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                MessageBox.Show("Canceled!");
            }
            else if (e.Error != null)
            {
                MessageBox.Show("Error: " + e.Error.Message);
            }
            else
            {
                MessageBox.Show("Done!");
            }

            //Need to enable other input fields.
            btnRun.Enabled = true;
            btnHalt.Enabled = false;
        }

        #region private methods
        private void LoadData()
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlPath);
            //Todo: Validate Schema
            XmlNode xnScequencer = xmlDoc.SelectSingleNode("scequencer");

            LoadDllPath(xnScequencer);

            //LoadSiteConfig
            LoadSiteConfig(xnScequencer);

            //LoadStopContinueConfig
            LoadStopContinueConfig(xnScequencer);

            //LoadTestFlowConfig
            LoadTestFlowConfig(xnScequencer);
        }

        private void LoadDllPath(XmlNode xnScequencer)
        {
            XmlNode xnDllPath = xnScequencer.SelectSingleNode("dllpath");
            txtDLLPath.Text = xnDllPath.InnerText;
            LoadDictFuncs(txtDLLPath.Text);
        }

        private void LoadSiteConfig(XmlNode xnScequencer)
        { }

        private void LoadStopContinueConfig(XmlNode xnScequencer)
        { }

        private void LoadTestFlowConfig(XmlNode xnScequencer)
        {
            dgvTestFlow.Rows.Clear();
            cParamsByID.Clear();
            oListTestFlow = new List<TestFlow>();
            XmlNode xnTestFlow = xnScequencer.SelectSingleNode("testflowconfig");

            foreach (XmlNode xn in xnTestFlow.ChildNodes)
            {
                TestFlow testFlow = new TestFlow();
                testFlow.ID = Convert.ToInt32(xn.SelectSingleNode("id").InnerText);
                cParamsByID.Add(testFlow.ID, new List<FunctionParameter>());
                testFlow.TestNumber = Convert.ToInt32(xn.SelectSingleNode("testnumber").InnerText);
                testFlow.TestName = xn.SelectSingleNode("testname").InnerText;
                testFlow.TestFunction = xn.SelectSingleNode("testfunction").InnerText;
                testFlow.UpperLimit = Convert.ToDouble(xn.SelectSingleNode("upperlimit").InnerText);
                testFlow.LowerLimit = Convert.ToDouble(xn.SelectSingleNode("lowerlimit").InnerText);
                testFlow.Unit = xn.SelectSingleNode("unit").InnerText;
                testFlow.SoftBin = Convert.ToInt32(xn.SelectSingleNode("softbin").InnerText);
                testFlow.HardBin = Convert.ToInt32(xn.SelectSingleNode("hardbin").InnerText);
                testFlow.Action = xn.SelectSingleNode("action").InnerText;

                XmlNode xnTestFunctionParameters = xn.SelectSingleNode("testfunctionparameters");
                foreach (XmlNode x in xnTestFunctionParameters)
                {
                    FunctionParameter fp = new FunctionParameter();
                    fp.ParameterName = x.SelectSingleNode("parametername").InnerText;
                    fp.ParameterType = x.SelectSingleNode("parametertype").InnerText;
                    fp.ParameterValue = x.SelectSingleNode("parametervalue").InnerText;
                    testFlow.TestFunctionParameters.Add(fp);
                    cParamsByID[testFlow.ID].Add(fp);
                }

                oListTestFlow.Add(testFlow);

                int index = dgvTestFlow.Rows.Add();
                dgvTestFlow.Rows[index].Cells["ID"].Value = testFlow.ID;
                dgvTestFlow.Rows[index].Cells["TestNumber"].Value = testFlow.TestNumber;
                dgvTestFlow.Rows[index].Cells["TestName"].Value = testFlow.TestName;
                dgvTestFlow.Rows[index].Cells["TestFunction"].Value = testFlow.TestFunction;
                dgvTestFlow.Rows[index].Cells["UpperLimit"].Value = testFlow.UpperLimit;
                dgvTestFlow.Rows[index].Cells["LowerLimit"].Value = testFlow.LowerLimit;
                dgvTestFlow.Rows[index].Cells["Unit"].Value = testFlow.Unit;
                dgvTestFlow.Rows[index].Cells["SoftBin"].Value = testFlow.SoftBin;
                dgvTestFlow.Rows[index].Cells["HardBin"].Value = testFlow.HardBin;
                dgvTestFlow.Rows[index].Cells["Action"].Value = testFlow.Action;
            }

            dgvTestFlow.ClearSelection();

            //Todo: Delete this later
            cListTestFlow = oListTestFlow;
        }

        private void LoadFunctionParameter(int id, string sTestFunction)
        {
            dgvFunctionParameter.Rows.Clear();

            foreach (var dict in dictFuncs)
            {
                if (dict.Key == sTestFunction)
                {
                    foreach(FunctionParameter fp in dict.Value)
                    {
                        int index = dgvFunctionParameter.Rows.Add();
                        dgvFunctionParameter.Rows[index].Cells["ParameterName"].Value = fp.ParameterName;
                        dgvFunctionParameter.Rows[index].Cells["ParameterType"].Value = fp.ParameterType;

                        if(cParamsByID.Any(x => x.Key == id))
                        {
                            if(cParamsByID[id].Any(x => x.ParameterName == fp.ParameterName))
                            {
                                dgvFunctionParameter.Rows[index].Cells["ParameterValue"].Value = cParamsByID[id].FirstOrDefault(x => x.ParameterName == fp.ParameterName).ParameterValue;
                            }
                        }
                    }
                    break;
                }
            }
        }

        private void LoadDictFuncs(string a)
        {
            if(!System.IO.File.Exists(a))
            {
                throw new Exception("Could not find the dll from " + a);
            }

            assAPP = System.Reflection.Assembly.LoadFile(a);
            string[] aa = a.Split('\\');
            a = aa[aa.Length - 1].ToLower();
            a = a.Substring(0, a.IndexOf(".dll"));
            tApp = assAPP.GetType(assAPP.GetTypes().FirstOrDefault(x => x.Name.ToLower() == a).FullName);

            MethodInfo[] mis = tApp.GetMethods(BindingFlags.Public | BindingFlags.Static);
            dictFuncs = new Dictionary<string, List<FunctionParameter>>();
            foreach (MethodInfo mi in mis)
            {
                List<FunctionParameter> lFuncParams = new List<FunctionParameter>();
                lFuncParams.Add(new FunctionParameter() { ParameterName = "Return Value", ParameterType = TypeTranslator(mi.ReturnType) });

                ParameterInfo[] pis = mi.GetParameters();
                for (int i = 0; i < pis.Length; i++)
                {
                    lFuncParams.Add(new FunctionParameter() { ParameterName = pis[i].Name, ParameterType = TypeTranslator(pis[i].ParameterType) });
                }

                dictFuncs.Add(mi.Name, lFuncParams);
            }

            BindFuncs();
        }

        private void SaveData()
        {
            XElement xeScequencer = XElement.Load(xmlPath);

            SaveDllPath(xeScequencer);

            SaveTestFlowConfig(xeScequencer);

            xeScequencer.Save(xmlPath);
        }

        private void SaveDllPath(XElement xeScequencer)
        {
            XElement xeDllPath = xeScequencer.Element("dllpath");
            xeDllPath.Value = txtDLLPath.Text;
        }

        private void SaveTestFlowConfig(XElement xeScequencer)
        {
            XElement xeTestflowconfig = xeScequencer.Element("testflowconfig");
            xeTestflowconfig.RemoveAll();

            foreach (TestFlow tf in cListTestFlow)
            {
                XElement xeTestFlow = new XElement("testflow",
                                                        new XElement("id", tf.ID.ToString()),
                                                        new XElement("testnumber", tf.TestNumber.ToString()),
                                                        new XElement("testname", tf.TestName.ToString()),
                                                        new XElement("testfunction", tf.TestFunction.ToString()),
                                                        new XElement("upperlimit", tf.UpperLimit.ToString()),
                                                        new XElement("lowerlimit", tf.LowerLimit.ToString()),
                                                        new XElement("unit", tf.Unit.ToString()),
                                                        new XElement("softbin", tf.SoftBin.ToString()),
                                                        new XElement("hardbin", tf.HardBin.ToString()),
                                                        new XElement("action", tf.Action.ToString()));

                XElement xetestfunctionparameters = new XElement("testfunctionparameters");

                foreach (FunctionParameter fp in tf.TestFunctionParameters)
                {
                    XElement xetestfunctionparameter = new XElement("testfunctionparameter",
                                                            new XElement("parametername", fp.ParameterName.ToString()),
                                                            new XElement("parametertype", fp.ParameterType.ToString()),
                                                            new XElement("parametervalue", fp.ParameterValue.ToString()));
                    xetestfunctionparameters.Add(xetestfunctionparameter);
                }

                xeTestFlow.Add(xetestfunctionparameters);

                xeTestflowconfig.Add(xeTestFlow);
            }
        }

        private void CacheData()
        {
            CacheTestFlowConfig();
        }

        private bool ValidateData()
        {
            bool bResult = true;

            return bResult;
        }

        private bool ValidateTestFlowConfig()
        {
            bool bResult = true;

            if (dgvTestFlow.Rows.Count == 0)
            {
                MessageBox.Show("No data to be inserted. Please enter proper data.");
                return false;
            }

            for (int i = 0; i < dgvTestFlow.Rows.Count; i++)
            {
                if (dgvTestFlow.Rows[i].Cells["TestNumber"].Value != null)
                {
                    try
                    {
                        Convert.ToInt32(dgvTestFlow.Rows[i].Cells["TestNumber"].Value);
                    }
                    catch(Exception)
                    {
                        bResult = PopupError(dgvTestFlow, "TestNumber", i, "Test Number");
                        break;
                    }
                }
                else
                {
                    bResult = PopupError(dgvTestFlow, "TestNumber", i, "Test Number");
                    break;
                }

                if (dgvTestFlow.Rows[i].Cells["TestName"].Value != null)
                {
                    if (!(dgvTestFlow.Rows[i].Cells["TestName"].Value is string))
                    {
                        bResult = PopupError(dgvTestFlow, "TestName", i, "Test Name");
                        break;
                    }
                }
                else
                {
                    bResult = PopupError(dgvTestFlow, "TestName", i, "Test Name");
                    break;
                }

                if (dgvTestFlow.Rows[i].Cells["TestFunction"].Value != null)
                {
                    if (!(dgvTestFlow.Rows[i].Cells["TestFunction"].Value is string))
                    {
                        bResult = PopupError(dgvTestFlow, "TestFunction", i, "Test Function");
                        break;
                    }
                }
                else
                {
                    bResult = PopupError(dgvTestFlow, "TestFunction", i, "Test Function");
                    break;
                }

                if (dgvTestFlow.Rows[i].Cells["UpperLimit"].Value != null)
                {
                    try
                    {
                        Convert.ToDouble(dgvTestFlow.Rows[i].Cells["UpperLimit"].Value);
                    }
                    catch (Exception)
                    {
                        bResult = PopupError(dgvTestFlow, "UpperLimit", i, "Upper Limit");
                        break;
                    }
                }
                else
                {
                    bResult = PopupError(dgvTestFlow, "UpperLimit", i, "Upper Limit");
                    break;
                }

                if (dgvTestFlow.Rows[i].Cells["LowerLimit"].Value != null)
                {
                    try
                    {
                        Convert.ToDouble(dgvTestFlow.Rows[i].Cells["LowerLimit"].Value);
                    }
                    catch (Exception)
                    {
                        bResult = PopupError(dgvTestFlow, "LowerLimit", i, "Lower Limit");
                        break;
                    }
                }
                else
                {
                    bResult = PopupError(dgvTestFlow, "LowerLimit", i, "Lower Limit");
                    break;
                }

                if (dgvTestFlow.Rows[i].Cells["Unit"].Value != null)
                {
                    if (!(dgvTestFlow.Rows[i].Cells["Unit"].Value is string))
                    {
                        bResult = PopupError(dgvTestFlow, "Unit", i, "Unit");
                        break;
                    }
                }
                else
                {
                    bResult = PopupError(dgvTestFlow, "Unit", i, "Unit");
                    break;
                }

                if (dgvTestFlow.Rows[i].Cells["SoftBin"].Value != null)
                {
                    try
                    {
                        Convert.ToInt32(dgvTestFlow.Rows[i].Cells["SoftBin"].Value);
                    }
                    catch (Exception)
                    {
                        bResult = PopupError(dgvTestFlow, "SoftBin", i, "Soft Bin");
                        break;
                    }
                }
                else
                {
                    bResult = PopupError(dgvTestFlow, "SoftBin", i, "Soft Bin");
                    break;
                }

                if (dgvTestFlow.Rows[i].Cells["HardBin"].Value != null)
                {
                    try
                    {
                        Convert.ToInt32(dgvTestFlow.Rows[i].Cells["HardBin"].Value);
                    }
                    catch (Exception)
                    {
                        bResult = PopupError(dgvTestFlow, "HardBin", i, "Hard Bin");
                        break;
                    }
                }
                else
                {
                    bResult = PopupError(dgvTestFlow, "HardBin", i, "Hard Bin");
                    break;
                }

                if (dgvTestFlow.Rows[i].Cells["Action"].Value != null)
                {
                    if (!(dgvTestFlow.Rows[i].Cells["Action"].Value is string))
                    {
                        bResult = PopupError(dgvTestFlow, "Action", i, "Action");
                        break;
                    }
                }
                else
                {
                    bResult = PopupError(dgvTestFlow, "Action", i, "Action");
                    break;
                }
            }

            return bResult;
        }

        private bool PopupError(DataGridView dgv, string columnName, int rowIndex, string msg)
        {
            dgv.ClearSelection();

            dgv.Rows[rowIndex].Cells[columnName].Style.BackColor = Color.Red;
            dgv.Rows[rowIndex].Cells[columnName].Style.SelectionBackColor = Color.Red;

            MessageBox.Show("Invalid " + msg);

            return false;
        }

        private void CacheTestFlowConfig()
        {
            cListTestFlow = new List<TestFlow>();

            for (int i = 0; i < dgvTestFlow.Rows.Count; i++)
            {
                TestFlow testFlow = new TestFlow();
                testFlow.ID = Convert.ToInt32(dgvTestFlow.Rows[i].Cells["ID"].Value);
                testFlow.TestNumber = Convert.ToInt32(dgvTestFlow.Rows[i].Cells["TestNumber"].Value);
                testFlow.TestName = Convert.ToString(dgvTestFlow.Rows[i].Cells["TestName"].Value);
                testFlow.TestFunction = Convert.ToString(dgvTestFlow.Rows[i].Cells["TestFunction"].Value);
                testFlow.UpperLimit = Convert.ToDouble(dgvTestFlow.Rows[i].Cells["UpperLimit"].Value);
                testFlow.LowerLimit = Convert.ToDouble(dgvTestFlow.Rows[i].Cells["LowerLimit"].Value);
                testFlow.Unit = Convert.ToString(dgvTestFlow.Rows[i].Cells["Unit"].Value);
                testFlow.SoftBin = Convert.ToInt32(dgvTestFlow.Rows[i].Cells["SoftBin"].Value);
                testFlow.HardBin = Convert.ToInt32(dgvTestFlow.Rows[i].Cells["HardBin"].Value);
                testFlow.Action = Convert.ToString(dgvTestFlow.Rows[i].Cells["Action"].Value);

                if(cParamsByID.Any(x => x.Key == testFlow.ID))
                {
                    testFlow.TestFunctionParameters = cParamsByID[testFlow.ID];
                }

                cListTestFlow.Add(testFlow);
            }
        }

        private void BackupFile()
        {
            //System.IO.Directory.Exists(AppDomain.CurrentDomain.BaseDirectory + @"..\..\Backup")
            //System.IO.Directory.Exists(Application.StartupPath + @"\..\..\Backup")
            //System.IO.Directory.Exists(@"..\..\Backup")
            //Above three are same
            string fldBackup = xmlPath.Substring(0, xmlPath.LastIndexOf("\\") + 1) + "Backup";
            if (!System.IO.Directory.Exists(fldBackup))
            {
                System.IO.Directory.CreateDirectory(fldBackup);
            }

            string fileName = xmlPath.Substring(xmlPath.LastIndexOf("\\"));
            string targetPath = fldBackup + fileName;
            System.IO.File.Copy(xmlPath, targetPath, true);
        }

        private static void GenerateXmlFile(string xmlPath)
        {
            try
            {
                //定义一个XDocument结构
                XDocument myXDoc = new XDocument(
                                       new XElement("scequencer",
                                           new XElement("dllpath"),
                                           new XElement("siteconfig"),
                                           new XElement("stopcontinueconfig"),
                                           new XElement("testflowconfig")
                                                 ));

                //保存此结构（即：我们预期的xml文件）
                myXDoc.Save(xmlPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private bool CheckContinue()
        {
            for (int i = 0; i < dgvTestFlow.SelectedRows.Count; i++)
            {
                bool result = true;

                for (int j = 0; j < dgvTestFlow.SelectedRows.Count; j++)
                {
                    int delta = Math.Abs(dgvTestFlow.SelectedRows[j].Index - dgvTestFlow.SelectedRows[i].Index);
                    if (delta == 0)
                        continue;
                    if (delta == 1)
                    {
                        result = true;
                        break;
                    }
                    else
                        result = false;
                }

                if (!result)
                    return result;
            }

            return true;
        }

        private void BindFuncs()
        {
            DataTable dtFunc = new DataTable();
            dtFunc.Columns.Add("TestFunction");

            DataRow drFirstRow;
            drFirstRow = dtFunc.NewRow();
            drFirstRow["TestFunction"] = string.Empty;
            dtFunc.Rows.Add(drFirstRow);

            foreach (var dict in dictFuncs)
            {
                DataRow drFunc;
                drFunc = dtFunc.NewRow();
                drFunc["TestFunction"] = dict.Key;
                dtFunc.Rows.Add(drFunc);
            }

            cboFuncs.ValueMember = "TestFunction";
            cboFuncs.DisplayMember = "TestFunction";
            cboFuncs.DataSource = dtFunc;
            cboFuncs.DropDownStyle = ComboBoxStyle.DropDownList;
            cboFuncs.Visible = false;
        }

        private string TypeTranslator(Type type)
        {
            //if (type.GetInterfaces().Contains(typeof(IEnumerable)))
            if (type.IsGenericType)
            {
                Type[] genericTypes = type.GetGenericArguments();
                return genericTypes[0].Name.ToString() + " Array";
            }

            return type.Name.ToString();
        }

        private object ValueTranslator(string type, string value)
        {
            object res;

            switch (type)
            {
                case "Int32":
                    res = Convert.ToInt32(value);
                    break;
                case "Double":
                    res = Convert.ToDouble(value);
                    break;
                case "Boolean":
                    if (value.ToUpper() == "TRUE")
                    {
                        res = true;
                    }
                    else
                    {
                        res = false;
                    }
                    break;
                case "Int32 Array":
                    res = ListValueTranslator<int>(value, int.Parse);
                    break;
                case "Double Array":
                    res = ListValueTranslator<double>(value, double.Parse);
                    break;
                default:
                    res = value;
                    break;
            }

            return res;
        }

        private List<T> ListValueTranslator<T>(string value, Func<string, T> TPhase)
        {
            List<T> li = new List<T>();
            string[] ls = value.Trim().Split(',');
            foreach (string s in ls)
            {
                li.Add(TPhase(s));
            }
            return li;
        }

        private void setButtonVisible(bool isEditMode)
        {
            btnAdd.Enabled = isEditMode;
            btnClone.Enabled = isEditMode;
            btnDelete.Enabled = isEditMode;
            btnSave.Enabled = isEditMode;
            btnReset.Enabled = isEditMode;
            btnUp.Enabled = isEditMode;
            btnDown.Enabled = isEditMode;
            btnClear.Enabled = isEditMode;

            btnStart.Enabled = !isEditMode;
            btnStop.Enabled = !isEditMode;
            btnRun.Enabled = !isEditMode;
            btnHalt.Enabled = !isEditMode;
        }
        #endregion
    }
}
