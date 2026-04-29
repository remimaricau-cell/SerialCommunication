using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;

namespace SerialCommunication
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                string[] portNames = SerialPort.GetPortNames().Distinct().ToArray();
                comboBoxPoort.Items.Clear();
                comboBoxPoort.Items.AddRange(portNames);
                if (comboBoxPoort.Items.Count > 0) comboBoxPoort.SelectedIndex = 0;

                comboBoxBaudrate.SelectedIndex = comboBoxBaudrate.Items.IndexOf("115200");
            }
            catch (Exception)
            { }
        }

        private void cboPoort_DropDown(object sender, EventArgs e)
        {
            try
            {
                string selected = (string)comboBoxPoort.SelectedItem;
                string[] portNames = SerialPort.GetPortNames().Distinct().ToArray();

                comboBoxPoort.Items.Clear();
                comboBoxPoort.Items.AddRange(portNames);

                comboBoxPoort.SelectedIndex = comboBoxPoort.Items.IndexOf(selected);
            }
            catch (Exception)
            {
                if (comboBoxPoort.Items.Count > 0) comboBoxPoort.SelectedIndex = 0;
            }
        }

        private void buttonConnect_Click(object sender, EventArgs e)
        {
            try
            {
                if (serialPortArduino.IsOpen)
                {
                    serialPortArduino.Close();
                    buttonConnect.Text = "Connect";
                    radioButtonVerbonden.Checked = false;
                    labelStatus.Text = "Verbinding verbroken";
                }
                else
                {
                    if (comboBoxPoort.SelectedItem == null)
                    {
                        MessageBox.Show("Selecteer eerst een poort.", "Geen poort geselecteerd",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    serialPortArduino.PortName = comboBoxPoort.SelectedItem.ToString();
                    serialPortArduino.BaudRate = int.Parse(comboBoxBaudrate.SelectedItem.ToString());
                    serialPortArduino.DataBits = (int)numericUpDownDatabits.Value;
                    serialPortArduino.Parity = GetSelectedParity();
                    serialPortArduino.StopBits = GetSelectedStopBits();
                    serialPortArduino.Handshake = GetSelectedHandshake();
                    serialPortArduino.DtrEnable = checkBoxDtrEnable.Checked;
                    serialPortArduino.RtsEnable = checkBoxRtsEnable.Checked;

                    serialPortArduino.Open();
                    buttonConnect.Text = "Disconnect";
                    radioButtonVerbonden.Checked = true;
                    labelStatus.Text = "Verbonden met " + serialPortArduino.PortName;
                }
            }
            catch (Exception ex)
            {
                buttonConnect.Text = serialPortArduino.IsOpen ? "Disconnect" : "Connect";
                radioButtonVerbonden.Checked = serialPortArduino.IsOpen;
                labelStatus.Text = ex.Message;

                MessageBox.Show(ex.Message, "Verbindingsfout",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private Parity GetSelectedParity()
        {
            if (radioButtonParityEven.Checked) return Parity.Even;
            if (radioButtonParityOdd.Checked) return Parity.Odd;
            if (radioButtonParityMark.Checked) return Parity.Mark;
            if (radioButtonParitySpace.Checked) return Parity.Space;

            return Parity.None;
        }

        private StopBits GetSelectedStopBits()
        {
            if (radioButtonStopbitsNone.Checked) return StopBits.None;
            if (radioButtonStopbitsOnePointFive.Checked) return StopBits.OnePointFive;
            if (radioButtonStopbitsTwo.Checked) return StopBits.Two;

            return StopBits.One;
        }

        private Handshake GetSelectedHandshake()
        {
            if (radioButtonHandshakeRTS.Checked) return Handshake.RequestToSend;
            if (radioButtonHandshakeRTSXonXoff.Checked) return Handshake.RequestToSendXOnXOff;
            if (radioButtonHandshakeXonXoff.Checked) return Handshake.XOnXOff;

            return Handshake.None;
        }
    }
}
