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
        private enum AlarmToestand
        {
            OK,
            ALARM,
            BEVESTIGD
        }

        private const int AlarmPotPin = 0;
        private const int Lm35Pin = 1;
        private const int LedPin = 2;
        private const int BuzzerPin = 3;
        private const int BevestigKnopPin = 5;

        private Timer timerOefening5;
        private Timer timerOefening6;
        private AlarmToestand alarmToestand = AlarmToestand.OK;

        public Form1()
        {
            InitializeComponent();
            checkBoxDigital2.CheckedChanged += checkBoxDigital2_CheckedChanged;
            checkBoxDigital3.CheckedChanged += checkBoxDigital3_CheckedChanged;
            checkBoxDigital4.CheckedChanged += checkBoxDigital4_CheckedChanged;
            trackBarPWM9.Scroll += trackBarPWM9_Scroll;
            trackBarPWM10.Scroll += trackBarPWM10_Scroll;
            trackBarPWM11.Scroll += trackBarPWM11_Scroll;

            timerOefening5 = new Timer(components);
            timerOefening5.Interval = 1000;
            timerOefening5.Tick += timerOefening5_Tick;

            timerOefening6 = new Timer(components);
            timerOefening6.Interval = 250;
            timerOefening6.Tick += timerOefening6_Tick;

            tabControl.SelectedIndexChanged += tabControl_SelectedIndexChanged;
            UpdateTimerOefening5();
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
            bool verbindingMaken = !serialPortArduino.IsOpen;

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
                    serialPortArduino.WriteLine("ping");

                    string antwoord = serialPortArduino.ReadLine().Trim();
                    if (antwoord != "pong")
                    {
                        serialPortArduino.Close();
                        buttonConnect.Text = "Connect";
                        radioButtonVerbonden.Checked = false;
                        labelStatus.Text = "Geen geldig antwoord van Arduino";

                        MessageBox.Show("Arduino antwoordde niet met pong.", "Verbindingsfout",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    buttonConnect.Text = "Disconnect";
                    radioButtonVerbonden.Checked = true;
                    labelStatus.Text = "Verbonden met " + serialPortArduino.PortName;
                }
            }
            catch (Exception ex)
            {
                if (verbindingMaken && serialPortArduino.IsOpen)
                {
                    serialPortArduino.Close();
                }

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

        private void checkBoxDigital2_CheckedChanged(object sender, EventArgs e)
        {
            SendDigitalOutputCommand(2, checkBoxDigital2.Checked);
        }

        private void checkBoxDigital3_CheckedChanged(object sender, EventArgs e)
        {
            SendDigitalOutputCommand(3, checkBoxDigital3.Checked);
        }

        private void checkBoxDigital4_CheckedChanged(object sender, EventArgs e)
        {
            SendDigitalOutputCommand(4, checkBoxDigital4.Checked);
        }

        private void SendDigitalOutputCommand(int pin, bool high)
        {
            try
            {
                if (!serialPortArduino.IsOpen)
                {
                    labelStatus.Text = "Geen open seriële verbinding";
                    return;
                }

                string command = "set d" + pin + (high ? " high" : " low");
                serialPortArduino.WriteLine(command);
                string response = serialPortArduino.ReadLine().Trim();
                if (response != "set done")
                {
                    labelStatus.Text = response;
                    MessageBox.Show(response, "Fout digitale uitgang",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                labelStatus.Text = command;
            }
            catch (Exception ex)
            {
                labelStatus.Text = ex.Message;
                MessageBox.Show(ex.Message, "Fout digitale uitgang",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void trackBarPWM9_Scroll(object sender, EventArgs e)
        {
            SendPwmOutputCommand(9, trackBarPWM9.Value);
        }

        private void trackBarPWM10_Scroll(object sender, EventArgs e)
        {
            SendPwmOutputCommand(10, trackBarPWM10.Value);
        }

        private void trackBarPWM11_Scroll(object sender, EventArgs e)
        {
            SendPwmOutputCommand(11, trackBarPWM11.Value);
        }

        private void SendPwmOutputCommand(int pin, int value)
        {
            try
            {
                if (!serialPortArduino.IsOpen)
                {
                    labelStatus.Text = "Geen open seriële verbinding";
                    return;
                }

                string command = "set pwm" + pin + " " + value;
                serialPortArduino.WriteLine(command);
                string response = serialPortArduino.ReadLine().Trim();
                if (response != "set done")
                {
                    labelStatus.Text = response;
                    MessageBox.Show(response, "Fout analoge uitgang",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                labelStatus.Text = command;
            }
            catch (Exception ex)
            {
                labelStatus.Text = ex.Message;
                MessageBox.Show(ex.Message, "Fout analoge uitgang",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl.SelectedTab == tabPageOefening6)
            {
                alarmToestand = AlarmToestand.OK;
                VisualiseerAlarmToestand();

                try
                {
                    if (serialPortArduino.IsOpen)
                    {
                        ApplyAlarmOutputs();
                    }
                }
                catch (Exception ex)
                {
                    HandleSerialRuntimeError(ex);
                }
            }

            UpdateTimerOefening5();
        }

        private void UpdateTimerOefening5()
        {
            timerOefening5.Enabled = tabControl.SelectedTab == tabPageOefening5;
            timerOefening6.Enabled = tabControl.SelectedTab == tabPageOefening6;
        }

        private void timerOefening5_Tick(object sender, EventArgs e)
        {
            HandleOefening5TimerTick();
        }

        private void timerOefening6_Tick(object sender, EventArgs e)
        {
            HandleOefening6TimerTick();
        }

        private void HandleOefening5TimerTick()
        {
            try
            {
                if (!serialPortArduino.IsOpen)
                {
                    labelStatus.Text = "Geen open seriële verbinding";
                    return;
                }

                int analog0 = ReadAnalogInput(0);
                int analog1 = ReadAnalogInput(1);

                double gewensteRichtingscoefficient = (45.0 - 5.0) / 1023.0;
                double gewensteOffset = 5.0;
                double gewensteTemperatuur = gewensteRichtingscoefficient * analog1 + gewensteOffset;

                double huidigeRichtingscoefficient = (500.0 - 0.0) / 1023.0;
                double huidigeOffset = 0.0;
                double huidigeTemperatuur = huidigeRichtingscoefficient * analog0 + huidigeOffset;

                labelGewensteTemp.Text = gewensteTemperatuur.ToString("0.0") + " °C";
                labelHuidigeTemp.Text = huidigeTemperatuur.ToString("0.0") + " °C";

                WriteSetCommand("set d2 " + (huidigeTemperatuur < gewensteTemperatuur ? "low" : "high"));
            }
            catch (Exception ex)
            {
                HandleSerialRuntimeError(ex);
            }
        }

        private void HandleOefening6TimerTick()
        {
            try
            {
                if (!serialPortArduino.IsOpen)
                {
                    labelStatus.Text = "Geen open seriële verbinding";
                    return;
                }

                bool bevestigd = ReadDigitalInput(BevestigKnopPin);
                int alarmValue = ReadAnalogInput(AlarmPotPin);
                int lm35Value = ReadAnalogInput(Lm35Pin);

                double alarmRichtingscoefficient = (100.0 - -10.0) / 1023.0;
                double alarmOffset = -10.0;
                double alarmTemperatuur = alarmRichtingscoefficient * alarmValue + alarmOffset;

                double huidigeTemperatuur = lm35Value * 500.0 / 1023.0;

                labelAlarmTemp.Text = alarmTemperatuur.ToString("0.0") + " °C";
                label15.Text = huidigeTemperatuur.ToString("0.0") + " °C";

                UpdateAlarmToestand(huidigeTemperatuur, alarmTemperatuur, bevestigd);
                ApplyAlarmOutputs();
                VisualiseerAlarmToestand();
            }
            catch (Exception ex)
            {
                HandleSerialRuntimeError(ex);
            }
        }

        private int ReadAnalogInput(int pin)
        {
            serialPortArduino.WriteLine("get a" + pin);

            string response = serialPortArduino.ReadLine().Trim();
            string prefix = "a" + pin + ": ";
            if (!response.StartsWith(prefix))
            {
                throw new InvalidOperationException(response);
            }

            string valueText = response.Substring(prefix.Length);
            int value;
            if (!int.TryParse(valueText, out value))
            {
                throw new InvalidOperationException("Ongeldige analoge waarde: " + response);
            }

            return value;
        }

        private bool ReadDigitalInput(int pin)
        {
            serialPortArduino.WriteLine("get d" + pin);

            string response = serialPortArduino.ReadLine().Trim();
            string prefix = "d" + pin + ": ";
            if (!response.StartsWith(prefix))
            {
                throw new InvalidOperationException(response);
            }

            string valueText = response.Substring(prefix.Length);
            if (valueText == "0") return false;
            if (valueText == "1") return true;

            throw new InvalidOperationException("Ongeldige digitale waarde: " + response);
        }

        private void UpdateAlarmToestand(double huidigeTemperatuur, double alarmTemperatuur, bool bevestigd)
        {
            switch (alarmToestand)
            {
                case AlarmToestand.OK:
                    if (huidigeTemperatuur > alarmTemperatuur)
                    {
                        alarmToestand = AlarmToestand.ALARM;
                    }
                    break;

                case AlarmToestand.ALARM:
                    if (bevestigd)
                    {
                        alarmToestand = huidigeTemperatuur > alarmTemperatuur
                            ? AlarmToestand.BEVESTIGD
                            : AlarmToestand.OK;
                    }
                    break;

                case AlarmToestand.BEVESTIGD:
                    if (huidigeTemperatuur < alarmTemperatuur)
                    {
                        alarmToestand = AlarmToestand.OK;
                    }
                    break;
            }
        }

        private void ApplyAlarmOutputs()
        {
            bool ledAan = alarmToestand == AlarmToestand.ALARM || alarmToestand == AlarmToestand.BEVESTIGD;
            bool buzzerAan = alarmToestand == AlarmToestand.ALARM;

            WriteSetCommand("set d" + LedPin + (ledAan ? " high" : " low"));
            WriteSetCommand("set d" + BuzzerPin + (buzzerAan ? " high" : " low"));
        }

        private void VisualiseerAlarmToestand()
        {
            label16.Text = alarmToestand.ToString();
            labelStatus.Text = alarmToestand.ToString();
        }

        private void HandleSerialRuntimeError(Exception ex)
        {
            labelStatus.Text = ex.Message;

            if (ex is TimeoutException || ex is System.IO.IOException || ex is UnauthorizedAccessException)
            {
                try
                {
                    if (serialPortArduino.IsOpen)
                    {
                        serialPortArduino.Close();
                    }
                }
                catch (Exception)
                { }

                timerOefening5.Stop();
                timerOefening6.Stop();
                buttonConnect.Text = "Connect";
                radioButtonVerbonden.Checked = false;
            }
        }

        private void WriteSetCommand(string command)
        {
            serialPortArduino.WriteLine(command);

            string response = serialPortArduino.ReadLine().Trim();
            if (response != "set done")
            {
                throw new InvalidOperationException(response);
            }
        }
    }
}
