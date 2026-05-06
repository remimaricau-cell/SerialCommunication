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

        private const int LedPin = 2;
        private const int BuzzerPin = 3;
        private const int BevestigKnopPin = 5;

        private Timer timerOefening5;
        private AlarmToestand alarmToestand = AlarmToestand.OK;
        private Label labelToestandTitel;
        private Label labelToestand;

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
            tabControl.SelectedIndexChanged += tabControl_SelectedIndexChanged;
            InitializeOefening5StatusLabels();
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
            UpdateTimerOefening5();
        }

        private void UpdateTimerOefening5()
        {
            if (tabControl.SelectedTab == tabPageOefening5)
            {
                timerOefening5.Start();
            }
            else
            {
                timerOefening5.Stop();
            }
        }

        private void InitializeOefening5StatusLabels()
        {
            label9.Text = "Alarmwaarde";

            labelToestandTitel = new Label();
            labelToestandTitel.AutoSize = true;
            labelToestandTitel.Font = label9.Font;
            labelToestandTitel.Location = new Point(694, 464);
            labelToestandTitel.Text = "Toestand";

            labelToestand = new Label();
            labelToestand.Font = labelGewensteTemp.Font;
            labelToestand.Location = new Point(886, 458);
            labelToestand.Size = new Size(190, 35);
            labelToestand.TextAlign = System.Drawing.ContentAlignment.MiddleRight;

            tabPageOefening5.Controls.Add(labelToestandTitel);
            tabPageOefening5.Controls.Add(labelToestand);
            UpdateToestandLabel();
        }

        private void timerOefening5_Tick(object sender, EventArgs e)
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
                bool alarmBevestigd = ReadDigitalInput(BevestigKnopPin);

                double alarmRichtingscoefficient = (60.0 - -10.0) / 1023.0;
                double alarmOffset = -10.0;
                double alarmTemperatuur = alarmRichtingscoefficient * analog1 + alarmOffset;

                double huidigeRichtingscoefficient = (500.0 - 0.0) / 1023.0;
                double huidigeOffset = 0.0;
                double huidigeTemperatuur = huidigeRichtingscoefficient * analog0 + huidigeOffset;

                labelGewensteTemp.Text = alarmTemperatuur.ToString("0.0") + " °C";
                labelHuidigeTemp.Text = huidigeTemperatuur.ToString("0.0") + " °C";

                UpdateAlarmToestand(huidigeTemperatuur, alarmTemperatuur, alarmBevestigd);
                ApplyAlarmOutputs();
                UpdateToestandLabel();
            }
            catch (Exception ex)
            {
                labelStatus.Text = ex.Message;
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

        private void UpdateAlarmToestand(double huidigeTemperatuur, double alarmTemperatuur, bool alarmBevestigd)
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
                    if (alarmBevestigd)
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

        private void UpdateToestandLabel()
        {
            labelToestand.Text = alarmToestand.ToString();
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
