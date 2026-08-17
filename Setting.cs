using ORAN_Initial_System_GUI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Dish_ORAN_InitialSystem_GUI
{

    public partial class Setting : Form
    {
        public Setting() {
            InitializeComponent();
        }
        string Settings = @"C:\Carrier Oran Settings";
        string LOLOIP = @"C:\Carrier Oran Settings\LOLO IP.txt";
        string PCSIP = @"C:\Carrier Oran Settings\PCS Band IP.txt";
        string LOLORuId = @"C:\Carrier Oran Settings\LOLO RU ID.txt";
        string PCSRuId = @"C:\Carrier Oran Settings\PCS RU ID.txt";
        string FirmwareFile = @"C:\Carrier Oran Settings\Firmware.txt";
        public bool getJoke = true;


        private void button1_Click(object sender, EventArgs e) {
            string firmware = Firmwareinfo.Text; // Get firmware input from TextBox
            string ipaddress = comboBox2.Text;
            string ruid = comboBox3.Text;
            string modelid = comboBox1.Text;
            Directory.CreateDirectory(Settings);
            string FirmwareFile = Path.Combine(Settings, "Firmware.txt");
            File.WriteAllText(FirmwareFile, firmware);

            if (checkBox1.Checked) {
                using (FileStream fstriruid = File.Create(LOLORuId)) {
                    // Add some text to file
                    Byte[] ruidtri = new UTF8Encoding(true).GetBytes(ruid);
                    fstriruid.Write(ruidtri, 0, ruidtri.Length);
                }
                using (FileStream fsdualruid = File.Create(PCSRuId)) {
                    // Add some text to file
                    Byte[] ruiddual = new UTF8Encoding(true).GetBytes(ruid);
                    fsdualruid.Write(ruiddual, 0, ruiddual.Length);
                }
                this.Close();
            } else {
                try {
                    if (modelid == "LOLO") {
                        using (FileStream fstri = File.Create(LOLOIP)) {
                            // Add some text to file
                            Byte[] ipaddrestri = new UTF8Encoding(true).GetBytes(ipaddress);
                            fstri.Write(ipaddrestri, 0, ipaddrestri.Length);
                        }
                        using (FileStream fstriruid = File.Create(LOLORuId)) {
                            // Add some text to file
                            Byte[] ruidtri = new UTF8Encoding(true).GetBytes(ruid);
                            fstriruid.Write(ruidtri, 0, ruidtri.Length);
                        }


                    }

                    if (modelid == "PCS") {
                        using (FileStream fstri = File.Create(PCSIP)) {
                            // Add some text to file
                            Byte[] ipaddrestri = new UTF8Encoding(true).GetBytes(ipaddress);
                            fstri.Write(ipaddrestri, 0, ipaddrestri.Length);
                        }
                        using (FileStream fsdualruid = File.Create(PCSRuId)) {
                            // Add some text to file
                            Byte[] ruiddual = new UTF8Encoding(true).GetBytes(ruid);
                            fsdualruid.Write(ruiddual, 0, ruiddual.Length);
                        }

                    }

                    this.Close();
                }
                catch {
                    MessageBox.Show("One of the item was not Selected..!");

                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) {
            switch (comboBox1.SelectedIndex) {
                case 0:
                    comboBox3.Items.Clear();
                    comboBox3.Items.Add("701");
                    comboBox3.Items.Add("702");
                    comboBox3.Items.Add("703");
                    break;

                case 1:
                    comboBox3.Items.Clear();
                    comboBox3.Items.Add("700");
                    comboBox3.Items.Add("701");
                    comboBox3.Items.Add("702");
                    comboBox3.Items.Add("800");
                    comboBox3.Items.Add("801");
                    comboBox3.Items.Add("802");
                    break;
            }
        }

        private void label3_DoubleClick(object sender, EventArgs e) {
            if (getJoke == true) {
                getJoke = false;
                MessageBox.Show("Jokes off");
            }else {
                getJoke = true;
                MessageBox.Show("Jokes on");
            }
        }
    }
}
