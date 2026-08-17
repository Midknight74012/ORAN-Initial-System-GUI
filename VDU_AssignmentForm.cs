using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Renci.SshNet;
using ORAN_Initial_System_GUI;

namespace Dish_ORAN_InitialSystem_GUI
{
    public partial class VDU_AssignmentForm : Form
    {
        VZ_LOLO vZ_LOLO = new VZ_LOLO();
        VZ_PCS vZ_PCS = new VZ_PCS();

        public VDU_AssignmentForm()
        {
            InitializeComponent();
        }
        private string SendSSHCommand(ShellStream stream, string command, string endpoint, int seconds = 15)
        {
            string reader = string.Empty;
            DateTime timeout = DateTime.UtcNow.AddSeconds(seconds);
            stream.Read();
            stream.WriteLine(command);
            while (!reader.Contains(endpoint) && DateTime.UtcNow < timeout)
            {
                if (stream.DataAvailable)
                {
                    reader += stream.Read();
                }
            }
            return reader;
        }
        #region FullPowerButtons
        private void sendFullPowerPCS_Click(object sender, EventArgs e)
        {
            string reader = string.Empty;
            // Disable button immediately on UI thread
            sendFullPowerPCS.Enabled = false;
            progressBarPCS.Value = progressBarPCS.Minimum;

            if (!cb900.Checked && !cb901.Checked && !cb902.Checked)
            {
                // nothing selected — re-enable and return
                sendFullPowerPCS.Enabled = true;
                return;
            }
            int totalCommands = 5;
            if (cb900.Checked) totalCommands += vZ_PCS.full900.Length;
            if (cb901.Checked) totalCommands += vZ_PCS.full901.Length;
            if (cb902.Checked) totalCommands += vZ_PCS.full902.Length;
            int commandsCompleted = 0;
            progressBarPCS.Maximum = totalCommands;
            Task.Run(() => {
                try
                {
                    using var client = new SshClient("REDACTED_IP", "REDACTED_USER", "REDACTED_PASSWORD");
                    client.Connect();
                    using var stream = client.CreateShellStream("xterm", 80, 50, 1024, 1024, 1024);
                    Thread.Sleep(50);

                    // initial handshake commands
                    SendSSHCommand(stream, @"./conn dmp", "/]$");
                    this.Invoke(new MethodInvoker(delegate {
                        progressBarPCS.Value++;
                    }));
                    SendSSHCommand(stream, "nrconfd_cli -u vsmuser", "Welcome to the ConfD CLI");
                    this.Invoke(new MethodInvoker(delegate {
                        progressBarPCS.Value++;
                    }));
                    Thread.Sleep(50);
                    reader += SendSSHCommand(stream, "set paginate false", "[ok]");
                    this.Invoke(new MethodInvoker(delegate {
                        progressBarPCS.Value++;
                    }));
                    Thread.Sleep(50);
                    if (cb900.Checked)
                    {
                        foreach (string command in vZ_PCS.full900)
                        {
                            SendSSHCommand(stream, command, ">");
                            this.Invoke(new MethodInvoker(delegate {
                                progressBarPCS.Value++;
                            }));
                            Thread.Sleep(50);
                        }
                    }

                    if (cb901.Checked)
                    {
                        foreach (string command in vZ_PCS.full901)
                        {
                            SendSSHCommand(stream, command, ">");
                            this.Invoke(new MethodInvoker(delegate {
                                progressBarPCS.Value++;
                            }));
                            Thread.Sleep(50);
                        }
                    }

                    if (cb902.Checked)
                    {
                        foreach (string command in vZ_PCS.full902)
                        {
                            SendSSHCommand(stream, command, ">");
                            this.Invoke(new MethodInvoker(delegate {
                                progressBarPCS.Value++;
                            }));
                            Thread.Sleep(50);
                        }
                    }
                    client.Disconnect();
                }
                catch (Exception ex)
                {
                    // Report error back to UI thread (keep message brief)
                    this.Invoke(new MethodInvoker(delegate {
                        MessageBox.Show($"Error sending commands: {ex.Message}", "SSH Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                }
                finally
                {
                    this.Invoke(new MethodInvoker(delegate {
                        sendFullPowerPCS.Enabled = true;
                        progressBarPCS.Value = progressBarPCS.Maximum;
                    }));
                }
            });

        }
        private void sendFullPowerXLLOLO_Click(object sender, EventArgs e)
        {
            string reader = string.Empty;
            // Disable button immediately on UI thread
            sendFullPowerXLLOLO.Enabled = false;
            progressBarXLLOLO.Value = progressBarXLLOLO.Minimum;

            if (!cb700.Checked && !cb701.Checked && cb702.Checked && cb703.Checked && !cb800.Checked && !cb801.Checked)
            {
                // nothing selected — re-enable and return
                sendFullPowerXLLOLO.Enabled = true;
                return;
            }
            int totalCommands = 5;
            if (cb700.Checked) totalCommands += vZ_LOLO.full700.Length;
            if (cb701.Checked) totalCommands += vZ_LOLO.full701.Length;
            if (cb702.Checked) totalCommands += vZ_LOLO.full702.Length;
            if (cb703.Checked) totalCommands += vZ_PCS.full703.Length;
            if (cb800.Checked) totalCommands += vZ_LOLO.full800.Length;
            if (cb801.Checked) totalCommands += vZ_LOLO.full801.Length;
            int commandsCompleted = 0;
            progressBarXLLOLO.Maximum = totalCommands;
            Task.Run(() => {
                try
                {
                    using var client = new SshClient("REDACTED_IP", "REDACTED_USER", "REDACTED_PASSWORD");
                    client.Connect();
                    using var stream = client.CreateShellStream("xterm", 80, 50, 1024, 1024, 1024);
                    Thread.Sleep(50);

                    // initial handshake commands
                    SendSSHCommand(stream, @"./conn dmp", "/]$");
                    this.Invoke(new MethodInvoker(delegate {
                        progressBarXLLOLO.Value++;
                    }));
                    SendSSHCommand(stream, "nrconfd_cli -u vsmuser", "Welcome to the ConfD CLI");
                    this.Invoke(new MethodInvoker(delegate {
                        progressBarXLLOLO.Value++;
                    }));
                    Thread.Sleep(50);
                    reader += SendSSHCommand(stream, "set paginate false", "[ok]");
                    this.Invoke(new MethodInvoker(delegate {
                        progressBarXLLOLO.Value++;
                    }));
                    Thread.Sleep(50);
                    if (cb700.Checked)
                    {
                        foreach (string command in vZ_LOLO.full700)
                        {
                            SendSSHCommand(stream, command, ">");
                            this.Invoke(new MethodInvoker(delegate {
                                progressBarXLLOLO.Value++;
                            }));
                            Thread.Sleep(50);
                        }
                    }

                    if (cb701.Checked)
                    {
                        foreach (string command in vZ_LOLO.full701)
                        {
                            SendSSHCommand(stream, command, ">");
                            this.Invoke(new MethodInvoker(delegate {
                                progressBarXLLOLO.Value++;
                            }));
                            Thread.Sleep(50);
                        }
                    }
                    if (cb702.Checked)
                    {
                        foreach (string command in vZ_LOLO.full702)
                        {
                            SendSSHCommand(stream, command, ">");
                            this.Invoke(new MethodInvoker(delegate {
                                progressBarXLLOLO.Value++;
                            }));
                            Thread.Sleep(50);
                        }
                    }

                    if (cb703.Checked)
                    {
                        foreach (string command in vZ_PCS.full703)
                        {
                            SendSSHCommand(stream, command, ">");
                            this.Invoke(new MethodInvoker(delegate {
                                progressBarXLLOLO.Value++;
                            }));
                            Thread.Sleep(50);
                        }
                    }
                    if (cb800.Checked)
                    {
                        foreach (string command in vZ_LOLO.full800)
                        {
                            SendSSHCommand(stream, command, ">");
                            this.Invoke(new MethodInvoker(delegate {
                                progressBarXLLOLO.Value++;
                            }));
                            Thread.Sleep(50);
                        }
                    }

                    if (cb801.Checked)
                    {
                        foreach (string command in vZ_LOLO.full801)
                        {
                            SendSSHCommand(stream, command, ">");
                            this.Invoke(new MethodInvoker(delegate {
                                progressBarXLLOLO.Value++;
                            }));
                            Thread.Sleep(50);
                        }
                    }
                    client.Disconnect();
                }
                catch (Exception ex)
                {
                    // Report error back to UI thread (keep message brief)
                    this.Invoke(new MethodInvoker(delegate {
                        MessageBox.Show($"Error sending commands: {ex.Message}", "SSH Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                }
                finally
                {
                    this.Invoke(new MethodInvoker(delegate {
                        sendFullPowerXLLOLO.Enabled = true;
                        progressBarXLLOLO.Value = progressBarXLLOLO.Maximum;
                    }));
                }
            });

        }
        #endregion FullPowerButtons

        #region AssignSerialNumbers
        private void assignSN_PCS_Click(object sender, EventArgs e)
        {
            string reader = string.Empty;
            // Disable button immediately on UI thread
            assignSN_PCS.Enabled = false;
            progressBarPCS.Value = progressBarPCS.Minimum;

            if (!cb900.Checked && !cb901.Checked && !cb902.Checked)
            {
                // nothing selected — re-enable and return
                assignSN_PCS.Enabled = true;
                return;
            }
            int totalCommands = 6;
            if (cb900.Checked) totalCommands++;
            if (cb901.Checked) totalCommands++;
            if (cb902.Checked) totalCommands++;
            int commandsCompleted = 0;
            progressBarPCS.Maximum = totalCommands;

            Task.Run(() => {
                try
                {
                    using var client = new SshClient("REDACTED_IP", "REDACTED_USER", "REDACTED_PASSWORD");
                    client.Connect();
                    using var stream = client.CreateShellStream("xterm", 80, 50, 1024, 1024, 1024);
                    Thread.Sleep(50);

                    // initial handshake commands
                    SendSSHCommand(stream, @"./conn dmp", "/]$");
                    this.Invoke(new MethodInvoker(delegate {
                        progressBarPCS.Value++;
                    }));
                    SendSSHCommand(stream, "nrconfd_cli -u vsmuser", "Welcome to the ConfD CLI");
                    this.Invoke(new MethodInvoker(delegate {
                        progressBarPCS.Value++;
                    }));
                    Thread.Sleep(50);
                    reader += SendSSHCommand(stream, "set paginate false", ">");
                    this.Invoke(new MethodInvoker(delegate {
                        progressBarPCS.Value++;
                    }));
                    Thread.Sleep(50);
                    SendSSHCommand(stream, "configure", "%");
                    if (cb900.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 900 serial-number " + textBox900.Text, "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarPCS.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    if (cb901.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 901 serial-number " + textBox901.Text, "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarPCS.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    if (cb902.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 902 serial-number " + textBox902.Text, "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarPCS.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    SendSSHCommand(stream, "commit", "%");
                    SendSSHCommand(stream, "exit", "[ok]");
                }
                catch (Exception ex)
                {
                    // Report error back to UI thread (keep message brief)
                    this.Invoke(new MethodInvoker(delegate {
                        MessageBox.Show($"Error sending commands: {ex.Message}", "SSH Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                }
                finally
                {
                    this.Invoke(new MethodInvoker(delegate {
                        assignSN_PCS.Enabled = true;
                        progressBarPCS.Value = progressBarPCS.Maximum;
                    }));
                }
            });
        }
        private void assignSN_XLLOLO_Click(object sender, EventArgs e)
        {
            string reader = string.Empty;
            // Disable button immediately on UI thread
            assignSN_XLLOLO.Enabled = false;
            progressBarXLLOLO.Value = progressBarXLLOLO.Minimum;

            if (!cb700.Checked && !cb701.Checked && !cb702.Checked && !cb703.Checked && !cb800.Checked && !cb801.Checked)
            {
                // nothing selected — re-enable and return
                assignSN_XLLOLO.Enabled = true;
                return;
            }
            int totalCommands = 6;
            if (cb700.Checked) totalCommands++;
            if (cb701.Checked) totalCommands++;
            if (cb702.Checked) totalCommands++;
            if (cb703.Checked) totalCommands++;
            if (cb800.Checked) totalCommands++;
            if (cb801.Checked) totalCommands++;
            int commandsCompleted = 0;
            progressBarXLLOLO.Maximum = totalCommands;

            Task.Run(() => {
                try
                {
                    using var client = new SshClient("REDACTED_IP", "REDACTED_USER", "REDACTED_PASSWORD");
                    client.Connect();
                    using var stream = client.CreateShellStream("xterm", 80, 50, 1024, 1024, 1024);
                    Thread.Sleep(50);

                    // initial handshake commands
                    SendSSHCommand(stream, @"./conn dmp", "/]$");
                    this.Invoke(new MethodInvoker(delegate {
                        progressBarXLLOLO.Value++;
                    }));
                    SendSSHCommand(stream, "nrconfd_cli -u vsmuser", "Welcome to the ConfD CLI");
                    this.Invoke(new MethodInvoker(delegate {
                        progressBarXLLOLO.Value++;
                    }));
                    Thread.Sleep(50);
                    reader += SendSSHCommand(stream, "set paginate false", "[ok]");
                    this.Invoke(new MethodInvoker(delegate {
                        progressBarXLLOLO.Value++;
                    }));
                    SendSSHCommand(stream, "configure", "%");
                    Thread.Sleep(50);
                    if (cb700.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 700 serial-number " + textBox700.Text, "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarXLLOLO.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    if (cb701.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 701 serial-number " + textBox701.Text, "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarXLLOLO.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    if (cb702.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 702 serial-number " + textBox702.Text, "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarXLLOLO.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    if (cb703.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 703 serial-number " + textBox703.Text, "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarXLLOLO.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    if (cb800.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 800 serial-number " + textBox800.Text, "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarXLLOLO.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    if (cb801.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 801 serial-number " + textBox801.Text, "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarXLLOLO.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    SendSSHCommand(stream, "commit", "%");
                    SendSSHCommand(stream, "exit", "[ok]");
                }
                catch (Exception ex)
                {
                    // Report error back to UI thread (keep message brief)
                    this.Invoke(new MethodInvoker(delegate {
                        MessageBox.Show($"Error sending commands: {ex.Message}", "SSH Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                }
                finally
                {
                    this.Invoke(new MethodInvoker(delegate {
                        assignSN_XLLOLO.Enabled = true;
                        progressBarXLLOLO.Value = progressBarXLLOLO.Maximum;
                    }));
                }
            });
        }
        #endregion AssignSerialNumbers

        #region ClearSerialNumbers
        private void clearSN_PCS_Click(object sender, EventArgs e)
        {
            string reader = string.Empty;
            // Disable button immediately on UI thread
            clearSN_PCS.Enabled = false;
            progressBarPCS.Value = progressBarPCS.Minimum;

            if (!cb900.Checked && !cb901.Checked && !cb902.Checked)
            {
                // nothing selected — re-enable and return
                clearSN_PCS.Enabled = true;
                return;
            }
            int totalCommands = 5;
            if (cb900.Checked) totalCommands += 2;
            if (cb901.Checked) totalCommands += 2;
            if (cb902.Checked) totalCommands += 2;
            int commandsCompleted = 0;
            progressBarPCS.Maximum = totalCommands;

            Task.Run(() => {
                try
                {
                    using var client = new SshClient("REDACTED_IP", "REDACTED_USER", "REDACTED_PASSWORD");
                    client.Connect();
                    using var stream = client.CreateShellStream("xterm", 80, 50, 1024, 1024, 1024);
                    Thread.Sleep(50);

                    // initial handshake commands
                    SendSSHCommand(stream, @"./conn dmp", "/]$");
                    this.Invoke(new MethodInvoker(delegate {
                        progressBarPCS.Value++;
                    }));
                    SendSSHCommand(stream, "nrconfd_cli -u vsmuser", "Welcome to the ConfD CLI");
                    this.Invoke(new MethodInvoker(delegate {
                        progressBarPCS.Value++;
                    }));
                    Thread.Sleep(50);
                    reader += SendSSHCommand(stream, "set paginate false", "[ok]");
                    this.Invoke(new MethodInvoker(delegate {
                        progressBarPCS.Value++;
                    }));
                    SendSSHCommand(stream, "configure", "%");
                    Thread.Sleep(50);

                    if (cb900.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 900 serial-number S123456789", "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarPCS.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    if (cb901.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 901 serial-number S789456123", "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarPCS.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    if (cb902.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 902 serial-number S654987312", "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarPCS.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    SendSSHCommand(stream, "commit", "%");
                    if (cb900.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 900 serial-number -", "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarPCS.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    if (cb901.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 901 serial-number -", "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarPCS.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    if (cb902.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 902 serial-number -", "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarPCS.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    SendSSHCommand(stream, "commit", "%");
                    SendSSHCommand(stream, "exit", "[ok]");
                }
                catch (Exception ex)
                {
                    // Report error back to UI thread (keep message brief)
                    this.Invoke(new MethodInvoker(delegate {
                        MessageBox.Show($"Error sending commands: {ex.Message}", "SSH Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                }
                finally
                {
                    this.Invoke(new MethodInvoker(delegate {
                        clearSN_PCS.Enabled = true;
                        progressBarPCS.Value = progressBarPCS.Maximum;
                    }));
                }
            });
        }
        private void clearSN_XLLOLO_Click(object sender, EventArgs e)
        {
            string reader = string.Empty;
            // Disable button immediately on UI thread
            clearSN_XLLOLO.Enabled = false;
            progressBarXLLOLO.Value = progressBarXLLOLO.Minimum;

            if (!cb700.Checked && !cb701.Checked && !cb701.Checked && !cb702.Checked && !cb703.Checked && !cb800.Checked && !cb801.Checked)
            {
                // nothing selected — re-enable and return
                clearSN_XLLOLO.Enabled = true;
                return;
            }
            int totalCommands = 5;
            if (cb700.Checked) totalCommands += 2;
            if (cb701.Checked) totalCommands += 2;
            if (cb702.Checked) totalCommands += 2;
            if (cb703.Checked) totalCommands += 2;
            if (cb800.Checked) totalCommands += 2;
            if (cb801.Checked) totalCommands += 2;
            int commandsCompleted = 0;
            progressBarXLLOLO.Maximum = totalCommands;

            Task.Run(() => {
                try
                {
                    using var client = new SshClient("REDACTED_IP", "REDACTED_USER", "REDACTED_PASSWORD");
                    client.Connect();
                    using var stream = client.CreateShellStream("xterm", 80, 50, 1024, 1024, 1024);
                    Thread.Sleep(50);

                    // initial handshake commands
                    SendSSHCommand(stream, @"./conn dmp", "/]$");
                    this.Invoke(new MethodInvoker(delegate {
                        progressBarXLLOLO.Value++;
                    }));
                    SendSSHCommand(stream, "nrconfd_cli -u vsmuser", "Welcome to the ConfD CLI");
                    this.Invoke(new MethodInvoker(delegate {
                        progressBarXLLOLO.Value++;
                    }));
                    Thread.Sleep(50);
                    reader += SendSSHCommand(stream, "set paginate false", "[ok]");
                    this.Invoke(new MethodInvoker(delegate {
                        progressBarXLLOLO.Value++;
                    }));
                    SendSSHCommand(stream, "configure", "%");
                    Thread.Sleep(50);
                    if (cb700.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 700 serial-number S123456789", "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarXLLOLO.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    if (cb701.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 701 serial-number S789456123", "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarXLLOLO.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    if (cb702.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 702 serial-number S123456789", "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarXLLOLO.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    if (cb703.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 703 serial-number S789456123", "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarXLLOLO.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    if (cb800.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 800 serial-number S123456789", "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarXLLOLO.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    if (cb801.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 801 serial-number S789456123", "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarXLLOLO.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    SendSSHCommand(stream, "commit", "%");
                    if (cb700.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 700 serial-number -", "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarXLLOLO.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    if (cb701.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 701 serial-number -", "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarXLLOLO.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    if (cb702.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 702 serial-number -", "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarXLLOLO.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    if (cb703.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 703 serial-number -", "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarXLLOLO.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    if (cb800.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 800 serial-number -", "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarXLLOLO.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    if (cb801.Checked)
                    {
                        SendSSHCommand(stream, "set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info 801 serial-number -", "%");
                        this.Invoke(new MethodInvoker(delegate {
                            progressBarXLLOLO.Value++;
                        }));
                        Thread.Sleep(50);
                    }
                    SendSSHCommand(stream, "commit", "%");
                    SendSSHCommand(stream, "exit", "[ok]");
                }
                catch (Exception ex)
                {
                    // Report error back to UI thread (keep message brief)
                    this.Invoke(new MethodInvoker(delegate {
                        MessageBox.Show($"Error sending commands: {ex.Message}", "SSH Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                }
                finally
                {
                    this.Invoke(new MethodInvoker(delegate {
                        clearSN_XLLOLO.Enabled = true;
                        progressBarXLLOLO.Value = progressBarXLLOLO.Maximum;
                    }));
                }
            });
        }
        #endregion ClearSerialNumbers
    }
}

