using Renci.SshNet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Renci.SshNet;

namespace ORAN_Initial_System_GUI
{
    public class VZ_LOLO
    {
        byte[] interrupt = new byte[] { 0x03 };

        DataGridViewButtonCell[] startButton = new DataGridViewButtonCell[4];
        DataGridViewButtonCell[] stopButton = new DataGridViewButtonCell[4];
        DataGridViewButtonCell[] clearButton = new DataGridViewButtonCell[4];
        enum LoloDataRow : int
        {
            SN = 1,
            SN_Check = 2,
            Firmware = 3,
            Full_Power = 4,
            RET_Signal = 5,
            MAC_Addresses = 6,
            Measurements = 7,
            Alarms = 8,
            CSV_Logging = 9,
            Result = 10,
            Timer = 11,
            Start_Button_Row = 12,
            Clear_Button_Row = 13,
            Show_Log = 14,
        }
        public List<string> testInfo = new List<string>
            {
            "S/N", //Never edit or remove
            "S/N Check",
            "Firmware",
            "Full Power",
            "RET Signal",
            "MAC Addresses",
            "Measurements",
            "Alarms",
            "CSV Logging",
            "Result",
            "Timer"
        };

        public string[] full700 = {
            "request gnbdu-func:test-model-b6-mplane cell-identity 17 model-id 5g-tm-3-1a nr-physical-cell-id 1",
            "request gnbdu-func:test-model-b6-mplane cell-identity 18 model-id 5g-tm-3-1a nr-physical-cell-id 2",
            "request test-model-lte model-id etm-3-1-a cell-number 2 physical-cell-id 3",
            "request test-model-lte model-id etm-3-1-a cell-number 22 physical-cell-id 4"};

        public string[] full701 = {
            "request gnbdu-func:test-model-b6-mplane model-id 5g-tm-3-1a cell-identity 33 nr-physical-cell-id 5",
            "request gnbdu-func:test-model-b6-mplane model-id 5g-tm-3-1a cell-identity 34 nr-physical-cell-id 6",
            "request test-model-lte model-id etm-3-1-a cell-number 3 physical-cell-id 7",
            "request test-model-lte model-id etm-3-1-a cell-number 32 physical-cell-id 8"};

        public string[] full702 = {
            "request gnbdu-func:test-model-b6-mplane model-id 5g-tm-3-1a cell-identity 49 nr-physical-cell-id 9",
            "request gnbdu-func:test-model-b6-mplane model-id 5g-tm-3-1a cell-identity 50 nr-physical-cell-id 10",
            "request test-model-lte model-id etm-3-1-a cell-number 4 physical-cell-id 11",
            "request test-model-lte model-id etm-3-1-a cell-number 42 physical-cell-id 12"};

        public string[] full800 = {
            "request gnbdu-func:test-model-b6-mplane model-id 5g-tm-3-1a cell-identity 65 nr-physical-cell-id 65",
            "request gnbdu-func:test-model-b6-mplane model-id 5g-tm-3-1a cell-identity 66 nr-physical-cell-id 66",
            "request test-model-lte model-id etm-3-1-a cell-number 5 physical-cell-id 19",
            "request test-model-lte model-id etm-3-1-a cell-number 52 physical-cell-id 20"};

        public string[] full801 = {
            "request gnbdu-func:test-model-b6-mplane model-id 5g-tm-3-1a cell-identity 81 nr-physical-cell-id 21",
            "request gnbdu-func:test-model-b6-mplane model-id 5g-tm-3-1a cell-identity 82 nr-physical-cell-id 22",
            "request test-model-lte model-id etm-3-1-a cell-number 6 physical-cell-id 23",
            "request test-model-lte model-id etm-3-1-a cell-number 62 physical-cell-id 24"};

        /*public string[] full802 = {
            "request test-model-b6-mplane cell-identity 98 nr-physical-cell-id 11 model-id 5g-tm-3-1a",
            "request test-model-b6-mplane cell-identity 99 nr-physical-cell-id 12 model-id 5g-tm-3-1a",
            "request test-model-lte cell-number 7 physical-cell-id 21 model-id etm-3-1-a",
            "request test-model-lte cell-number 72 physical-cell-id 22 model-id etm-3-1-a"};*/

        public void LoloGridSetup(DataGridView grid) {
            grid.Rows.Add("RU ID", "700", "701", "702");
            for (int i = 0; i < testInfo.Count; i++) {
                grid.Rows.Add(testInfo[i]);
                grid.Rows[i + 1].ReadOnly = true;
            }
            for (int j = (int)LoloDataRow.SN_Check; j <= (int)LoloDataRow.Result; j++) {
                for (int i = 1; i < 4; i++) {
                    grid.Rows[(int)LoloDataRow.Timer].Cells[i].Value = "00:00:00";
                    grid.Rows[j].Cells[i] = new DataGridViewImageCell();
                    grid.Rows[j].Cells[i].Value = Image.FromFile(@"Resources\white.png");
                }
            }
            grid.Rows.Add("Start");
            grid.Rows.Add("Clear");

            for (int i = 1; i < 4; i++) {
                DataGridViewButtonCell btnCell = new DataGridViewButtonCell();
                grid.Columns[i].Width = (grid.Width / grid.ColumnCount) - 10;
                btnCell.Value = "Start Test";
                startButton[i - 1] = btnCell;
                grid.Rows[(int)LoloDataRow.Start_Button_Row].Cells[i] = btnCell;

                DataGridViewButtonCell stopBtn = new DataGridViewButtonCell();
                stopBtn.Value = "Stop Aging";
                stopButton[i - 1] = stopBtn;

                DataGridViewButtonCell clearCell = new DataGridViewButtonCell();
                clearCell.Value = "Clear";
                clearButton[i - 1] = clearCell;
                grid.Rows[(int)LoloDataRow.Clear_Button_Row].Cells[i] = clearCell;
            }
            grid.Rows.Add("Show Log");

            for (int i = 1; i < 4; i++) {
                DataGridViewButtonCell btnCell = new DataGridViewButtonCell();
                grid.Columns[i].Width = (grid.Width / grid.ColumnCount) - 10;
                btnCell.Value = "Show Log";
                startButton[i - 1] = btnCell;
                grid.Rows[(int)LoloDataRow.Show_Log].Cells[i] = btnCell;
            }
        }
    }
}
