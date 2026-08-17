using Renci.SshNet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Renci.SshNet;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Reflection;
using System.Reflection.Metadata.Ecma335;
using System.Reflection.PortableExecutable;
using System.Security.Claims;
using System.Data;
using Microsoft.VisualBasic.FileIO;

namespace ORAN_Initial_System_GUI
{
    public class VZ_PCS
    {

        public string[] full900 = {
            "request gnbdu-func:test-model-b6-mplane cell-identity 17 model-id 5g-tm-3-1a nr-physical-cell-id 1",
            "request gnbdu-func:test-model-b6-mplane cell-identity 18 model-id 5g-tm-3-1a nr-physical-cell-id 2",
            "request gnbdu-func:test-model-lte cell-number 6 model-id etm-3-1-a physical-cell-id 3",
            "request gnbdu-func:test-model-lte cell-number 62 model-id etm-3-1-a physical-cell-id 4"};

        public string[] full901 = {
            "request gnbdu-func:test-model-b6-mplane cell-identity 33 model-id 5g-tm-3-1a nr-physical-cell-id 3",
            "request gnbdu-func:test-model-b6-mplane cell-identity 34 model-id 5g-tm-3-1a nr-physical-cell-id 4",
            "request test-model-lte cell-number 7 model-id etm-3-1-a physical-cell-id 5",
            "request test-model-lte cell-number 72 model-id etm-3-1-a physical-cell-id 6"};

        public string[] full902 = {
            "request gnbdu-func:test-model-b6-mplane cell-identity 49 model-id 5g-tm-3-1a nr-physical-cell-id 5",
            "request gnbdu-func:test-model-b6-mplane cell-identity 50 model-id 5g-tm-3-1a nr-physical-cell-id 6",
            "request test-model-lte cell-number 8 model-id etm-3-1-a physical-cell-id 7",
            "request test-model-lte cell-number 82 model-id etm-3-1-a physical-cell-id 8"};

        public string[] full703 = {
            "request gnbdu-func:test-model-b6-mplane model-id 5g-tm-3-1a cell-identity 97 nr-physical-cell-id 13",
            "request gnbdu-func:test-model-b6-mplane model-id 5g-tm-3-1a cell-identity 98 nr-physical-cell-id 14",
            "request test-model-lte model-id etm-3-1-a cell-number 7 physical-cell-id 15",
            "request test-model-lte model-id etm-3-1-a cell-number 72 physical-cell-id 16"};

    };
}

    

