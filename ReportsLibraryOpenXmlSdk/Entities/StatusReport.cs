using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportsLibraryOpenXmlSdk.Entities
{
    public static class StatusReport
    {
        private static readonly string _examinedTrue = "Отчет испытуемый - сформирован!";
        private static readonly string _workloadTrue = "Отчет нагрузочный - сформирован!";

        private static readonly string _examinedFalse = "Отчет испытуемый - не сформирован!";
        private static readonly string _workloadFalse = "Отчет нагрузочный - не сформирован!";

        private static readonly string _status = "Подождите. Идет формирование отчетов!";
        
        private static bool? _examined = null;
        private static bool? _workload = null;

        public static string Status
        {
            get 
            {
                if (_examined == true && _workload == true)
                {
                    return $"{_examinedTrue} {_workloadTrue}";
                }
                else if (_examined == false && _workload == false)
                {
                    return $"{_examinedFalse} {_workloadFalse}";
                }
                else if (_examined == true && _workload == false)
                {
                    return $"{_examinedTrue} {_workloadFalse}";
                }
                else if (_examined == false && _workload == true)
                {
                    return $"{_examinedFalse} {_workloadTrue}";
                }
                else
                {
                    return _status;
                }
            }
        }

        public static void SetStatusExaminedError() => _examined = false;

        public static void SetStatusWorkloadError() => _workload = false;

        public static void SetStatusExaminedReady() => _examined = true;

        public static void SetStatusWorkloadReady() => _workload = true;

    }
}
