using System;
using ExcelInterop.InteropApi;

namespace ExcelInterop.Utils {
    public static class ExcelApiUtils {

        public static void CloseExcelApi(ref ExcelApi excelApi) {
            excelApi.Close();
            excelApi = null;
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
    }
}
