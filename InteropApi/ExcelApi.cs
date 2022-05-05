using System.Linq;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelInterop.InteropApi {

    public class ExcelApi {

        private readonly Excel.Application _application;
        private readonly Excel.Workbooks _workbooks;

        public ExcelApi(bool visible = false) {
            this._application = new Excel.Application {
                WindowState = Excel.XlWindowState.xlMinimized,
                Visible = visible
            };
            this._workbooks = this._application.Workbooks;
        }

        private Dictionary<string, ExcelWorkbookApi> _workbookDict = new Dictionary<string, ExcelWorkbookApi>();
        public readonly ObservableCollection<string> obsWorkbooks = new ObservableCollection<string>();

        public ExcelWorkbookApi this[int index] => this._workbookDict[this.obsWorkbooks[index]];
        public ExcelWorkbookApi this[string name] => this._workbookDict[name];

        internal void _Add(Excel.Workbook workbook) {
            this._workbookDict.Add(workbook.Name, new ExcelWorkbookApi(this, workbook));
            this.obsWorkbooks.Add(workbook.Name);
        }

        internal void _Remove(string workbookName) {
            _ = this._workbookDict.Remove(workbookName);
            _ = this.obsWorkbooks.Remove(workbookName);
        }

        public void OpenWorkbook() {
            this._Add(this._workbooks.Add());
        }

        public void Close() {
            string[] workbookNames = this.obsWorkbooks.ToArray();
            foreach (string workbooName in workbookNames) {
                this[workbooName].Close();
            }
            this._application.Quit();
        }
    }
}
