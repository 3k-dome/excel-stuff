using System.Collections.ObjectModel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelInterop.InteropApi {

    public class ExcelWorkbookApi {

        private readonly ExcelApi _parentApi;
        private readonly Excel.Workbook _workbook;
        private readonly Excel.Sheets _sheets;

        public ObservableCollection<string> obsWorksheets = new ObservableCollection<string>();

        public ExcelWorkbookApi(ExcelApi parentApi, Excel.Workbook workbook) {
            this._parentApi = parentApi;
            this._workbook = workbook;
            this._sheets = this._workbook.Worksheets;
            foreach (Excel.Worksheet worksheet in this._sheets) {
                this.obsWorksheets.Add(worksheet.Name);
            }
        }

        public Excel.Worksheet this[int index] => this._sheets.Item[this.obsWorksheets[index]];
        public Excel.Worksheet this[string name] => this._sheets.Item[name];

        public void AddSheet(string name = null) {
            Excel.Worksheet sheet = this._sheets.Add();
            if (name != null) { sheet.Name = name; }
            this.obsWorksheets.Add(sheet.Name);
        }

        #region "RemoveSheet"
        public void RemoveSheet(int index) {
            Excel.Worksheet sheet = this[index];
            this._RemoveSheet(sheet);
        }

        public void RemoveSheet(string name) {
            Excel.Worksheet sheet = this[name];
            this._RemoveSheet(sheet);
        }

        private void _RemoveSheet(Excel.Worksheet sheet) {
            _ = this.obsWorksheets.Remove(sheet.Name);
            sheet.Delete();
        }
        #endregion

        #region "RenameSheet"
        public void RenameSheet(int index, string newName) {
            Excel.Worksheet sheet = this[index];
            this._RenameSheet(sheet, newName);
        }

        public void RenameSheet(string name, string newName) {
            Excel.Worksheet sheet = this[name];
            this._RenameSheet(sheet, newName);
        }

        private void _RenameSheet(Excel.Worksheet sheet, string newName) {
            string oldName = sheet.Name;
            sheet.Name = newName;
            this.obsWorksheets[this.obsWorksheets.IndexOf(oldName)] = newName;
        }
        #endregion

        public void Close() {
            this._parentApi._Remove(this._workbook.Name);
            this._workbook.Close();
        }
    }
}
