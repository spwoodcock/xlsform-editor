import React from 'react';
import { IgrExcelXlsxModule } from 'igniteui-react-excel';
import { IgrExcelCoreModule } from 'igniteui-react-excel';
import { IgrExcelModule } from 'igniteui-react-excel';
import { IgrSpreadsheetModule } from 'igniteui-react-spreadsheet';
import { IgrSpreadsheet } from 'igniteui-react-spreadsheet';
import { ExcelUtility } from './ExcelUtility';

IgrExcelCoreModule.register();
IgrExcelModule.register();
IgrExcelXlsxModule.register();

IgrSpreadsheetModule.register();

class SpreadsheetOverview extends React.Component<any, any> {
    public spreadsheet: IgrSpreadsheet;

    constructor(props: any) {
        super(props);
        this.onSpreadsheetRef = this.onSpreadsheetRef.bind(this);
        this.openFile = this.openFile.bind(this);
        this.saveFile = this.saveFile.bind(this);
    }

    public render(): JSX.Element {
        return (
            <div className="container sample">
                <div className="options horizontal">
                    <input type="file" onChange={(e) => this.openFile(e.target.files as FileList)} accept=".xls, .xlt, .xlsx, .xlsm, .xltm, .xltx"/>
                    <button onClick={this.saveFile}>Save</button>
                </div>
                <IgrSpreadsheet ref={this.onSpreadsheetRef} height="calc(100% - 30px)" width="100%" />
            </div>
        );
    }

    public openFile(selectorFiles: FileList) {
        if (selectorFiles != null && selectorFiles.length > 0) {
            ExcelUtility.load(selectorFiles[0]).then((w) => {
                this.spreadsheet.workbook = w;
            }, (e) => {
                console.error("Workbook Load Error");
            });
        }
    }

    public saveFile() {
        ExcelUtility.save(this.spreadsheet.workbook, 'form');
    }

    public onSpreadsheetRef(spreadsheet: IgrSpreadsheet) {
        if (!spreadsheet) { return; }

        this.spreadsheet = spreadsheet;
    }
}

function App() {
  return (
    <SpreadsheetOverview />
  )
}

export default App
