import { saveAs } from 'file-saver';
import { Workbook } from 'igniteui-webcomponents-excel';
import { WorkbookFormat } from 'igniteui-webcomponents-excel';
import { WorkbookSaveOptions } from 'igniteui-webcomponents-excel';
import { WorkbookLoadOptions } from 'igniteui-webcomponents-excel';
import { IgcExcelXlsxModule } from 'igniteui-webcomponents-excel';
import { IgcExcelCoreModule } from 'igniteui-webcomponents-excel';
import { IgcExcelModule } from 'igniteui-webcomponents-excel';

IgcExcelCoreModule.register();
IgcExcelModule.register();
IgcExcelXlsxModule.register();

export class ExcelUtility {
  
  public static getExtension(format: WorkbookFormat) {
    switch (format) {
      case WorkbookFormat.StrictOpenXml:
      case WorkbookFormat.Excel2007:
      return ".xlsx";
      case WorkbookFormat.Excel2007MacroEnabled:
      return ".xlsm";
      case WorkbookFormat.Excel2007MacroEnabledTemplate:
      return ".xltm";
      case WorkbookFormat.Excel2007Template:
      return ".xltx";
      case WorkbookFormat.Excel97To2003:
      return ".xls";
      case WorkbookFormat.Excel97To2003Template:
      return ".xlt";
    }
  }
  
  public static load(file: File): Promise<Workbook> {
    return new Promise<Workbook>((resolve, reject) => {
      ExcelUtility.readFileAsUint8Array(file).then((a) => {
        Workbook.load(a, new WorkbookLoadOptions(), (w) => {
          resolve(w);
        }, (e) => {
          reject(e);
        });
      }, (e) => {
        reject(e);
      });
    });
  }
  
  public static loadFromUrl(url: string): Promise<Workbook> {
    return new Promise<Workbook>((resolve, reject) => {
      const req = new XMLHttpRequest();
      req.open("GET", url, true);
      req.responseType = "arraybuffer";
      req.onload = (_d) => {
        const data = new Uint8Array(req.response);
        Workbook.load(data, new WorkbookLoadOptions(), (w) => {
          resolve(w);
        }, (e) => {
          reject(e);
        });
      };
      req.send();
    });
  }
  
  public static save(workbook: Workbook, fileNameWithoutExtension: string): Promise<string> {
    return new Promise<string>((resolve, reject) => {
      const opt = new WorkbookSaveOptions();
      opt.type = "blob";
      
      workbook.save(opt, (d) => {
        const fileExt = ExcelUtility.getExtension(workbook.currentFormat);
        const fileName = fileNameWithoutExtension + fileExt;
        saveAs(d as Blob, fileName);
        resolve(fileName);
      }, (e) => {
        reject(e);
      });
    });
  }
  
  private static readFileAsUint8Array(file: File): Promise<Uint8Array> {
    return new Promise<Uint8Array>((resolve, reject) => {
      const fr = new FileReader();
      fr.onerror = (_e) => {
        reject(fr.error);
      };
      
      fr.onload = (_e) => {
        resolve(new Uint8Array(fr.result as ArrayBuffer));
      };
      fr.readAsArrayBuffer(file);
    });
  }
}
