import './style.css';
import { IgcSpreadsheetModule, IgcSpreadsheetComponent } from 'igniteui-webcomponents-spreadsheet';
import { ModuleManager } from 'igniteui-webcomponents-core';
import { ExcelUtility } from './ExcelUtility';

ModuleManager.register(IgcSpreadsheetModule);

export class SpreadsheetOverview {
  private spreadsheet!: IgcSpreadsheetComponent;
  private fileInput!: HTMLInputElement;
  private urlInput!: HTMLInputElement;
  private openFromUrlBtn!: HTMLButtonElement;
  private saveBtn!: HTMLButtonElement;

  constructor() {
    this.render();
    this.initialiseElements();
    this.attachEventListeners();
    this.checkForUrlParam();
  }

  private initialiseElements(): void {
    this.spreadsheet = document.getElementById('spreadsheet') as IgcSpreadsheetComponent;
    this.fileInput = document.getElementById('fileInput') as HTMLInputElement;
    this.urlInput = document.getElementById('urlInput') as HTMLInputElement;
    this.openFromUrlBtn = document.getElementById('openFromUrlBtn') as HTMLButtonElement;
    this.saveBtn = document.getElementById('saveBtn') as HTMLButtonElement;
  }

  private attachEventListeners(): void {
    this.fileInput.addEventListener('change', this.openFile.bind(this));
    this.openFromUrlBtn.addEventListener('click', this.openFileFromUrl.bind(this));
    this.saveBtn.addEventListener('click', this.saveFile.bind(this));
  }

  private render(): void {
    const html = String.raw;
    const appHtml = html`
      <div class="container sample">
        <div class="options horizontal">
          <input type="file" accept=".xls, .xlt, .xlsx, .xlsm, .xltm, .xltx" id="fileInput" />
          
          <div class="options horizontal">
            <button id="openFromUrlBtn">Open From URL</button>
            <input type="text" id="urlInput" placeholder="Enter URL" />
          </div>

          <button id="saveBtn">Save</button>
        </div>

        <igc-spreadsheet id="spreadsheet" width="100%" height="calc(100% - 30px)"></igc-spreadsheet>
      </div>
      <div id="toast" class="toast"></div>
    `;

    document.querySelector<HTMLDivElement>('#app')!.innerHTML = appHtml;
  }

  private showToast(message: string): void {
    const toast = document.getElementById('toast')!;
    toast.textContent = message;
    toast.classList.add('show');

    setTimeout(() => {
      toast.classList.remove('show');
    }, 3000);
  }

  private isWorkbookEmpty(): boolean {
    const content = this.spreadsheet?.textContent?.trim();
    return content === 'WjWjSheet1' || content === 'WjSheet1Sheet1';
  }

  private checkForUrlParam(): void {
    const urlParams = new URLSearchParams(window.location.search);
    const url = urlParams.get('url');

    if (url) {
      this.urlInput.value = url;
      this.openFileFromUrl();
    }
  }

  public openFile(e: Event): void {
    const target = e.target as HTMLInputElement;
    const files = target.files;
    
    if (!files || files.length === 0) {
      return;
    }
  
    this.urlInput.value = '';
  
    if (!this.isWorkbookEmpty()) {
      const confirmed = window.confirm('You have unsaved changes. Are you sure you want to open another file?');
      if (!confirmed) {
        return;
      }
    }
  
    ExcelUtility.load(files[0]).then(
      (w) => {
        this.spreadsheet!.workbook = w;
      },
      (error) => {
        console.error(error);
        this.showToast('Failed to load file.');
      }
    );
  }
  
  public openFileFromUrl(): void {
    const url = this.urlInput.value;
    
    if (!url) {
      return;
    }
  
    if (!this.isWorkbookEmpty()) {
      const confirmed = window.confirm('You have unsaved changes. Are you sure you want to open another file?');
      if (!confirmed) {
        return;
      }
    }
  
    this.fileInput.value = '';
    this.urlInput.value = '';
  
    ExcelUtility.loadFromUrl(url).then(
      (w) => {
        this.spreadsheet!.workbook = w;
      },
      (error) => {
        console.error(error);
        this.showToast('Failed to load from URL.');
      }
    );
  }
  
  public saveFile(): void {
    ExcelUtility.save(this.spreadsheet!.workbook, 'form');
  }
}

new SpreadsheetOverview();
