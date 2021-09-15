import { Component, VERSION } from '@angular/core';
import { ExcelService } from './service/excel.service';
import * as XLSX from 'xlsx';
import { IInvoice, Invoice } from './invoice.model';

@Component({
  selector: 'my-app',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  name = 'Angular ' + VERSION.major;
  willDownload = false;
  invoiceKeyList: string[] = [];
  invoices: any[] = [];

  constructor(private excelService: ExcelService) {
    const invoice = new Invoice();
    this.invoiceKeyList = Object.keys(invoice);
    console.log(this.invoiceKeyList);
  }

  onFileChange(ev) {
    let workBook = null;
    let jsonData = null;
    const reader = new FileReader();
    const file = ev.target.files[0];
    reader.onload = event => {
      const data = reader.result;
      workBook = XLSX.read(data, { type: 'binary' });
      jsonData = workBook.SheetNames.reduce((initial, name) => {
        const sheet = workBook.Sheets[name];
        initial[name] = XLSX.utils.sheet_to_json(sheet);
        return initial;
      }, {});
      const dataString = JSON.stringify(jsonData);
      document.getElementById('output').innerHTML = dataString
        .slice(0, 300)
        .concat('...');
      this.setDownload(dataString);

      const jsonArr = JSON.parse(dataString);
      this.invoices = [];
      if (workBook.SheetNames.length !== undefined) {
        for (let i = 0; i < workBook.SheetNames.length; i++) {
          jsonArr[workBook.SheetNames[i]].forEach(obj => {
            const invoiceObj = this.invoiceKeyList.reduce((carry, item) => {
              carry[item] = undefined;
              return carry;
            }, {});

            for (var key in obj) {
              this.invoiceKeyList.forEach(k => {
                // console.log("key: " + key + ", value: " + obj[key])
                // console.log("k: " + k + ", value: " + invoiceObj[k]);
                // console.log(key === k);
                if (key === k) {
                  invoiceObj[k] = obj[key];
                }
              });
              // console.log("key: " + key + ", value: " + obj[key])
            }
            console.log('invoiceObj: ' + invoiceObj);
            this.invoices.push(invoiceObj);
          });
          this.countLineNO();
          this.excelService.exportAsExcelFile(this.invoices, 'export-to-excel');
        }
      } else {
        jsonArr[workBook.SheetNames[0]].forEach(obj => {
          const invoiceObj = this.invoiceKeyList.reduce((carry, item) => {
            carry[item] = undefined;
            return carry;
          }, {});

          for (var key in obj) {
            this.invoiceKeyList.forEach(k => {
              // console.log("key: " + key + ", value: " + obj[key])
              // console.log("k: " + k + ", value: " + invoiceObj[k]);
              // console.log(key === k);
              if (key === k) {
                invoiceObj[k] = obj[key];
              }
            });
            // console.log("key: " + key + ", value: " + obj[key])
          }
          console.log('invoiceObj: ' + invoiceObj);
          this.invoices.push(invoiceObj);
        });
        this.countLineNO();
        this.excelService.exportAsExcelFile(this.invoices, 'export-to-excel');
      }
    };
    reader.readAsBinaryString(file);
  }

  countLineNO(): void {
    for (let i = 0; i < this.invoices.length; i++) {
      const item = this.invoices[i];
      let counting = 1;
      for (let j = i; i > 0; i--) {
        const compareObj = this.invoices[j];
        if (item.BILL_NO === compareObj.BILL_NO) {
          counting++;
        }
      }
      item.LINE_NO = String(counting);
    }
  }

  setDownload(data) {
    this.willDownload = true;
    setTimeout(() => {
      const el = document.querySelector('#download');
      el.setAttribute(
        'href',
        `data:text/json;charset=utf-8,${encodeURIComponent(data)}`
      );
      el.setAttribute('download', 'xlsxtojson.json');
    }, 1000);
  }
}
