import { Component, VERSION } from '@angular/core';
import { ExcelService } from './service/excel.service';
import * as XLSX from 'xlsx';
import { IInvoice, Invoice } from './invoice.model';
import { CdkDragDrop, moveItemInArray } from '@angular/cdk/drag-drop';
import { ErrorMsg, IErrorMsg } from './errorMsg.model';

@Component({
  selector: 'my-app',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  name = 'Cassie';
  willDownload = false;
  invoiceKeyList: string[] = [];
  invoices: any[] = [];
  errorMsg: IErrorMsg[] = [];
  isEdit = false;
  acceptExcelOnly = '.csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel';
  isExcelOnly = true;
  excelStyle = '  color: #141a46; background-color: #ec8b5e;';
  notExcelStyle = '  color: #ec8b5e; background-color: #141a46;';
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
      console.log('workBook.SheetNames.length' + workBook.SheetNames.length);
      if (workBook.SheetNames.length !== undefined) {
        for (let i = 0; i < workBook.SheetNames.length; i++) {
          this.invoices = [];
          jsonArr[workBook.SheetNames[i]].forEach(obj => {
            const invoiceObj = this.invoiceKeyList.reduce((carry, item) => {
              carry[item] = undefined;
              return carry;
            }, {});

            let isObjNotEmpty = false;
            for (var key in obj) {
              this.invoiceKeyList.forEach(k => {
                // console.log("key: " + key + ", value: " + obj[key])
                // console.log("k: " + k + ", value: " + invoiceObj[k]);
                // console.log(key === k);
                if (key === k) {
                  if (obj[key] !== undefined) {
                    invoiceObj[k] = obj[key];
                    isObjNotEmpty = true;
                  }
                }
              });
              // console.log("key: " + key + ", value: " + obj[key])
            }
            // console.log(isObjNotEmpty)
            // console.log(invoiceObj)
            if (isObjNotEmpty) {
              this.invoices.push(invoiceObj);
            }
          });
          this.countLineNO();
          if (this.invoices.length > 0) {
            this.excelService.exportAsExcelFile(
              this.invoices,
              'export-to-excel'
            );
          } else {
            const msgObj = new ErrorMsg();
            msgObj.msg = 'Sheet ' + i + ' does not match any field names that are shown in the botton of the list';
          }
        }
      } else {
        this.invoices = [];
        jsonArr[workBook.SheetNames[0]].forEach(obj => {
          const invoiceObj = this.invoiceKeyList.reduce((carry, item) => {
            carry[item] = undefined;
            return carry;
          }, {});

          let isObjNotEmpty = false;
          for (var key in obj) {
            this.invoiceKeyList.forEach(k => {
              // console.log("key: " + key + ", value: " + obj[key])
              // console.log("k: " + k + ", value: " + invoiceObj[k]);
              // console.log(key === k);
              if (key === k) {
                if (obj[key] !== undefined) {
                  invoiceObj[k] = obj[key];
                  isObjNotEmpty = true;
                }
              }
            });
            // console.log("key: " + key + ", value: " + obj[key])
          }
          // console.log(isObjNotEmpty)
          // console.log(invoiceObj)
          if (isObjNotEmpty) {
            this.invoices.push(invoiceObj);
          }
        });
        this.countLineNO();
        if (this.invoices.length > 0) {
          this.excelService.exportAsExcelFile(this.invoices, 'export-to-excel');
        } else {
          alert('None of the field matched');
        }
      }
    };
    reader.readAsBinaryString(file);
  }

  countLineNO(): void {
    console.log(this.invoices.length);
    for (let i = 0; i < this.invoices.length; i++) {
      const item = this.invoices[i];
      // console.log(item);
      // console.log(item.BILL_NO);
      let counting = 1;
      console.log('i = ' + i);
      for (let j = i-1; j >= 0; j--) {
        console.log('j = ' + j);
        const compareObj = this.invoices[j];
        // console.log(compareObj);
        console.log('counting before = ' + counting);
        if (item.BILL_NO === compareObj.BILL_NO) {
          console.log(item.BILL_NO + '===' + compareObj.BILL_NO);
          counting++;
          console.log('counting '+counting);
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

  drop(event: CdkDragDrop<string[]>) {
    moveItemInArray(this.invoiceKeyList, event.previousIndex, event.currentIndex);
  }

  editOrder(): void{
    this.isEdit = !this.isEdit;
  }

  changeAcceptedFile(): void{
    this.isExcelOnly = !this.isExcelOnly;
  }
}
