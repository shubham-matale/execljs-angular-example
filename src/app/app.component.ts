import { Component } from '@angular/core';
import {ExcelServiceService} from "./excel-service.service";
import { DatePipe } from '@angular/common';
import { Workbook } from 'exceljs';
import * as fs from 'file-saver';
@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'excel-sheet-try';
  file:any;
  constructor(private excelService: ExcelServiceService,private datePipe: DatePipe) {
  }
  generateExcel() {
    this.excelService.generateExcel();
  }

  generateTeamplate(){
    let header = ['Plant',"Asset","Date","ABC","ADE","DEC"]
    let data = [];
    var now = new Date();
    for(let i=0;i<100;i++){
      now.setMinutes(now.getMinutes() + 5);
      now = new Date(now);
      data.push(["GHANI","INV_01",this.datePipe.transform(now, 'dd/MM/YYYY hh:mm:00'),"","",""])
    }
    this.excelService.generateExcelForTheTemplate(header,data);
  }

  // @ts-ignore
  handleChange(e) {
    this.file = e.target.files[0]
    this.validateFile()
  }


  validateFile(){
      const wb = new Workbook();
      const reader = new FileReader()
      let newWorkbook = new Workbook();
      let worksheet = newWorkbook.addWorksheet('GHANI');
      let header = ['Plant',"Asset","Date","ABC","ADE","DEC"]
      worksheet.addRow(header);
      let isErrors = false;
      let requiredColumn = 6;
      reader.readAsArrayBuffer(this.file)
      reader.onload = () => {
        const buffer = reader.result;
        // @ts-ignore
        wb.xlsx.load(buffer).then(workbook => {
          workbook.eachSheet((sheet, id) => {
            sheet.eachRow((row, rowIndex) => {
              if(rowIndex!=1){
                let rowData = row.values;
                // @ts-ignore
                if(rowData.length<requiredColumn){
                  isErrors = true;
                  let row = worksheet.addRow(rowData);
                  row.eachCell(cell => {
                    cell.fill = {
                      type: 'pattern',
                      pattern: 'solid',
                      fgColor: { argb: "FF99FF99" }
                    }
                  })

                }else {
                  let rowTempData = worksheet.addRow(rowData);
                  // @ts-ignore
                  for(let i=4;i<rowData.length;i++){

                    // @ts-ignore
                    // @ts-ignore
                    if(typeof rowData[i]=="string"){
                      // @ts-ignore
                      if(parseInt(rowData[i])){

                        // @ts-ignore
                        console.log("data")
                      } else{
                        isErrors =true;
                        console.log("Nan")
                        let cell = rowTempData.getCell(i);
                        console.log(cell)
                        cell.fill = {
                          type: 'pattern',
                          pattern: 'solid',
                          fgColor :{ argb: "FF9999" },
                          bgColor: { argb: "FF99FF99" }
                        }


                      }
                    }
                  }
                }
                console.log(row.values, rowIndex)
              }


            });
            console.log("errors",isErrors)
            if(isErrors){

              newWorkbook.xlsx.writeBuffer().then((data) => {
                let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                fs.saveAs(blob, 'error.xlsx');
              })
            }
          })
        })
      }
  }
}
