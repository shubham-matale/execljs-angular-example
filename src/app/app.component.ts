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
  today = new Date();
  endTime = new Date(new Date().setHours(this.today.getHours()+5))
   inputData = {
    "planName":"Ghani",
    "startTime" : this.today,
    "assetName": "INV_01",
    "endTime": this.endTime,
     "tags" : ["tag1","tag2","tag3"],
     "granularity": 10
  }
  constructor(private excelService: ExcelServiceService,private datePipe: DatePipe) {
  }
  generateExcel() {
    this.excelService.generateExcel();
  }

  generateTeamplate(){


    let header = ['Plant',"Asset","Date"]
    for(let i of this.inputData["tags"]){
      header.push(i)
      header.push("Q_"+i);
    }
    console.log(this.inputData)

    // @ts-ignore
    let data:any = [];
    var now = this.inputData["startTime"];
    // @ts-ignore
    let iterations = (Math.round( (this.inputData["endTime"].getTime() - this.inputData["startTime"].getTime())/60000))/this.inputData["granularity"];


    for(let i=0;i<iterations;i++){
      now.setMinutes(now.getMinutes() + this.inputData["granularity"]);
      now = new Date(now);
      let tempArr = ["GHANI","INV_01",this.datePipe.transform(now, 'dd/MM/YYYY hh:mm:00')]
      for(let i=0;i<iterations;i++){
        tempArr.push("")
        tempArr.push("")
      }
      data.push(tempArr)
    }
    this.excelService.generateExcelForTheTemplate(header,data);
  }

  // @ts-ignore
  handleChange(e) {
    this.file = e.target.files[0]
    this.validateFile()
  }


  async validateFile() {
    const wb = new Workbook();
    const reader = new FileReader()
    let newWorkbook = new Workbook();
    let worksheet = newWorkbook.addWorksheet('GHANI');
    let header = ['Plant', "Asset", "Date", "ABC", "ADE", "DEC"]
    worksheet.addRow(header);
    let isErrors = false;
    let requiredColumn = 6;
    // @ts-ignore
    await reader.readAsArrayBuffer(this.file)
    reader.onload = () => {
      const buffer = reader.result;
      // @ts-ignore
      wb.xlsx.load(buffer).then(workbook => {
        workbook.eachSheet((sheet, id) => {
          sheet.eachRow((row, rowIndex) => {
            if (rowIndex != 1) {
              let rowData = row.values;
              // @ts-ignore
              if (rowData.length < requiredColumn) {
                isErrors = true;
                let row = worksheet.addRow(rowData);
                row.eachCell(cell => {
                  cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: {argb: "FF99FF99"}
                  }
                })

              } else {
                let rowTempData = worksheet.addRow(rowData);
                // @ts-ignore
                for (let i = 1; i < rowData.length; i++) {

                  // @ts-ignore
                  if(i<=3){

                  }else{
                    // @ts-ignore
                    if (rowData[i]==""){
                      isErrors =true;
                      let cell = rowTempData.getCell(i);
                      console.log(cell)
                      cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: {argb: "FF9999"},
                        bgColor: {argb: "FF99FF99"}
                      }
                      continue;
                    }
                    // @ts-ignore
                    if (typeof rowData[i] == "string" ) {
                      const reguarExp = new RegExp(/^([0-9]|[a-z])+([0-9a-z]+)$/i);
                      // @ts-ignore
                      if(rowData[i].match(reguarExp)){
                        console.log("alphnumeric")
                      }else {
                        console.log("not alph")
                      }
                      // @ts-ignore
                      if (parseInt(rowData[i])) {

                        // @ts-ignore
                        console.log(rowData[i]);
                        // @ts-ignore
                        console.log();
                      } else {
                        isErrors = true;
                        console.log("Nan")
                        let cell = rowTempData.getCell(i);
                        console.log(cell)
                        cell.fill = {
                          type: 'pattern',
                          pattern: 'solid',
                          fgColor: {argb: "FF9999"},
                          bgColor: {argb: "FF99FF99"}
                        }
                      }
                    }
                  }
                }
              }
              console.log(row.values, rowIndex)
            }


          });
          console.log("errors", isErrors)
          if (isErrors) {
            console.log("error function")
            newWorkbook.xlsx.writeBuffer().then((data) => {
              let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
              fs.saveAs(blob, 'error.xlsx');
            })
          }else{
            console.log("data is good to go")
          }
        })
      })
    }
    console.log("shubham")
  }


}
