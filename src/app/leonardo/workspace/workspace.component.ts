import { Component, OnInit, Input, ViewChild, AfterViewInit, ViewEncapsulation, ElementRef, Output, EventEmitter } from '@angular/core';

declare var LeonardoSpreadsheet: any;

@Component({
  selector: 'app-workspace',
  encapsulation:ViewEncapsulation.None,
  templateUrl: './workspace.component.html',
  styleUrls: ['./workspace.component.scss'],
  
})
export class WorkspaceComponent implements OnInit, AfterViewInit {
  @Input() kendoConfig: any;
  @ViewChild('leoHost') leoHost: ElementRef;
  @Output() gridEvent: EventEmitter<Object> = new EventEmitter();
  spreadsheet;
  configMap:any;
  container:any;
  currentConfig;
  constructor() {
    this.configMap = {
      rowHeader: this.showRowHeaders,
      colHeader: this.showColHeaders,
      showGridLines: this.showGridlines,
      ribbon: this.setRibbon,
      sheetbar: this.setSheetBar,
      formulabar: this.setFormulaBar,
      activeSheetIndex: "activesheetindex",
      nCols: "columns",
      nRows: "rows",
    }
  }

  ngOnInit() {
  }

  ngAfterViewInit(){
    this.container = document.querySelector("#leoHost");
    this.intialiseSpreadsheet();
  }

  showRowHeaders(args){
    let newjson = args[0];
    let configparam = args[1].rowHeader;

    return this.showHeaders("rowHeader", newjson, configparam)
  }

  showColHeaders(args){
    let newjson = args[0];
    let configparam = args[1].colHeader;

    return this.showHeaders("colHeader", newjson, configparam)
  }
  showHeaders(headerName, json, value) {
    json["grid"][headerName] = value;
    return json;
  }

  showGridlines(args) {
    let newjson = args[0];
    let spreadsheets = newjson["grid"]["sheets"];
    let activesheetName = newjson["grid"]["activeSheet"];
    let activesheetIndex = 0;
    for(let sheet =0;sheet<spreadsheets.length;sheet++){
      if(activesheetName == spreadsheets[sheet]["name"]){
        activesheetIndex = sheet;
        break;
      }
    }
    let configparam = args[1].showGridLines;
    newjson.grid.sheets[activesheetIndex].showGridLines = configparam;
    return newjson;
  }

  showActiveSheet(args) {
    let newjson = args[0];
    let configparam = args[1].activeSheetIndex
    let sheets = newjson["sheets"];
    if (configparam <= sheets.length && configparam > 0) {
      newjson["activeSheet"] = sheets[configparam - 1].name;
    }
    return newjson;
  }

  intialiseSpreadsheet(inputJson?) {

    this.currentConfig = inputJson || this.kendoConfig["json"]["json1"];
    this.spreadsheet = new LeonardoSpreadsheet("WB1",this.container,{config:this.currentConfig,
      events:{},
      uiStyle: {
        horizontalAlignment: 'center'
      }});
    this.spreadsheet.init()
  }


  showChanges(newConfiguration) {
    let newJSON = this.spreadsheet.getState();

    if(newConfiguration["extractorOutputChanged"]){
      let data = newConfiguration["extractorOutput"];
      this.spreadsheet.setData(data);
      return;
    }
    else if(newConfiguration["useCaseChanged"]){
      this.currentConfig = this.kendoConfig["json"][newConfiguration["useCase"]];
    }
    else{
      this.currentConfig = this.mergeJSON(newJSON, newConfiguration);
    }

    this.spreadsheet.destroy();
    this.intialiseSpreadsheet(this.currentConfig);
  }

  mergeJSON(spreadsheetconfig, changesconfig) {
    let newJSON = spreadsheetconfig;
    for (let key in changesconfig) {
      //for checkbox value
      if (typeof changesconfig[key] == "string" && changesconfig[key] == "")
        continue
      else {
        let mappedValue = this.configMap[key];
        if (typeof mappedValue == 'string') {
          newJSON[this.configMap[key]] = changesconfig[key];
        }
        else {
          newJSON = mappedValue.call(this, [newJSON, changesconfig]);
        }
      }
    }
    return newJSON;
  }

  setFormulaBar(args){
    let newjson = args[0];
    let visibility = args[1].formulabar;
    newjson["formulabar"]["visible"] = visibility;
    return newjson;
  }

  setSheetBar(args){
    let newjson = args[0];
    let visibility = args[1].sheetbar;
    newjson["sheetbar"]["visible"] = visibility;
    return newjson;
  }

  setRibbon(args){
    let newjson = args[0];
    let visibility = args[1].ribbon;
    newjson["ribbon"]["visible"] = visibility;
    return newjson;
  }

  adjustDim(dimJSON) {
    if (dimJSON.contWidth!="") {
      this.container.style.width = dimJSON.contWidth;
    }
    else{
      this.container.style.width = "100%";
    }
    if (dimJSON.contHeight) {
      this.container.style.height = dimJSON.contHeight;
    }
    else{
      this.container.style.height = "100%";
    }
   this.refreshSpreadsheet();
  }

  refreshSpreadsheet(){
    // this.spreadsheet.refresh();
  }

}
