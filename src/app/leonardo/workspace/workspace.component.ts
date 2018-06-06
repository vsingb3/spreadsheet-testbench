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
  @Output() activeSheetChanged: EventEmitter<any> = new EventEmitter();
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
    let configparam = args[1].showGridLines;

    Object.keys(spreadsheets).some( function (value){
      if(activesheetName == spreadsheets[value]["name"]){
        newjson.grid.sheets[value].showGridLines = configparam;
        return true;
      }
    });

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
      events:{ activeSheetChanged:this.activeSheetChangedHandler.bind(this) },
      uiStyle: {
        horizontalAlignment: 'center'
      }});
    this.spreadsheet.init();
  }

  activeSheetChangedHandler(sheetName){
    this.activeSheetChanged.emit(this.spreadsheet.getState());
  }


  showChanges(newConfiguration) {
    let currentState = this.spreadsheet.getState();
    this.currentConfig = this.mergeJSON(newConfiguration);

    //If ribbon is not changed, call set state
    if(currentState.ribbon.visible == this.currentConfig.ribbon.visible){
      this.spreadsheet.setState(this.currentConfig);
    }
    else{
      this.spreadsheet.destroy();
      this.spreadsheet = null;
      this.intialiseSpreadsheet(this.currentConfig);
    }
  }

  mergeJSON(changesconfig) {
    let newJSON = this.spreadsheet.getState();
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

  getData(){
    return this.spreadsheet.getData();
  }

  setData(extractorOutput){
    let currentState = this.spreadsheet.getState();
    this.currentConfig = this.spreadsheet.getState();
    this.currentConfig.grid = extractorOutput;

    try {
      this.spreadsheet.destroy();
      this.spreadsheet = null;
      this.intialiseSpreadsheet(this.currentConfig);
    } catch (exception) {

      if(this.spreadsheet){
        while (this.container.firstChild) {
          this.container.removeChild(this.container.firstChild);
        }
        this.spreadsheet = null;
        this.intialiseSpreadsheet(currentState);
      }

      alert("Provided input is not valid. Please check and provide a valid input");
    }
  }

  changeJson(type){
    this.currentConfig = this.kendoConfig["json"][type];
    this.spreadsheet.destroy();
    this.spreadsheet = null;
    this.intialiseSpreadsheet(this.currentConfig);
  }
}
