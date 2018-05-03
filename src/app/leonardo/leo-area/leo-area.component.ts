import { Component, ViewChild, ElementRef, OnInit, Input } from '@angular/core';

declare var window;
declare var document;

@Component({
  selector: 'leo-area',
  templateUrl: './leo-area.component.html',
  styleUrls: ['./leo-area.component.scss']
})
export class LeoAreaComponent implements OnInit {

  @Input() id: string;
  @Input() questionConfig: any;
  @ViewChild('workspace') workspace;
  @ViewChild('spreadsheetConfigurator') spreadsheetConfigurator;
  configuratorJSON: any;
  kendoConfig: any;
  configuratorVisible: boolean;
  toogleBtnText:String;
  constructor() {
    this.configuratorVisible = true;
    this.toogleBtnText = "<<";
  }
  ngOnInit() {
    this.configuratorJSON = this.questionConfig["configuratorJSON"];
    this.kendoConfig = this.questionConfig["kendoConfig"];
  }

  configChangedListener(configJSON) {
    this.workspace.showChanges(configJSON);
  }

  changeJsonListener(type){
    this.workspace.changeJson(type);
  }

  getDataEventListener(){
    let data=  this.workspace.getData();
    this.spreadsheetConfigurator.updateTextArea(data);
  }

  setDataEventListener(data){
    this.workspace.setData(data);
  }

  toggleConfiguratorPane(configVisible) {
    this.configuratorVisible =configVisible;
    if (configVisible) {
      document.querySelector(".configuratorpane").classList.remove("hideConfigPane");
      document.querySelector(".spreadsheetpane").classList.remove("expandSpreadsheetPane");
      this.toogleBtnText = "<<";
    }
    else{
      document.querySelector(".configuratorpane").classList.add("hideConfigPane");
      document.querySelector(".spreadsheetpane").classList.add("expandSpreadsheetPane");
      this.toogleBtnText = ">>";
    }
  }
}

