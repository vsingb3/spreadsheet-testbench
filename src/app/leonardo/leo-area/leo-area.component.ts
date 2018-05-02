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


  configJSONchangedListener(configJSON) {
    this.workspace.showChanges(configJSON);
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
    this.workspace.refreshSpreadsheet();
  }

  containerDimChangedListener(containerDimJSON) {
    this.workspace.adjustDim(containerDimJSON);
  }

  verticalPaneResizeListener() {
    this.workspace.refreshSpreadsheet();
  }


  //   showChanges(){

  //     let latestConfig = jsonconfigurator.mergeJSON(config,spreadsheetConfigJSON);
  //     $("#spreadsheet").data("kendoSpreadsheet").destroy()
  //     $("#spreadsheet").empty();
  //     kendoInstance = new kendoWrapper (latestConfig).init();
  // return false;
  // }
}

