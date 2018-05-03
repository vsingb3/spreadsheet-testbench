import { Component, OnInit, AfterViewInit, Input, Output, EventEmitter, ElementRef, ViewEncapsulation } from '@angular/core';
import { NgForm } from '@angular/forms';
declare var Leonardo: any;
@Component({
    selector: 'app-spreadsheetConfigurator',
    templateUrl: './spreadsheetConfigurator.component.html',
    styleUrls: ['./spreadsheetConfigurator.component.scss'],
    encapsulation: ViewEncapsulation.None
})
export class SpreadsheetConfiguratorComponent implements OnInit, AfterViewInit {
    @Input() configuratorJSON: any;
    @Input() kendoConfig: any;
    @Output() configChanged: EventEmitter<any> = new EventEmitter();
    @Output() changeJsonEvent: EventEmitter<any> = new EventEmitter();
    @Output() getDataEvent: EventEmitter<any> = new EventEmitter();
    @Output() setDataEvent: EventEmitter<any> = new EventEmitter();
    inpConfigObject;
    constructor() {
        this.inpConfigObject = {};
    }

    ngOnInit() {
        this.setDefaultConfigurations("json1");
    };

    ngAfterViewInit() {
        for (let section of this.configuratorJSON) {
            for (let elem of section.configElem) {
                let inpElement = elem;
                if (inpElement["type"] == "checkbox") {
                    document.querySelector("#" + inpElement["id"] + "+i").addEventListener("click", function () {
                        let newValue = !this.inpConfigObject[section.name][inpElement["id"]];
                        this.inpConfigObject[section.name][inpElement["id"]] = newValue;
                        this.configChanged.emit(this.inpConfigObject["kendoConfig"]);
                    }.bind(this));
                }
            }
        }
    }

    setDefaultConfigurations(jsonType){

        let initialJson = this.kendoConfig["json"][jsonType];
        let initialConfig = this.setInitialConfiguration(initialJson);

        for (let section of this.configuratorJSON) {
            this.inpConfigObject[section["name"]] = {};
            for (let elem of section["configElem"]) {
                let inpElement = elem;
                if (inpElement["type"] == "checkbox") {
                    this.inpConfigObject[section["name"]][inpElement["id"]] = initialConfig[inpElement["id"]];
                }
                else if (inpElement["type"] == "dropdown") {
                    this.inpConfigObject[section["name"]][inpElement["id"]] = jsonType;
                }
                else {
                    this.inpConfigObject[section["name"]][inpElement["id"]] = "";
                }
            }
        }
    }

    setInitialConfiguration(inputJson){

        let defaultValues = {};
        defaultValues["formulabar"] = inputJson["formulabar"]["visible"];
        defaultValues["sheetbar"] = inputJson["sheetbar"]["visible"];
        defaultValues["ribbon"] = inputJson["ribbon"]["visible"];
        defaultValues["rowHeader"] = inputJson["grid"]["rowHeader"];
        defaultValues["colHeader"] = inputJson["grid"]["colHeader"];
        defaultValues["showGridLines"] = this.getDefaultGridLineValue(inputJson);
        return defaultValues;
    }

    getDefaultGridLineValue(inputJson){
        let spreadsheets = inputJson["grid"]["sheets"];
        let activesheetName = inputJson["grid"]["activeSheet"];
        let activesheetIndex = 0;
        for(let sheet =0;sheet<spreadsheets.length;sheet++){
          if(activesheetName == spreadsheets[sheet]["name"]){
            activesheetIndex = sheet;
            break;
          }
        }
        return inputJson.grid.sheets[activesheetIndex].showGridLines;
    }

    setExtractorOutput() {
        try {
          let extractorOutput = JSON.parse(this.inpConfigObject["initialConfig"]["extractorOutput"]);
          this.setDataEvent.emit(extractorOutput);
        } catch (exception) {
          alert("Provided Input is not a valid JSON");
        }
    }

    updateData(){
        this.getDataEvent.emit(true);
    }

    updateTextArea(data){
        let stringData = JSON.stringify(data, null, 2);
        this.inpConfigObject["initialConfig"]["extractorOutput"] = stringData;
    }
    changeJson(event){
        let dataType = event.target.value;
        this.changeJsonEvent.emit(dataType);
    }
}