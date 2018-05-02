import { Component, ViewChild, ElementRef, OnInit } from '@angular/core';
import { DataService } from '../data.service';
import { ActivatedRoute } from '@angular/router';

declare var window;
declare var document;
@Component({
  selector: 'app-leonardo',
  templateUrl: './leonardo.component.html',
  styleUrls: ['./leonardo.component.scss']
})
export class LeonardoComponent implements OnInit {

  questionConfig:any;

  constructor(private dataService: DataService, private route: ActivatedRoute) {
    this.questionConfig = this.dataService.getQuestionConfig("v1");
  }
  ngOnInit() {
  }
}
