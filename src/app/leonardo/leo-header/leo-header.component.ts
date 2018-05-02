import { Component, OnInit, Input } from '@angular/core';

@Component({
  selector: 'app-leo-header',
  templateUrl: './leo-header.component.html',
  styleUrls: ['./leo-header.component.scss']
})
export class LeoHeaderComponent implements OnInit {
  @Input() questionConfig: any;
  assignmmentConfig:any;
  timeRemaining:any;
  constructor() { }

  ngOnInit() {
    this.assignmmentConfig = this.questionConfig.assignmmentConfig;
  }
}
