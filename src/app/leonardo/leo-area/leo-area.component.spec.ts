import { async, ComponentFixture, TestBed } from '@angular/core/testing';

import { LeoAssessmentPlayerComponent } from './leo-assessment-player.component';

describe('LeoAssessmentPlayerComponent', () => {
  let component: LeoAssessmentPlayerComponent;
  let fixture: ComponentFixture<LeoAssessmentPlayerComponent>;

  beforeEach(async(() => {
    TestBed.configureTestingModule({
      declarations: [ LeoAssessmentPlayerComponent ]
    })
    .compileComponents();
  }));

  beforeEach(() => {
    fixture = TestBed.createComponent(LeoAssessmentPlayerComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
