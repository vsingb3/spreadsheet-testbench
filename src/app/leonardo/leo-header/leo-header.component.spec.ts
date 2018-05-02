import { async, ComponentFixture, TestBed } from '@angular/core/testing';

import { LeoHeaderComponent } from './leo-header.component';

describe('LeoHeaderComponent', () => {
  let component: LeoHeaderComponent;
  let fixture: ComponentFixture<LeoHeaderComponent>;

  beforeEach(async(() => {
    TestBed.configureTestingModule({
      declarations: [ LeoHeaderComponent ]
    })
    .compileComponents();
  }));

  beforeEach(() => {
    fixture = TestBed.createComponent(LeoHeaderComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
