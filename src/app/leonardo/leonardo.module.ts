import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { RouterModule, Routes } from '@angular/router';
import { FormsModule } from '@angular/forms';
import { LeonardoComponent } from './leonardo.component';
import { SpreadsheetConfiguratorComponent } from './spreadsheetConfigurator/spreadsheetConfigurator.component';
import { WorkspaceComponent } from './workspace/workspace.component';
import { LeoAreaComponent } from './leo-area/leo-area.component';
import { LeoHeaderComponent } from './leo-header/leo-header.component';

export const leoRoutes: Routes = [
  {
    path: '',
    component: LeonardoComponent
  }
];

@NgModule({
  imports: [
    CommonModule,
    FormsModule,
    RouterModule
  ],
  exports: [
    LeonardoComponent
  ],
  declarations: [
    LeonardoComponent,
    SpreadsheetConfiguratorComponent,
    WorkspaceComponent,
    LeoAreaComponent,
    LeoHeaderComponent
  ]
})
export class LeonardoModule { }
