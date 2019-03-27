import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { FormsModule } from "@angular/forms";
import { CommonModule } from '@angular/common';
import { HttpClientModule } from '@angular/common/http';


//Angular material from: https://material.angular.io/guide/getting-started
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';
import { MatButtonModule, MatCheckboxModule, MatTabsModule } from '@angular/material';


import AppComponent from './app.component';
import attachments from './attachments/attachments.component';
import item from './item/item.component';
import message from './message/message.component';
import calendar from './calendar/calendar.component';
import userprofile from './userprofile/userprofile.component';
import diagnostics from './diagnostics/diagnostics.component';
import other from './other/other.component';
import context from './context/context.component';
import newcase from './newcase/newcase.component';
import excel from './excel/excel.component';


@NgModule({
  declarations: [AppComponent, attachments, item, message, calendar, userprofile, diagnostics, other, context, newcase, excel],
  imports: [BrowserModule, FormsModule, CommonModule, HttpClientModule,
              //MsalModule.forRoot( { clientID: "a1d2f69a-d480-42b4-b9f3-a9abc66de29f" } )

    //  ,MatButtonModule, MatCheckboxModule, MatTabsModule,
    //  BrowserAnimationsModule
  ],
  // exports: [
  //           MatButtonModule, MatCheckboxModule, MatTabsModule,
  //           BrowserAnimationsModule
  //          ],

  bootstrap: [AppComponent]
})
export default class AppModule {


}