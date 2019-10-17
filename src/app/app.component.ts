import { Component, OnInit } from '@angular/core';
import { FormsModule } from '@angular/Forms';
import { OperationsModel } from './Models/Operations.model';
import { CommonModule } from '@angular/common';


const template = require('./app.component.html');

@Component({
    selector: 'app-home',
    template
})
export default class AppComponent implements OnInit {
    welcomeMessage = 'Welcome';


    public AddinTitle: string = "DEMO Addin"
    public AddinSubTitle: string = "Developer Support for Messaging "

    //public CurrentItem: Office.Item = Office.context.mailbox.item;
   
    Operations: OperationsModel[] = [
        { id: 1, name: "View Item Properties" },
        { id: 2, name: "View Attachments " },
        { id: 3, name: "View Message properties" },
        { id: 4, name: "View Calendar properties" },
        { id: 5, name: "View User Profile properties" },
        { id: 6, name: "View Diagnostics properties" },
        { id: 7, name: "Other Operations" },
        { id: 8, name: "View Context" },
        { id: 9, name: "New Case Announce on Teams" },
        { id: 10, name: "Excel ?" },
        { id: 11, name: "EWS" }



    ]
    SelectedOperationIndex: number;

    async run() {
        /**
        * Insert your Outlook code here
        */
        //this.Terminal = "Terminal de testes";    


        //this.CurrentItem = Office.context.mailbox.item;
        
    }

    ngOnInit() {
        this.SelectedOperationIndex = 1;
    }

    public onChange(event): void {  // event will give you full breif of action
        const newVal = event.target.value;
        
        console.log(newVal);
    }
}