import { Component, OnInit } from '@angular/core';
import { MatCheckbox, MatCheckboxModule } from '@angular/material';
import { NewCaseModel } from './model/newcase.model';
import { HttpClient, HttpClientModule } from '@angular/common/http';
import { stringify } from 'querystring';
import { debug } from 'util';



const template = require('./newcase.component.html');

@Component({
    selector: 'newcase',
    template
})
export default class attachments implements OnInit {

    async ngOnInit() {
        // console.log( (Office.context as any).auth.GetAcessTokenAsync() );
    }

    //#region Properties
    private _CurrentItem: Office.Item = Office.context.mailbox.item;
    public get CurrentItem(): Office.Item {
        return this._CurrentItem;
    }
    public set CurrentItem(value: Office.Item) {
        this._CurrentItem = value;

    }

    Model: NewCaseModel = new NewCaseModel();
    //#endregion

    constructor(private http:HttpClient) {
        this.CurrentItem = Office.context.mailbox.item;
    }


    async ResetForm() {
        this.Model = new NewCaseModel();
    }

    async AnnounceOnTeams(http: HttpClient) {
        /**
        * Call Team API
        */
       var that=this;

       debugger;
        (Office.context as any).auth.getAccessTokenAsync(function (result) {
            var token: string = "InitialToken-Invalid";
            
            debugger;

            console.log("status=" + result.status);
            console.log("value= " + result.value);
            // console.log("error code= " + result.error.code);
            // console.log("message= " + result.error.message);

            token = result.value;

            http.get('https://graph.microsoft.com/v1.0/me/joinedTeams', 
            { headers: { 'Authorization': 'Bearer ' + token } }
            ).subscribe( data=> { console.log(data) }, 
                         error=> { console.log(error)}
                       );


        });



    }
}