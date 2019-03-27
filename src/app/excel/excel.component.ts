import { Component } from '@angular/core';
import { MatCheckbox } from '@angular/material';
import utils from '../utils.component';
import * as XLSX from 'xlsx';
import { HttpClient } from '@angular/common/http';


const template = require('./excel.component.html');

@Component({
    selector: 'excel',
    template
})
export default class excel {

    public Utils:utils = new utils();

    private _CurrentItem: Office.Item = Office.context.mailbox.item;
    public get CurrentItem(): Office.Item {
        return this._CurrentItem;
    }
    public set CurrentItem(value: Office.Item) {
        this._CurrentItem = value;
    }
    public currentsheet:XLSX.SheetProps;

    constructor(private httpClient: HttpClient) {
        this.CurrentItem = Office.context.mailbox.item;
        var file:XLSX.WorkBook ;
        var url:string = "https://microsoft-my.sharepoint.com/:x:/r/personal/francese_microsoft_com/_layouts/15/doc2.aspx?sourcedoc=%7Bb4a58081-f066-46b4-b1e9-50d0a0a76c30%7D&action=default&uid=%7BB4A58081-F066-46B4-B1E9-50D0A0A76C30%7D&ListItemId=88356&ListId=%7B2CA34F11-0AF2-40B6-B4C0-CA1D5AB5D17E%7D&odsp=1&env=prodbubble&cid=a0291295-ebbf-4ede-886f-35f466d58c10"
        const bstr: string = "";





        this.httpClient.get(url).subscribe((res)=>{
            const wb: XLSX.WorkBook = XLSX.read(bstr, {type: 'binary'});
            console.log(res);
            this.currentsheet = wb.Workbook.Sheets[0];
        });
        
        
        
        
    }

}