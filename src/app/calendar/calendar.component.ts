import { Component } from '@angular/core';

const template = require('./calendar.component.html');

@Component({
    selector: 'calendar',
    template
})
export default class calendar {

    private _CurrentItem: Office.Item = Office.context.mailbox.item;
    public get CurrentItem(): Office.Item {
        return this._CurrentItem;
    }
    public set CurrentItem(value: Office.Item) {
        this._CurrentItem = value;

    }

    constructor() {
        this.CurrentItem = Office.context.mailbox.item;
        
    }

    splitArray(data:any):string
    {
        var retval:string="";

        data.array.forEach(element => {
            retval += ( element + ";" );
        });


        return retval;
    }

}