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

    
    private _body : string;
    public get body() : string {
        return this._body;
    }
    public set body(v : string) {
        this._body = v;
    }

    
    private _requiredAttendees : string;
    public get requiredAttendees() : string {
        return this._requiredAttendees;
    }
    public set requiredAttendees(v : string) {
        this._requiredAttendees = v;
    }
    
    
    private _organizer : string;
    public get organizer() : string {
        return this._organizer;
    }
    public set organizer(v : string) {
        this._organizer = v;
    }
    
    
    private _start : string;
    public get start() : string {
        return this._start;
    }
    public set start(v : string) {
        this._start = v;
    }
    
    private _end : string;
    public get end() : string {
        return this._end;
    }
    public set end(v : string) {
        this._end = v;
    }
    
    
    async refreshCalendarData()
    {
        debugger;
        this.body = "";
        this.requiredAttendees = "";
        this.organizer = "";
        this.start="";
        this.end = "";
        
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, x=> { this.body = x.value; });
        Office.context.mailbox.item.requiredAttendees.getAsync(x=>{x.value.forEach(y=>{this.requiredAttendees += y.displayName + " - " + y.emailAddress +";"})} )
        Office.context.mailbox.item.organizer.getAsync(x=>{ this.organizer = x.value.emailAddress;})
        this.start = Office.context.mailbox.item.start.toISOString();
        this.end = Office.context.mailbox.item.end.toISOString();
        
    }
    constructor() {
        this.CurrentItem = Office.context.mailbox.item;
        this.refreshCalendarData();
    }

    splitArray(data:any):string
    {
        debugger;
        var retval:string="";
        if (data==null) {
            return "";
        }
        try {
            data.array.forEach(element => {
                retval += ( element + ";" );
            });

        } catch (error) {
            retval = retval + "";    
        }

        
        return retval;
    }

}