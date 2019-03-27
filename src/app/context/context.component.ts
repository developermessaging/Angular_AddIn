import { Component } from '@angular/core';

const template = require('./context.component.html');

@Component({
    selector: 'context',
    template
})
export default class context {

    //#region "Class Properties"

    private _CurrentItem: Office.Item = Office.context.mailbox.item;
    public get CurrentItem(): Office.Item {
        return this._CurrentItem;
    }
    public set CurrentItem(value: Office.Item) {
        this._CurrentItem = value;

    }

    
    private _Mailbox : Office.Mailbox;
    public get Mailbox() : Office.Mailbox {
        return this._Mailbox;
    }
    public set Mailbox(v : Office.Mailbox) {
        this._Mailbox = v;
    }

    
    private _Context : Office.Context;
    public get Context() : Office.Context {
        return this._Context;
    }
    public set Context(v : Office.Context) {
        this._Context = v;
    }
    
    //#endregion

    constructor() {
        this._CurrentItem = Office.context.mailbox.item;
        this._Mailbox = Office.context.mailbox;
        this.Context = Office.context;
        debugger;
        //
        // this.Context.displayLanguage;
        // this.Context.license;
        // this.Context.officeTheme.bodyBackgroundColor
        // this.Context.officeTheme.bodyForegroundColor
        // this.Context.officeTheme.controlBackgroundColor
        // this.Context.officeTheme.controlForegroundColor
        // this.Context.platform;
        // //https://docs.microsoft.com/en-us/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets
        // var x = this.Context.requirements.isSetSupported( "a",1.3);
        // this.Context.touchEnabled
        


    }
} 