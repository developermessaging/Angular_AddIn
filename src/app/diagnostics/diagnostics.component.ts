import { Component } from '@angular/core';

const template = require('./diagnostics.component.html');

@Component({
    selector: 'diagnostics',
    template
})
export default class diagnostics {

    private _Diagnostics: Office.Diagnostics = Office.context.mailbox.diagnostics;
    public get Diagnostics(): Office.Diagnostics {
        return this._Diagnostics;
    }
    public set Diagnostics(value: Office.Diagnostics) {
        this._Diagnostics = value;

    }

    constructor() {
        this._Diagnostics = Office.context.mailbox.diagnostics;
        debugger;
        
    }


}