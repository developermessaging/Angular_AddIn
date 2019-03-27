import { Component } from '@angular/core';
import { MatCheckbox } from '@angular/material';

const template = require('./attachments.component.html');

@Component({
    selector: 'attachments',
    template
})
export default class attachments {

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

}
