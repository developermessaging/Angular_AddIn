import { Component } from '@angular/core';
import { MatCheckbox } from '@angular/material';
import utils from '../utils.component';


const template = require('./item.component.html');

@Component({
    selector: 'item',
    template
})
export default class item {

    public Utils:utils = new utils();

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