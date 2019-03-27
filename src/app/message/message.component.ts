import { Component } from '@angular/core';

const template = require('./message.component.html');

@Component({
    selector: 'message',
    template
})
export default class message {

    private _CurrentItem: Office.Item = Office.context.mailbox.item;
    public get CurrentItem(): Office.Item {
        return this._CurrentItem;
    }
    public set CurrentItem(value: Office.Item) {
        this._CurrentItem = value;

    }

    constructor() {
        this.CurrentItem = Office.context.mailbox.item;
        // this.CurrentItem.from.getAsync(this.callback);
        
    }

    

    callback(asyncResult) {
        var from = asyncResult.value;
        console.log("From " + from);
    }

    // Format an EmailAddressDetails object as
    // GivenName Surname <emailaddress>
    buildEmailAddressString(address) {
        return address.displayName + " &lt;" + address.emailAddress + "&gt;";
    }
     // Take an array of EmailAddressDetails objects and
    // build a list of formatted strings, separated by a line-break
    buildEmailAddressesString(addresses) {
        if (addresses && addresses.length > 0) {
            var returnString = "";

            for (var i = 0; i < addresses.length; i++) {
                if (i > 0) {
                    returnString = returnString + "<br/>";
                }
                returnString = returnString + this.buildEmailAddressString(addresses[i]);
            }

            return returnString;
        }

        return "None";
    }
}